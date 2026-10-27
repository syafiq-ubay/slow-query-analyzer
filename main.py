"""Export MySQL and Nginx records from Elasticsearch to Excel."""

import argparse
import csv
import os
from datetime import date, datetime, time, timezone
from pathlib import Path
from typing import Any, Dict, Iterator, List
from zoneinfo import ZoneInfo

from correlator import correlate_all, normalize_query, query_operation
from elastic_client import (
    ELASTIC_PASS,
    ELASTIC_URL,
    ELASTIC_USER,
    MYSQL_DATA_VIEW_ID,
    NGINX_DATA_VIEW_ID,
    ElasticLogClient,
)
from exporter import (
    MERGED_FIELDS,
    MYSQL_FIELDS,
    NGINX_FIELDS,
    export_to_excel,
)


def _http_status_code(value: str) -> str:
    try:
        code = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("HTTP status code must be an integer.") from exc
    if code < 100 or code > 599:
        raise argparse.ArgumentTypeError("HTTP status code must be between 100 and 599.")
    return str(code)


def _add_time_range(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--start-time",
        default=os.getenv("START_TIME", "auto"),
        help="Start time (default: auto from input/slow-queries.log, else fallback).",
    )
    parser.add_argument("--end-time", default=os.getenv("END_TIME", "now"),
                        help="End time (default: now).")


def _add_connection_options(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--max-records",
        type=int,
        default=0,
        help="Maximum records from each source (default 0=unlimited).",
    )


def _add_nginx_options(
    parser: argparse.ArgumentParser,
    *,
    default_exclude_codes: List[str] | None = None,
) -> None:
    help_text = "Exclude an Nginx HTTP status code; may be repeated."
    if default_exclude_codes:
        help_text += f" Defaults: {', '.join(default_exclude_codes)}."
    parser.add_argument(
        "--exclude-code",
        action="append",
        default=default_exclude_codes or [],
        type=_http_status_code,
        metavar="HTTP_CODE",
        help=help_text,
    )


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export MySQL slow-query and Nginx records from Elasticsearch."
    )
    configured_exclude_codes = [
        code.strip()
        for code in os.getenv("NGINX_EXCLUDE_CODES", "").split(",")
        if code.strip()
    ]
    try:
        configured_exclude_codes = [
            _http_status_code(code) for code in configured_exclude_codes
        ]
    except argparse.ArgumentTypeError as exc:
        parser.error(f"Invalid NGINX_EXCLUDE_CODES: {exc}")

    modes = parser.add_subparsers(dest="mode", required=True)

    mysql_parser = modes.add_parser("mysql", help="Export MySQL slow queries from Elastic.")
    _add_time_range(mysql_parser)
    _add_connection_options(mysql_parser)
    mysql_parser.add_argument("--output", default="output/mysql_slow_queries.xlsx")

    nginx_parser = modes.add_parser("nginx", help="Export Nginx access logs from Elastic.")
    _add_time_range(nginx_parser)
    _add_connection_options(nginx_parser)
    _add_nginx_options(nginx_parser, default_exclude_codes=configured_exclude_codes)
    nginx_parser.add_argument("--output", default="output/nginx_access_logs.xlsx")

    merge_parser = modes.add_parser(
        "merge", help="Fetch both Elastic sources and correlate records by time."
    )
    _add_time_range(merge_parser)
    _add_connection_options(merge_parser)
    _add_nginx_options(merge_parser, default_exclude_codes=configured_exclude_codes)
    merge_parser.add_argument("--output", default="output/merge_logs.xlsx")

    args = parser.parse_args()
    if args.max_records < 0:
        parser.error("--max-records cannot be negative; use 0 for unlimited.")
    args.mysql_query_format = os.getenv("MYSQL_QUERY_FORMAT", "short").strip().lower()
    if args.mysql_query_format not in {"full", "short", "super_short"}:
        parser.error("MYSQL_QUERY_FORMAT must be 'full', 'short', or 'super_short'.")
    if not ELASTIC_URL:
        parser.error("Set ELASTIC_URL in .env before running the exporter.")
    if args.mode in ("mysql", "merge") and not MYSQL_DATA_VIEW_ID:
        parser.error("Set MYSQL_DATA_VIEW_ID in .env before running this mode.")
    if args.mode in ("nginx", "merge") and not NGINX_DATA_VIEW_ID:
        parser.error("Set NGINX_DATA_VIEW_ID in .env before running this mode.")
    return args


def _elastic_client(args: argparse.Namespace) -> ElasticLogClient:
    return ElasticLogClient(
        base_url=ELASTIC_URL,
        username=ELASTIC_USER,
        password=ELASTIC_PASS,
    )


def _stream_records(
    client: ElasticLogClient,
    *,
    data_view_id: str,
    start_time: str,
    end_time: str,
    fields: List[str],
    max_records: int,
    exclude_codes: List[str] | None = None,
) -> Iterator[Dict[str, Any]]:
    source_fields = fields
    if fields is MYSQL_FIELDS:
        # Some data views provide a pre-normalized SQL shape. Fetch it when
        # available, while keeping the public MySQL export columns unchanged.
        source_fields = [*MYSQL_FIELDS, "query_normalized"]
    for page in client.stream_records(
        data_view_id=data_view_id,
        start_time=start_time,
        end_time=end_time,
        fields=source_fields,
        exclude_codes=exclude_codes,
        query_filter=os.getenv(
            "MYSQL_FILTER" if fields is MYSQL_FIELDS else "NGINX_FILTER", ""
        ).strip() or None,
        max_records=max_records,
    ):
        yield from page


def _mysql_query_records(
    records: Iterator[Dict[str, Any]], query_format: str
) -> Iterator[Dict[str, Any]]:
    """Choose the full SQL or a compact, value-normalized SQL shape."""
    for record in records:
        prepared = dict(record)
        raw_query = prepared.get("query", "")
        normalized_query = prepared.get("query_normalized", "")
        source_query = normalized_query or raw_query
        if query_format == "short":
            prepared["query"] = normalize_query(source_query)
        elif query_format == "super_short":
            prepared["query"] = query_operation(source_query)
        else:
            prepared["query"] = raw_query or normalized_query
        yield prepared


def main() -> None:
    args = _parse_args()
    if args.start_time == "auto":
        args.start_time = _input_start_time()
    if args.mode in ("mysql", "merge"):
        print(f"Format query MySQL: {args.mysql_query_format}")
    client = _elastic_client(args)

    if args.mode == "mysql":
        records = _mysql_query_records(
            _stream_records(
                client,
                data_view_id=MYSQL_DATA_VIEW_ID,
                start_time=args.start_time,
                end_time=args.end_time,
                fields=MYSQL_FIELDS,
                max_records=args.max_records,
            ),
            args.mysql_query_format,
        )
        total = export_to_excel(
            records, args.output, fields=MYSQL_FIELDS, sheet_title="MySQL Slow Queries"
        )
    elif args.mode == "nginx":
        records = _stream_records(
            client,
            data_view_id=NGINX_DATA_VIEW_ID,
            start_time=args.start_time,
            end_time=args.end_time,
            fields=NGINX_FIELDS,
            max_records=args.max_records,
            exclude_codes=args.exclude_code,
        )
        total = export_to_excel(
            records, args.output, fields=NGINX_FIELDS, sheet_title="Nginx Access Logs"
        )
    else:
        mysql_records = _mysql_query_records(
            _stream_records(
                client,
                data_view_id=MYSQL_DATA_VIEW_ID,
                start_time=args.start_time,
                end_time=args.end_time,
                fields=MYSQL_FIELDS,
                max_records=args.max_records,
            ),
            args.mysql_query_format,
        )
        nginx_records = _stream_records(
            client,
            data_view_id=NGINX_DATA_VIEW_ID,
            start_time=args.start_time,
            end_time=args.end_time,
            fields=NGINX_FIELDS,
            max_records=args.max_records,
            exclude_codes=args.exclude_code,
        )
        total = export_to_excel(
            correlate_all(mysql_records, nginx_records),
            args.output,
            fields=MERGED_FIELDS,
            sheet_title="Combined Logs",
        )

    print(f"Export selesai: {total:,} baris -> {args.output}")


def _input_start_time() -> str:
    """Use the earliest timestamp in input CSV/XLSX files, or today's start."""
    input_dir = Path(os.getenv("INPUT_DIR", "input"))
    timestamps = []
    input_files = sorted((*input_dir.glob("*.csv"), *input_dir.glob("*.xlsx")))
    for path in input_files:
        try:
            timestamps.extend(_timestamps_from_file(path))
        except (OSError, ValueError, ImportError) as exc:
            print(f"Peringatan: tanggal dari {path} tidak dapat dibaca: {exc}")

    if timestamps:
        earliest = min(timestamps).astimezone(timezone.utc)
        print(f"Tanggal awal otomatis dari file input: {earliest.isoformat()}")
        return earliest.isoformat().replace("+00:00", "Z")

    # Timestamp tanpa zona waktu dianggap waktu lokal Indonesia.
    jakarta = ZoneInfo("Asia/Jakarta")
    local_today = datetime.combine(date.today(), time.min, tzinfo=jakarta)
    if input_files:
        print(
            f"Tidak ditemukan timestamp yang didukung di file dalam {input_dir}; "
            "memakai awal hari ini."
        )
    else:
        print(f"Tidak ada file CSV/Excel di {input_dir}; memakai awal hari ini.")
    return local_today.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")


def _timestamps_from_file(path: Path) -> List[datetime]:
    """Read recognized timestamp columns from CSV or Excel input files."""
    records = []
    if path.suffix.lower() == ".csv":
        with path.open(newline="", encoding="utf-8-sig", errors="replace") as source:
            records.extend(csv.DictReader(source))
    else:
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            for worksheet in workbook.worksheets:
                rows = worksheet.iter_rows(values_only=True)
                headers = next(rows, None)
                if not headers:
                    continue
                names = [str(value).strip() if value is not None else "" for value in headers]
                records.extend(dict(zip(names, row)) for row in rows)
        finally:
            workbook.close()

    timestamp_names = {
        "@timestamp",
        "mysql_@timestamp",
        "http_timestamp",
        "timestamp",
        "date",
        "datetime",
    }
    result = []
    for record in records:
        for name, value in record.items():
            if name and str(name).strip().lower() in timestamp_names:
                parsed = _parse_input_datetime(value)
                if parsed is not None:
                    result.append(parsed)
    return result


def _parse_input_datetime(value: Any) -> datetime | None:
    if isinstance(value, datetime):
        parsed = value
    elif value is None or not str(value).strip():
        return None
    else:
        text = str(value).strip()
        try:
            parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
        except ValueError:
            for fmt in (
                "%b %d, %Y @ %H:%M:%S.%f",
                "%b %d, %Y @ %H:%M:%S",
                "%Y-%m-%d %H:%M:%S",
                "%m/%d/%Y %H:%M:%S",
            ):
                try:
                    parsed = datetime.strptime(text, fmt)
                    break
                except ValueError:
                    continue
            else:
                return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=ZoneInfo("Asia/Jakarta"))
    return parsed


if __name__ == "__main__":
    main()

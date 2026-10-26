"""Time-based matching of MySQL slow queries and Nginx access records."""

from datetime import datetime, timezone
from collections import defaultdict
from typing import Any, Dict, Iterable, Iterator, List, Optional


def parse_timestamp(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if text.isdigit():
        return float(text)
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.timestamp()


def correlate_all(
    mysql_records: Iterable[Dict[str, Any]],
    nginx_records: Iterable[Dict[str, Any]],
) -> Iterator[Dict[str, Any]]:
    """Join records whose timestamps match to the displayed second."""
    by_timestamp: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for nginx in nginx_records:
        key = _timestamp_key(nginx.get("@timestamp"))
        if key is not None:
            by_timestamp[key].append(nginx)

    for mysql in mysql_records:
        key = _timestamp_key(mysql.get("@timestamp"))
        matches = by_timestamp.get(key, []) if key is not None else []
        if not matches:
            yield _merged_row(mysql, None)
            continue
        for nginx in matches:
            yield _merged_row(mysql, nginx)


def _timestamp_key(value: Any) -> Optional[str]:
    """Normalize a timestamp to the `MMM DD, YYYY @ HH:mm:ss` second."""
    timestamp = parse_timestamp(value)
    if timestamp is None:
        return None
    parsed = datetime.fromtimestamp(timestamp, tz=timezone.utc)
    return parsed.strftime("%b %d, %Y @ %H:%M:%S")


def _merged_row(
    mysql: Dict[str, Any],
    nginx: Optional[Dict[str, Any]],
) -> Dict[str, Any]:
    return {
        "mysql_@timestamp": mysql.get("@timestamp", ""),
        "mysql_query_time": mysql.get("query_time", ""),
        "mysql_schema": mysql.get("schema", ""),
        "mysql_user": mysql.get("user", ""),
        "mysql_host": mysql.get("host", ""),
        "mysql_query_type": mysql.get("query_type", ""),
        "mysql_rows_examined": mysql.get("rows_examined", ""),
        "mysql_rows_sent": mysql.get("rows_sent", ""),
        "mysql_query": mysql.get("query_normalized") or mysql.get("query", ""),
        "http_timestamp": nginx.get("@timestamp", "") if nginx else "",
        "http_method": nginx.get("method", "") if nginx else "",
        "http_status_code": nginx.get("code", "") if nginx else "",
        "http_response_time": nginx.get("rt", "") if nginx else "",
        "http_path": nginx.get("path", "") if nginx else "",
        "http_query": nginx.get("query", "") if nginx else "",
        "http_client_ip": nginx.get("remote", "") if nginx else "",
        "http_x_forwarded_for": nginx.get("http_x_forwarded_for", "") if nginx else "",
        "http_referer": nginx.get("referer", "") if nginx else "",
        "http_user_agent": nginx.get("agent", "") if nginx else "",
    }

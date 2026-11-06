"""Time-based matching of MySQL slow queries and Nginx access records."""

from datetime import datetime, timezone
from collections import defaultdict
from bisect import bisect_right
import re
from typing import Any, Dict, Iterable, Iterator, List, Optional

MAX_TIMESTAMP_GAP: float = 120.0  # seconds (2 menit)

# Daftar user Nginx yang menjadi proxy/kandidat untuk user MySQL 'akange'.
# Tidak ada log nginx access.akange, sehingga pencarian dilakukan di log-log ini.
AKANGE_PROXY_USERS: frozenset = frozenset({"simbca", "simbcav3", "simcorn", "simpadi"})

# Kata-kata SQL yang tidak relevan sebagai kata kunci path/referer matching.
_SQL_STOP_WORDS: frozenset = frozenset({
    "select", "from", "where", "and", "or", "not", "in", "is", "null",
    "like", "between", "join", "left", "right", "inner", "outer", "on",
    "group", "order", "by", "having", "limit", "offset", "union", "all",
    "distinct", "as", "case", "when", "then", "else", "end", "into",
    "values", "set", "update", "delete", "insert", "create", "drop",
    "alter", "table", "index", "view", "exists", "count", "sum", "avg",
    "min", "max", "coalesce", "ifnull", "if", "concat", "length", "trim",
    "upper", "lower", "substring", "date", "year", "month", "day",
    "now", "curdate", "timestamp", "cast", "convert", "char", "varchar",
    "int", "bigint", "float", "double", "decimal", "bool", "boolean",
    "true", "false", "asc", "desc", "with", "recursive", "rowcount",
    "aggregate", "use", "database", "schema",
})


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
    """Join MySQL queries to the closest plausible Nginx request.

    Nginx timestamps must be strictly later than the MySQL timestamp, but
    within 2 minutes (mysql_time < nginx_time <= mysql_time + 120).
    Existing user and response-duration checks still apply. Older/incomplete
    records retain exact-second key matching only when their full timestamps
    still satisfy the later-than rule.
    """
    by_timestamp: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    timed_nginx_by_user: Dict[str, List[tuple[float, Dict[str, Any]]]] = defaultdict(list)
    for nginx in nginx_records:
        key = _timestamp_key(nginx.get("@timestamp"))
        if key is not None:
            by_timestamp[key].append(nginx)
        timestamp = parse_timestamp(nginx.get("@timestamp"))
        if timestamp is not None:
            log_user = _log_user(nginx.get("_log_name"))
            if log_user:
                timed_nginx_by_user[log_user].append((timestamp, nginx))
    nginx_times_by_user = {}
    for log_user, records in timed_nginx_by_user.items():
        records.sort(key=lambda item: item[0])
        nginx_times_by_user[log_user] = [timestamp for timestamp, _ in records]

    # Bangun pool gabungan Nginx untuk user 'akange'.
    # Karena tidak ada log access.akange, kita cari di log proxy (simbca, simcorn, simpadi, dll).
    akange_nginx_pool: List[tuple] = []
    for proxy_user in AKANGE_PROXY_USERS:
        akange_nginx_pool.extend(timed_nginx_by_user.get(proxy_user, []))
    akange_nginx_pool.sort(key=lambda item: item[0])
    akange_times: List[float] = [t for t, _ in akange_nginx_pool]

    # Setiap MySQL record mencari pasangan Nginx-nya secara independen.
    # Tidak ada prioritas antar user — kecocokan ditentukan murni oleh rules:
    #   • timestamp window (mysql_time < nginx_time ≤ mysql_time + 120s)
    #   • response time  (query_time ≤ rt ≤ query_time + 120)
    #   • untuk akange: keyword query cocok ke path/referer Nginx
    # Jika dua MySQL record dari user berbeda sama-sama memenuhi rules untuk
    # satu Nginx record yang sama, keduanya valid dan muncul di output.
    for mysql in mysql_records:
        mysql_time = parse_timestamp(mysql.get("@timestamp"))
        user = _normal_user(mysql.get("user"))
        query_time = _duration(mysql.get("query_time"))
        closest_strict_match: Optional[Dict[str, Any]] = None

        if mysql_time is not None:
            if user == "akange" and query_time is not None:
                # Strategi khusus akange: cari di pool proxy (simbca/simcorn/simpadi)
                # lalu saring dengan kata kunci dari query SQL ke path/referer.
                keywords = _extract_query_keywords(mysql.get("query", ""))
                start = bisect_right(akange_times, mysql_time)
                end = bisect_right(akange_times, mysql_time + MAX_TIMESTAMP_GAP)
                for record_index in range(start, end):
                    _, nginx = akange_nginx_pool[record_index]
                    response_time = _duration(nginx.get("rt"))
                    if response_time is None:
                        continue
                    if not (query_time <= response_time <= query_time + 120):
                        continue
                    # Keyword wajib cocok jika berhasil diekstrak.
                    if keywords and not _akange_path_matches(nginx, keywords):
                        continue
                    closest_strict_match = nginx
                    break
            elif user and query_time is not None:
                # User native: cari langsung di index log miliknya sendiri.
                user_records = timed_nginx_by_user.get(user, [])
                user_times = nginx_times_by_user.get(user, [])
                start = bisect_right(user_times, mysql_time)
                end = bisect_right(user_times, mysql_time + MAX_TIMESTAMP_GAP)
                for record_index in range(start, end):
                    _, nginx = user_records[record_index]
                    response_time = _duration(nginx.get("rt"))
                    if response_time is None:
                        continue
                    if query_time <= response_time <= query_time + 120:
                        closest_strict_match = nginx
                        break

        # Fallback: exact-second matching untuk record tanpa metadata lengkap.
        if closest_strict_match is not None:
            matches = [closest_strict_match]
        else:
            key = _timestamp_key(mysql.get("@timestamp"))
            matches = by_timestamp.get(key, []) if key is not None else []
            matches = [
                n for n in matches
                if mysql_time is not None
                and (nginx_time := parse_timestamp(n.get("@timestamp"))) is not None
                and nginx_time > mysql_time
                and _needs_legacy_fallback(mysql, n, user, query_time)
            ]
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
        "mysql_query": mysql.get("query", ""),
        "http_timestamp": nginx.get("@timestamp", "") if nginx else "",
        "http_log_name": nginx.get("_log_name", "") if nginx else "",
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


def _duration(value: Any) -> Optional[float]:
    """Parse a duration in seconds, including common unit-suffixed values."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)\s*(ms|s)?\s*", str(value), re.I)
    if not match:
        return None
    number = float(match.group(1))
    return number / 1000 if (match.group(2) or "").lower() == "ms" else number


def _normal_user(value: Any) -> str:
    return _account_key(value)


def _log_user(value: Any) -> str:
    """Extract and normalize the account suffix from an Nginx log name."""
    name = str(value or "").strip().lower()
    name = name.rsplit(".", 1)[-1] if "." in name else name
    return _account_key(name)


def _account_key(value: Any) -> str:
    """Normalize account-name variants such as ``simbcav3`` and ``simbca-v3``."""
    name = str(value or "").strip().lower()
    # Nginx log names can insert an API namespace before the account name.
    name = re.sub(r"^api[-_.]+", "", name)
    # The MySQL and Nginx sources inconsistently include separators before
    # deployment/version suffixes (for example ``simbcav3`` vs ``simbca-v3``).
    # Standardize version suffixes like ``-v3`` or ``_v3`` to ``v3`` without
    # removing the version number.
    name = re.sub(r"[-_.]?v0*(\d+)$", r"v\1", name)
    # Ignore remaining punctuation/case differences while preserving digits
    # that are part of the account itself.
    return re.sub(r"[^a-z0-9]", "", name)


def _needs_legacy_fallback(
    mysql: Dict[str, Any], nginx: Dict[str, Any], user: str, query_time: Optional[float]
) -> bool:
    """Allow exact-second matching only when a required field is unavailable."""
    log_name = str(nginx.get("_log_name") or "").strip()
    response_time = _duration(nginx.get("rt"))
    if user and log_name and _log_user(log_name) != user:
        return False
    if user and log_name and query_time is not None and response_time is not None:
        # All correlation fields exist, so an invalid timing pair must stay rejected.
        return False
    return True


# ---------------------------------------------------------------------------
# Fungsi helper khusus akange
# ---------------------------------------------------------------------------

_SQL_IDENTIFIER = re.compile(r"`?([a-zA-Z_][a-zA-Z0-9_]*)`?")
# Karakter pemisah dalam URL path/referer
_URL_SEG = re.compile(r"[^a-z0-9]+")


def _extract_query_keywords(query: Any) -> frozenset:
    """Ekstrak kata kunci bermakna dari teks SQL untuk dicocokkan ke path/referer Nginx.

    Langkah-langkah:
    1. Hapus preamble SET timestamp=... dan USE db;
    2. Hapus string literal dan komentar agar tidak ada false positive
    3. Ambil semua identifier SQL (nama tabel, kolom, alias)
    4. Filter: hanya identifier >= 4 karakter dan bukan SQL stop word
    """
    if not query:
        return frozenset()
    text = str(query)
    # Hapus preamble umum dari MySQL slow log
    text = re.sub(r"(?i)^\s*SET\s+timestamp\s*=\s*\d+\s*;?\s*", "", text)
    text = re.sub(r"(?i)^\s*USE\s+\w+\s*;?\s*", "", text)
    # Hapus string literal agar isinya tidak ikut diekstrak
    text = re.sub(r"'(?:''|\\'|[^'])*'", " ", text)
    text = re.sub(r'"(?:""|\\"|[^"])*"', " ", text)
    # Hapus komentar SQL
    text = re.sub(r"/\*[\s\S]*?\*/", " ", text)
    text = re.sub(r"--[^\r\n]*", " ", text)
    text = re.sub(r"#[^\r\n]*", " ", text)
    # Ekstrak identifier
    identifiers: set = set()
    for m in _SQL_IDENTIFIER.finditer(text):
        word = m.group(1).lower()
        if len(word) >= 4 and word not in _SQL_STOP_WORDS:
            identifiers.add(word)
    return frozenset(identifiers)


def _akange_path_matches(nginx: Dict[str, Any], keywords: frozenset) -> bool:
    """Cek apakah path atau referer Nginx mengandung setidaknya satu kata kunci query.

    Kata kunci dicocokkan sebagai substring setelah membagi URL path menjadi
    segmen-segmen kecil (pisah berdasarkan karakter non-alfanumerik).
    """
    path = str(nginx.get("path") or "").lower()
    referer = str(nginx.get("referer") or "").lower()
    combined = path + " " + referer
    # Bagi URL menjadi token (segmen path, parameter, dll.)
    segments = set(_URL_SEG.split(combined))
    # Cocokkan: keyword ada di salah satu segmen, ATAU segmen ada di dalam keyword
    # (menangani kasus keyword lebih panjang dari segmen, misal 'ttbsks' vs 'ttbs-ks')
    for kw in keywords:
        if kw in combined:
            return True
        # Cek juga apakah keyword merupakan substring dari segmen path gabungan
        for seg in segments:
            if seg and len(seg) >= 4 and seg in kw:
                return True
    return False


_SQL_STRING = re.compile(r"'(?:''|\\.|[^'])*'|\"(?:\"\"|\\.|[^\"])*\"")
_SQL_NUMBER = re.compile(r"\b(?:0x[0-9a-f]+|0b[01]+|\d+(?:\.\d+)?(?:e[+-]?\d+)?)\b", re.IGNORECASE)
_SQL_IN_VALUES = re.compile(r"\bIN\s*\(\s*\?(?:\s*,\s*\?)*\s*\)", re.IGNORECASE)
_SQL_OPERATIONS = {
    "SELECT", "INSERT", "REPLACE", "UPDATE", "DELETE", "CREATE", "ALTER",
    "DROP", "TRUNCATE", "RENAME", "BEGIN", "START", "COMMIT", "ROLLBACK",
    "SAVEPOINT", "RELEASE", "CALL", "EXECUTE", "SHOW", "DESCRIBE", "EXPLAIN",
}


def query_operation(query: Any) -> str:
    """Return the top-level SQL operation, or OTHER if it is not recognized."""
    if query is None:
        return "OTHER"
    text = str(query)
    text = re.sub(r"/\*[\s\S]*?\*/", " ", text)
    text = re.sub(r"--[^\r\n]*|\#[^\r\n]*", " ", text)
    depth = 0
    quote = ""
    index = 0
    while index < len(text):
        char = text[index]
        if quote:
            if char == "\\":
                index += 2
                continue
            if char == quote:
                if index + 1 < len(text) and text[index + 1] == quote:
                    index += 2
                    continue
                quote = ""
            index += 1
            continue
        if char in ("'", '"', "`"):
            quote = char
            index += 1
            continue
        if char == "(":
            depth += 1
            index += 1
            continue
        if char == ")":
            depth = max(0, depth - 1)
            index += 1
            continue
        if depth == 0 and (char.isalpha() or char == "_"):
            end = index + 1
            while end < len(text) and (text[end].isalnum() or text[end] == "_"):
                end += 1
            word = text[index:end].upper()
            if word in _SQL_OPERATIONS:
                return word
            index = end
            continue
        index += 1
    return "OTHER"


def normalize_query(query: Any, *, max_length: int = 1200) -> str:
    """Return a compact SQL shape by replacing literal values with placeholders."""
    if query is None:
        return ""
    text = str(query).strip()
    if not text:
        return ""

    # Replace literals before stripping comments so comment-like text inside a
    # quoted value is retained as a single placeholder.
    text = _SQL_STRING.sub("?", text)
    text = re.sub(r"/\*[\s\S]*?\*/", " ", text)
    text = re.sub(r"--[^\r\n]*|\#[^\r\n]*", " ", text)
    text = _SQL_NUMBER.sub("?", text)
    text = _SQL_IN_VALUES.sub("IN (?)", text)
    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_length:
        text = text[:max_length].rstrip() + " …[query shape shortened]"
    return text

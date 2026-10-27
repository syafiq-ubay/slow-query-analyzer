"""Time-based matching of MySQL slow queries and Nginx access records."""

from datetime import datetime, timezone
from collections import defaultdict
import re
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
        "mysql_query": mysql.get("query", ""),
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

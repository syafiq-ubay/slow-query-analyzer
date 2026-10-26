"""Paginated Elasticsearch queries for MySQL slow logs and Nginx access logs."""

import logging
import os
import shlex
import time
from typing import Any, Dict, Generator, List, Optional

import requests

from exporter import MYSQL_FIELDS, NGINX_FIELDS


def _load_dotenv(path: str = ".env") -> None:
    """Load simple KEY=VALUE entries without overriding shell environment values."""
    if not os.path.isfile(path):
        return
    with open(path, encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if line.startswith("export "):
                line = line[7:].lstrip()
            key, separator, value = line.partition("=")
            key = key.strip()
            if not separator or not key or key in os.environ:
                continue
            value = value.strip()
            if value[:1] in ("'", '"'):
                try:
                    parsed = shlex.split(value, comments=True)
                    value = parsed[0] if parsed else ""
                except (ValueError, IndexError):
                    value = value[1:]
            else:
                value = value.split(" #", 1)[0].rstrip()
            os.environ[key] = value


_load_dotenv()


logger = logging.getLogger(__name__)

ELASTIC_URL = os.getenv("ELASTIC_URL", "")
ELASTIC_USER = os.getenv("ELASTIC_USER", "")
ELASTIC_PASS = os.getenv("ELASTIC_PASS", "")
MYSQL_DATA_VIEW_ID = os.getenv("MYSQL_DATA_VIEW_ID", "")
NGINX_DATA_VIEW_ID = os.getenv("NGINX_DATA_VIEW_ID", "")

SEARCH_PAGE_SIZE = 1000


class ElasticLogClient:
    def __init__(
        self,
        base_url: str = ELASTIC_URL,
        username: str = ELASTIC_USER,
        password: str = ELASTIC_PASS,
        verify_ssl: bool = True,
    ) -> None:
        self.base_url = base_url.rstrip("/")
        self.session = requests.Session()
        self.session.auth = (username, password) if username or password else None
        self.session.verify = verify_ssl
        self.session.headers.update({
            "kbn-xsrf": "true",
            "Content-Type": "application/json",
            "elastic-api-version": "1",
        })
        self._resolved_data_views: Dict[str, str] = {}

    def resolve_index_pattern(self, data_view_id: str) -> str:
        if data_view_id in self._resolved_data_views:
            return self._resolved_data_views[data_view_id]

        endpoints = (
            f"{self.base_url}/api/data_views/data_view/{data_view_id}",
            f"{self.base_url}/api/saved_objects/index-pattern/{data_view_id}",
        )
        errors = []
        for url in endpoints:
            try:
                response = self.session.get(url, timeout=15)
                response.raise_for_status()
                data = response.json()
                title = (
                    data.get("data_view", {}).get("title")
                    or data.get("attributes", {}).get("title")
                )
                if title:
                    self._resolved_data_views[data_view_id] = title
                    logger.info("Resolved data view %s to %s", data_view_id, title)
                    return title
                errors.append(f"{url}: response did not contain a data-view title")
            except (requests.RequestException, ValueError) as exc:
                errors.append(f"{url}: {exc}")

        details = "; ".join(errors)
        raise RuntimeError(
            f"Unable to resolve Kibana data view '{data_view_id}'. {details}"
        )

    @staticmethod
    def _build_query(
        start_time: str,
        end_time: str,
        timestamp_field: str,
        fields: List[str],
        exclude_codes: Optional[List[str]],
        exclude_code_field: str,
        query_filter: Optional[str],
        page_size: int,
        search_after: Optional[List[Any]],
    ) -> Dict[str, Any]:
        body: Dict[str, Any] = {
            "size": page_size,
            "track_total_hits": False,
            "query": {
                "bool": {
                    "filter": [{
                        "range": {
                            timestamp_field: {
                                "gte": start_time,
                                "lte": end_time,
                                "format": "strict_date_optional_time||epoch_millis",
                            }
                        }
                    }]
                }
            },
            "sort": [{timestamp_field: {"order": "desc"}}],
            "_source": fields,
        }
        if exclude_codes:
            body["query"]["bool"]["must_not"] = [
                {"terms": {exclude_code_field: exclude_codes}}
            ]
        if query_filter:
            body["query"]["bool"].setdefault("filter", []).append(
                {"query_string": {"query": query_filter}}
            )
        if search_after:
            body["search_after"] = search_after
        return body

    def _post_search(
        self, index_pattern: str, body: Dict[str, Any]
    ) -> Dict[str, Any]:
        urls = (
            f"{self.base_url}/{index_pattern}/_search",
            f"{self.base_url}:9200/{index_pattern}/_search",
            f"{self.base_url}/elasticsearch/{index_pattern}/_search",
        )
        failures = []
        for url in urls:
            try:
                response = self.session.post(url, json=body, timeout=45)
                if response.status_code == 200:
                    return response.json()
                details = response.text.strip().replace("\n", " ")[:300]
                failures.append(f"{url}: HTTP {response.status_code} {details}")
            except (requests.RequestException, ValueError) as exc:
                failures.append(f"{url}: {exc}")

        # Kibana's Console proxy requires the Elasticsearch method as a query
        # parameter; omitting it can result in a rejected request.
        proxy_url = f"{self.base_url}/api/console/proxy"
        try:
            response = self.session.post(
                proxy_url,
                params={"path": f"/{index_pattern}/_search", "method": "POST"},
                json=body,
                timeout=45,
            )
            if response.status_code == 200:
                return response.json()
            details = response.text.strip().replace("\n", " ")[:300]
            failures.append(
                f"{response.url}: HTTP {response.status_code} {details}"
            )
        except (requests.RequestException, ValueError) as exc:
            failures.append(f"{proxy_url}: {exc}")

        bsearch_url = f"{self.base_url}/internal/bsearch"
        payload = {
            "batch": [{
                "request": {
                    "params": {"index": index_pattern, "body": body}
                }
            }]
        }
        try:
            response = self.session.post(bsearch_url, json=payload, timeout=45)
            response.raise_for_status()
            result = response.json().get("result", {})
            if result.get("isRunning"):
                search_id = result.get("id")
                for _ in range(60):
                    time.sleep(0.5)
                    poll_response = self.session.post(
                        bsearch_url,
                        json={"batch": [{"request": {"id": search_id}}]},
                        timeout=15,
                    )
                    poll_response.raise_for_status()
                    result = poll_response.json().get("result", {})
                    if not result.get("isRunning"):
                        break
                else:
                    raise RuntimeError("Kibana asynchronous search did not complete in 30 seconds.")
            raw_response = result.get("rawResponse")
            if raw_response is not None:
                return raw_response
            details = str(response.json())[:300]
            failures.append(f"{bsearch_url}: response contained no search result: {details}")
        except (requests.RequestException, ValueError) as exc:
            failures.append(f"{bsearch_url}: {exc}")

        raise RuntimeError(
            f"Elasticsearch search failed for index '{index_pattern}': "
            + "; ".join(failures)
        )

    def stream_records(
        self,
        *,
        data_view_id: str,
        start_time: str,
        end_time: str,
        fields: List[str],
        timestamp_field: str = "@timestamp",
        exclude_codes: Optional[List[str]] = None,
        exclude_code_field: str = "code",
        query_filter: Optional[str] = None,
        max_records: int = 0,
        page_size: int = SEARCH_PAGE_SIZE,
    ) -> Generator[List[Dict[str, Any]], None, None]:
        if page_size <= 0:
            raise ValueError("page_size must be greater than zero.")
        if max_records < 0:
            raise ValueError("max_records cannot be negative.")

        index_pattern = self.resolve_index_pattern(data_view_id)
        print(f"      [Elastic] Data view: {data_view_id} ({index_pattern})")
        print(f"      [Elastic] Range: {start_time} -> {end_time}")
        if exclude_codes:
            print(f"      [Elastic] Excluding {exclude_code_field}: {', '.join(exclude_codes)}")

        total = 0
        page_number = 0
        search_after = None
        while max_records == 0 or total < max_records:
            request_size = page_size
            if max_records:
                request_size = min(page_size, max_records - total)
            body = self._build_query(
                start_time,
                end_time,
                timestamp_field,
                fields,
                exclude_codes,
                exclude_code_field,
                query_filter,
                request_size,
                search_after,
            )
            result = self._post_search(index_pattern, body)
            hits = result.get("hits", {}).get("hits", [])
            if not hits:
                break

            records = [
                {field: hit.get("_source", {}).get(field, "") for field in fields}
                for hit in hits
            ]
            total += len(records)
            page_number += 1
            print(f"      [Elastic] Page {page_number}: {len(records)} records (total {total:,})")
            yield records

            search_after = hits[-1].get("sort")
            if not search_after:
                raise RuntimeError(
                    "Elasticsearch hit is missing a sort cursor; stopping pagination "
                    "to avoid returning duplicate or incomplete records."
                )
            if len(hits) < request_size:
                break

        print(f"      [Elastic] Complete: {total:,} records")

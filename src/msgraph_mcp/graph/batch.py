"""Graph $batch executor.

Provides a generic chunked batch executor against Microsoft Graph's /$batch
endpoint, decoupled from the SDK via the BatchTransport protocol so it can
be unit-tested with a fake transport.
"""

from __future__ import annotations

import asyncio
from collections.abc import Iterator
from dataclasses import dataclass
from typing import Any, Protocol, runtime_checkable


@dataclass
class BatchRequest:
    """A single sub-request to include in a Graph $batch call."""

    id: str
    method: str
    url: str
    body: dict[str, Any] | None = None
    headers: dict[str, str] | None = None


@dataclass
class BatchResult:
    """Per-sub-request outcome, in the same order as the caller's input."""

    id: str
    ok: bool
    status: int = 0
    error: str | None = None
    body: dict[str, Any] | None = None


@runtime_checkable
class BatchTransport(Protocol):
    """Minimal HTTP transport for Graph $batch.

    Implementations POST the payload to https://graph.microsoft.com/v1.0/$batch
    and return the parsed JSON body.
    """

    async def post_batch(self, payload: dict[str, Any]) -> dict[str, Any]: ...


async def execute_batch(
    transport: BatchTransport,
    requests: list[BatchRequest],
    *,
    chunk_size: int = 20,
) -> list[BatchResult]:
    """Execute Graph $batch in chunks. Preserves input order in results.

    Sub-requests within a single $batch are numbered "1", "2", ... per chunk;
    that numbering is internal to the executor and never exposed to callers.
    """
    if not requests:
        return []

    results_by_id: dict[str, BatchResult] = {}

    for chunk in _chunked(requests, chunk_size):
        subid_to_caller: dict[str, str] = {}
        payload_requests: list[dict[str, Any]] = []
        for i, req in enumerate(chunk, start=1):
            sub_id = str(i)
            subid_to_caller[sub_id] = req.id
            entry: dict[str, Any] = {
                "id": sub_id,
                "method": req.method,
                "url": req.url,
            }
            if req.body is not None:
                entry["body"] = req.body
                entry["headers"] = {"Content-Type": "application/json"}
            if req.headers:
                merged = dict(entry.get("headers", {}))
                merged.update(req.headers)
                entry["headers"] = merged
            payload_requests.append(entry)

        response = await transport.post_batch({"requests": payload_requests})

        throttled = _throttled_entries(response, payload_requests)
        if throttled:
            wait = _max_retry_after(response.get("responses", []))
            await asyncio.sleep(wait or 0)
            # Renumber the retry chunk starting at 1 again.
            retry_payload_requests: list[dict[str, Any]] = []
            retry_subid_to_caller: dict[str, str] = {}
            for j, orig in enumerate(throttled, start=1):
                new_sub_id = str(j)
                retry_subid_to_caller[new_sub_id] = subid_to_caller[orig["id"]]
                retry_payload_requests.append({**orig, "id": new_sub_id})
            retry_response = await transport.post_batch({"requests": retry_payload_requests})

            # Merge non-throttled originals with retry results.
            keep = [r for r in response.get("responses", []) if r.get("status") != 429]
            for sub in keep:
                caller_id = subid_to_caller.get(str(sub.get("id")))
                if caller_id is not None:
                    results_by_id[caller_id] = _sub_to_result(caller_id, sub)
            for sub in retry_response.get("responses", []):
                caller_id = retry_subid_to_caller.get(str(sub.get("id")))
                if caller_id is not None:
                    results_by_id[caller_id] = _sub_to_result(caller_id, sub)
        else:
            for sub in response.get("responses", []):
                caller_id = subid_to_caller.get(str(sub.get("id")))
                if caller_id is not None:
                    results_by_id[caller_id] = _sub_to_result(caller_id, sub)

    # Preserve input order; fill in any missing as failures.
    out: list[BatchResult] = []
    for req in requests:
        out.append(
            results_by_id.get(req.id)
            or BatchResult(id=req.id, ok=False, status=0, error="No response from $batch")
        )
    return out


def _chunked(xs: list[BatchRequest], n: int) -> Iterator[list[BatchRequest]]:
    for i in range(0, len(xs), n):
        yield xs[i : i + n]


def _sub_to_result(caller_id: str, sub: dict[str, Any]) -> BatchResult:
    status = int(sub.get("status") or 0)
    ok = 200 <= status < 300
    body = sub.get("body") if isinstance(sub.get("body"), dict) else None
    error = None if ok else _format_error(status, body)
    return BatchResult(id=caller_id, ok=ok, status=status, error=error, body=body)


def _format_error(status: int, body: dict[str, Any] | None) -> str:
    if isinstance(body, dict):
        err = body.get("error", {}) if isinstance(body.get("error"), dict) else {}
        code = err.get("code", "")
        message = err.get("message", "")
        prefix = " ".join(filter(None, [str(status), code]))
        return f"{prefix}: {message}" if message else prefix
    return str(status)


def _throttled_entries(response: dict[str, Any], payload_requests: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Return the original payload entries whose responses came back 429."""
    throttled_ids = {
        str(r.get("id"))
        for r in response.get("responses", [])
        if r.get("status") == 429
    }
    return [p for p in payload_requests if p["id"] in throttled_ids]


def _max_retry_after(responses: list[dict[str, Any]]) -> float | None:
    max_ra: float | None = None
    for r in responses:
        if r.get("status") != 429:
            continue
        ra = (r.get("headers") or {}).get("Retry-After")
        if ra is None:
            continue
        try:
            v = float(ra)
        except (TypeError, ValueError):
            continue
        if max_ra is None or v > max_ra:
            max_ra = v
    return max_ra


class GraphBatchTransport:
    """Default BatchTransport using msgraph-sdk's request adapter.

    Posts to https://graph.microsoft.com/v1.0/$batch and returns the parsed
    JSON. Reuses the existing GraphClient's auth and retry policy.

    Implementation note: if msgraph-sdk grows first-class batch primitives
    (BatchRequestContentCollection / BatchResponseContent) that fit the
    BatchTransport contract more cleanly, this class is the only place that
    needs to change.
    """

    GRAPH_BATCH_URL = "https://graph.microsoft.com/v1.0/$batch"

    def __init__(self, graph) -> None:
        self._graph = graph

    async def post_batch(self, payload: dict[str, Any]) -> dict[str, Any]:
        import json

        from kiota_abstractions.method import Method  # type: ignore[import-untyped]
        from kiota_abstractions.request_information import RequestInformation  # type: ignore[import-untyped]

        ri = RequestInformation()
        ri.url = self.GRAPH_BATCH_URL
        ri.http_method = Method.POST
        ri.headers.try_add("Content-Type", "application/json")
        ri.content = json.dumps(payload).encode("utf-8")

        raw = await self._graph.raw.request_adapter.send_primitive_async(ri, "bytes", {})
        return json.loads(raw)

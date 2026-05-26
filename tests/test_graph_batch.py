import asyncio
from unittest.mock import AsyncMock, MagicMock

import pytest

from outlook_mcp.graph.batch import BatchRequest, BatchResult, BatchTransport, GraphBatchTransport, execute_batch


def test_batch_request_dataclass_defaults():
    req = BatchRequest(id="m1", method="POST", url="/me/messages/m1/move")
    assert req.id == "m1"
    assert req.method == "POST"
    assert req.url == "/me/messages/m1/move"
    assert req.body is None
    assert req.headers is None


def test_batch_result_dataclass_defaults():
    res = BatchResult(id="m1", ok=True, status=200)
    assert res.id == "m1"
    assert res.ok is True
    assert res.status == 200
    assert res.error is None
    assert res.body is None


def test_batch_transport_is_runtime_checkable_protocol():
    class Stub:
        async def post_batch(self, payload):
            return {"responses": []}

    # Protocol type check at runtime — Stub satisfies BatchTransport.
    assert isinstance(Stub(), BatchTransport)


class FakeTransport:
    """Records each post_batch call and returns canned responses in order."""

    def __init__(self, responses: list[dict]) -> None:
        self.calls: list[dict] = []
        self._responses = list(responses)

    async def post_batch(self, payload):
        self.calls.append(payload)
        return self._responses.pop(0)


@pytest.mark.asyncio
async def test_execute_batch_empty_input_returns_empty():
    transport = FakeTransport(responses=[])
    out = await execute_batch(transport, [])
    assert out == []
    assert transport.calls == []


@pytest.mark.asyncio
async def test_execute_batch_single_chunk_preserves_order():
    transport = FakeTransport(responses=[{
        "responses": [
            {"id": "1", "status": 204},
            {"id": "2", "status": 204},
            {"id": "3", "status": 204},
        ],
    }])
    reqs = [
        BatchRequest(id="m1", method="POST", url="/me/messages/m1/move", body={"destinationId": "archive"}),
        BatchRequest(id="m2", method="POST", url="/me/messages/m2/move", body={"destinationId": "archive"}),
        BatchRequest(id="m3", method="POST", url="/me/messages/m3/move", body={"destinationId": "archive"}),
    ]
    out = await execute_batch(transport, reqs)
    assert [r.id for r in out] == ["m1", "m2", "m3"]
    assert all(r.ok for r in out)
    assert all(r.status == 204 for r in out)
    # Payload shape sanity
    assert len(transport.calls) == 1
    sent = transport.calls[0]["requests"]
    assert [s["id"] for s in sent] == ["1", "2", "3"]
    assert all(s["method"] == "POST" for s in sent)
    assert sent[0]["url"] == "/me/messages/m1/move"
    assert sent[0]["body"] == {"destinationId": "archive"}
    assert sent[0]["headers"] == {"Content-Type": "application/json"}


@pytest.mark.asyncio
async def test_execute_batch_chunks_at_default_size_20():
    # 45 requests → 3 chunks of 20/20/5
    reqs = [
        BatchRequest(id=f"m{i}", method="POST", url=f"/me/messages/m{i}/move", body={"destinationId": "archive"})
        for i in range(45)
    ]
    transport = FakeTransport(responses=[
        {"responses": [{"id": str(j + 1), "status": 204} for j in range(20)]},
        {"responses": [{"id": str(j + 1), "status": 204} for j in range(20)]},
        {"responses": [{"id": str(j + 1), "status": 204} for j in range(5)]},
    ])
    out = await execute_batch(transport, reqs)
    assert [r.id for r in out] == [f"m{i}" for i in range(45)]
    assert all(r.ok for r in out)
    assert len(transport.calls) == 3
    assert [len(c["requests"]) for c in transport.calls] == [20, 20, 5]


@pytest.mark.asyncio
async def test_execute_batch_mixed_success_failure_in_one_chunk():
    transport = FakeTransport(responses=[{
        "responses": [
            {"id": "1", "status": 204},
            {"id": "2", "status": 404, "body": {"error": {"code": "ErrorItemNotFound", "message": "not found"}}},
            {"id": "3", "status": 500, "body": {"error": {"code": "InternalServerError", "message": "boom"}}},
        ],
    }])
    reqs = [
        BatchRequest(id="m1", method="POST", url="/me/messages/m1/move"),
        BatchRequest(id="m2", method="POST", url="/me/messages/m2/move"),
        BatchRequest(id="m3", method="POST", url="/me/messages/m3/move"),
    ]
    out = await execute_batch(transport, reqs)
    assert out[0].ok and out[0].status == 204 and out[0].error is None
    assert not out[1].ok and out[1].status == 404
    assert "ErrorItemNotFound" in (out[1].error or "")
    assert "not found" in (out[1].error or "")
    assert not out[2].ok and out[2].status == 500


@pytest.mark.asyncio
async def test_execute_batch_missing_response_marked_failed():
    # Transport returns no entry for one of the sub-ids.
    transport = FakeTransport(responses=[{
        "responses": [{"id": "1", "status": 204}],
    }])
    reqs = [
        BatchRequest(id="m1", method="POST", url="/me/messages/m1/move"),
        BatchRequest(id="m2", method="POST", url="/me/messages/m2/move"),
    ]
    out = await execute_batch(transport, reqs)
    assert out[0].ok
    assert not out[1].ok
    assert out[1].error is not None and "No response" in out[1].error


@pytest.mark.asyncio
async def test_execute_batch_429_waits_and_retries_only_throttled(monkeypatch):
    sleeps: list[float] = []

    async def fake_sleep(secs):
        sleeps.append(secs)

    monkeypatch.setattr(asyncio, "sleep", fake_sleep)

    transport = FakeTransport(responses=[
        {
            "responses": [
                {"id": "1", "status": 204},
                {"id": "2", "status": 429, "headers": {"Retry-After": "2"}},
                {"id": "3", "status": 429, "headers": {"Retry-After": "1"}},
            ],
        },
        {
            # Retry chunk contains only the throttled sub-requests; renumbered.
            "responses": [
                {"id": "1", "status": 204},
                {"id": "2", "status": 204},
            ],
        },
    ])
    reqs = [
        BatchRequest(id="m1", method="POST", url="/me/messages/m1/move"),
        BatchRequest(id="m2", method="POST", url="/me/messages/m2/move"),
        BatchRequest(id="m3", method="POST", url="/me/messages/m3/move"),
    ]
    out = await execute_batch(transport, reqs)
    assert [r.ok for r in out] == [True, True, True]
    assert sleeps == [2.0]  # max Retry-After across the throttled subs
    assert len(transport.calls) == 2
    assert len(transport.calls[1]["requests"]) == 2
    # Retried subs are exactly the originally-throttled ones, in caller order.
    retry_urls = [s["url"] for s in transport.calls[1]["requests"]]
    assert retry_urls == ["/me/messages/m2/move", "/me/messages/m3/move"]


@pytest.mark.asyncio
async def test_execute_batch_429_persistent_failure_reported(monkeypatch):
    async def fake_sleep(secs):  # noqa: ARG001
        return None

    monkeypatch.setattr(asyncio, "sleep", fake_sleep)

    transport = FakeTransport(responses=[
        {"responses": [{"id": "1", "status": 429, "headers": {"Retry-After": "1"}}]},
        {"responses": [{"id": "1", "status": 429, "headers": {"Retry-After": "1"}}]},
    ])
    reqs = [BatchRequest(id="m1", method="POST", url="/me/messages/m1/move")]
    out = await execute_batch(transport, reqs)
    assert not out[0].ok
    assert out[0].status == 429



@pytest.mark.asyncio
async def test_graph_batch_transport_posts_to_v1_batch_endpoint():
    # Mock graph.raw.request_adapter.send_primitive_async to capture the
    # RequestInformation we build, and return canned bytes.
    fake_response = b'{"responses":[{"id":"1","status":204}]}'
    send_async = AsyncMock(return_value=fake_response)
    graph = MagicMock()
    graph.raw.request_adapter.send_primitive_async = send_async

    transport = GraphBatchTransport(graph)
    payload = {"requests": [{"id": "1", "method": "POST", "url": "/me/messages/m1/move"}]}
    out = await transport.post_batch(payload)

    assert out == {"responses": [{"id": "1", "status": 204}]}
    assert send_async.await_count == 1
    # The RequestInformation passed in should target the v1.0 $batch endpoint
    # with method POST and JSON body equal to our payload.
    assert send_async.await_args is not None
    args, _ = send_async.await_args
    req_info = args[0]
    assert "/v1.0/$batch" in (req_info.url or "")
    # http_method may be an enum value or string depending on SDK version.
    assert str(req_info.http_method).lower().endswith("post")
    # Body bytes should JSON-decode to our payload
    import json
    assert json.loads(req_info.content.decode("utf-8")) == payload

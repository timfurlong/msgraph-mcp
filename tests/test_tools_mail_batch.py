from unittest.mock import MagicMock

import pytest

from outlook_mcp.graph.batch import BatchResult
from outlook_mcp.graph.errors import GraphValidationError
from outlook_mcp.tools import mail_batch


@pytest.mark.asyncio
async def test_batch_archive_messages_happy_path(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=204) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    graph = MagicMock()
    out = await mail_batch.batch_archive_messages(
        graph=graph, message_ids=["m1", "m2", "m3"]
    )

    assert out["summary"] == {"total": 3, "succeeded": 3, "failed": 0}
    assert [r["id"] for r in out["results"]] == ["m1", "m2", "m3"]
    assert all(r["ok"] for r in out["results"])
    # URLs should be the per-id Graph /move endpoint pointing at "archive".
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1/move",
        "/me/messages/m2/move",
        "/me/messages/m3/move",
    ]
    assert all(r.body == {"destinationId": "archive"} for r in captured["requests"])


@pytest.mark.asyncio
async def test_batch_archive_messages_uses_users_endpoint_when_mailbox_set(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=204) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    out = await mail_batch.batch_archive_messages(
        graph=MagicMock(),
        message_ids=["m1"],
        mailbox="shared@example.com",
    )
    assert out["summary"]["total"] == 1
    assert captured["requests"][0].url == "/users/shared@example.com/messages/m1/move"


@pytest.mark.asyncio
async def test_batch_archive_messages_validates_empty_input():
    with pytest.raises(GraphValidationError):
        await mail_batch.batch_archive_messages(graph=MagicMock(), message_ids=[])


@pytest.mark.asyncio
async def test_batch_archive_messages_validates_max_input():
    with pytest.raises(GraphValidationError):
        await mail_batch.batch_archive_messages(
            graph=MagicMock(), message_ids=["m"] * 1001
        )


@pytest.mark.asyncio
async def test_batch_archive_messages_reports_per_id_failures(monkeypatch):
    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        return [
            BatchResult(id="m1", ok=True, status=204),
            BatchResult(id="m2", ok=False, status=404, error="404 ErrorItemNotFound: gone"),
        ]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    out = await mail_batch.batch_archive_messages(
        graph=MagicMock(), message_ids=["m1", "m2"]
    )
    assert out["summary"] == {"total": 2, "succeeded": 1, "failed": 1}
    assert out["results"][1]["error"] == "404 ErrorItemNotFound: gone"


@pytest.mark.asyncio
async def test_batch_move_messages_sends_destination_id(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=204) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    await mail_batch.batch_move_messages(
        graph=MagicMock(),
        message_ids=["m1", "m2"],
        destination="inbox",
    )
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1/move",
        "/me/messages/m2/move",
    ]
    assert all(r.body == {"destinationId": "inbox"} for r in captured["requests"])


@pytest.mark.asyncio
async def test_batch_move_messages_validates_destination_required():
    with pytest.raises(GraphValidationError):
        await mail_batch.batch_move_messages(
            graph=MagicMock(), message_ids=["m1"], destination=""
        )


@pytest.mark.asyncio
async def test_batch_move_messages_empty_input_validates():
    with pytest.raises(GraphValidationError):
        await mail_batch.batch_move_messages(
            graph=MagicMock(), message_ids=[], destination="archive"
        )


@pytest.mark.asyncio
async def test_batch_mark_read_uses_patch_isRead_true(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=200) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    await mail_batch.batch_mark_read(graph=MagicMock(), message_ids=["m1", "m2"])

    assert [r.method for r in captured["requests"]] == ["PATCH", "PATCH"]
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1",
        "/me/messages/m2",
    ]
    assert all(r.body == {"isRead": True} for r in captured["requests"])


@pytest.mark.asyncio
async def test_batch_mark_unread_uses_patch_isRead_false(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=200) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    await mail_batch.batch_mark_unread(graph=MagicMock(), message_ids=["m1", "m2"])

    assert [r.method for r in captured["requests"]] == ["PATCH", "PATCH"]
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1",
        "/me/messages/m2",
    ]
    assert all(r.body == {"isRead": False} for r in captured["requests"])


@pytest.mark.asyncio
async def test_batch_flag_messages_patches_flag_flagged(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=200) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    await mail_batch.batch_flag_messages(graph=MagicMock(), message_ids=["m1", "m2"])

    assert [r.method for r in captured["requests"]] == ["PATCH", "PATCH"]
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1",
        "/me/messages/m2",
    ]
    assert all(r.body == {"flag": {"flagStatus": "flagged"}} for r in captured["requests"])


@pytest.mark.asyncio
async def test_batch_unflag_messages_patches_flag_notFlagged(monkeypatch):
    captured = {}

    async def fake_execute_batch(transport, requests, *, chunk_size=20):
        captured["requests"] = requests
        return [BatchResult(id=r.id, ok=True, status=200) for r in requests]

    monkeypatch.setattr(mail_batch, "execute_batch", fake_execute_batch)
    monkeypatch.setattr(mail_batch, "GraphBatchTransport", lambda graph: object())

    await mail_batch.batch_unflag_messages(graph=MagicMock(), message_ids=["m1", "m2"])

    assert [r.method for r in captured["requests"]] == ["PATCH", "PATCH"]
    assert [r.url for r in captured["requests"]] == [
        "/me/messages/m1",
        "/me/messages/m2",
    ]
    assert all(r.body == {"flag": {"flagStatus": "notFlagged"}} for r in captured["requests"])


def test_register_wires_all_six_batch_tools():
    """register(mcp, graph=graph) decorates six MCP tools by name."""
    registered_names: list[str] = []

    class FakeMCP:
        def tool(self, *, name, description=""):  # noqa: ARG002
            registered_names.append(name)

            def deco(fn):
                return fn

            return deco

    mail_batch.register(FakeMCP(), graph=MagicMock())

    assert registered_names == [
        "batch_archive_messages",
        "batch_move_messages",
        "batch_mark_read",
        "batch_mark_unread",
        "batch_flag_messages",
        "batch_unflag_messages",
    ]

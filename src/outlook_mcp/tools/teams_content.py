"""Teams inline hosted-content download.

Inline images and other hosted content live behind an authenticated Graph
endpoint (.../hostedContents/{id}/$value) that requires the server's
delegated token, so an agent cannot fetch them directly. This tool proxies
that authenticated fetch and returns base64, parallel to download_attachment
for mail. SharePoint/OneDrive-backed file attachments are a different Graph
surface and are out of scope (see README).
"""

from __future__ import annotations

import base64

from outlook_mcp.auth.token import NotAuthenticatedError
from outlook_mcp.graph.errors import GraphValidationError, map_kiota_error
from outlook_mcp.graph.trimming import trim_hosted_content_download


def _sniff_content_type(data: bytes) -> str | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return None


def _hosted_content_builder(graph, *, chat_id, team_id, channel_id, message_id, hosted_content_id):
    if chat_id is not None:
        message = graph.raw.chats.by_chat_id(chat_id).messages.by_chat_message_id(message_id)
    else:
        message = (
            graph.raw.teams.by_team_id(team_id)
            .channels.by_channel_id(channel_id)
            .messages.by_chat_message_id(message_id)
        )
    return message.hosted_contents.by_chat_message_hosted_content_id(hosted_content_id).content


async def download_hosted_content(
    *,
    graph,
    message_id: str,
    hosted_content_id: str,
    chat_id: str | None = None,
    team_id: str | None = None,
    channel_id: str | None = None,
    include_raw: bool = False,
) -> dict:
    """Download one inline hosted-content item (usually an image) as base64.

    Specify the message location with exactly one of:
      - chat_id (for a chat message), or
      - team_id AND channel_id (for a channel message).

    Get message_id and hosted_content_id from a message's hosted_content_refs
    (returned by list_chat_messages / list_channel_messages / list_message_replies).

    Args:
        message_id: Graph id of the message the content is attached to.
        hosted_content_id: Graph hosted-content id (from hosted_content_refs).
        chat_id: Chat id, if the message is in a chat.
        team_id: Team id, if the message is in a channel.
        channel_id: Channel id, if the message is in a channel.
        include_raw: Include the raw payload under "raw".

    Returns:
        {"content_type": str | None, "size_bytes": int, "content_base64": str | None}
        content_type is sniffed from the bytes (Graph does not return it on
        the $value endpoint); it may be None for unrecognized formats.
    """
    has_chat = chat_id is not None
    has_team = team_id is not None
    has_channel = channel_id is not None
    if has_chat and (has_team or has_channel):
        raise GraphValidationError("Provide either chat_id or (team_id and channel_id), not both.")
    if not has_chat and not (has_team and has_channel):
        raise GraphValidationError(
            "Provide a message location: chat_id, or both team_id and channel_id."
        )

    builder = _hosted_content_builder(
        graph, chat_id=chat_id, team_id=team_id, channel_id=channel_id,
        message_id=message_id, hosted_content_id=hosted_content_id,
    )
    try:
        data = await builder.get()
    except NotAuthenticatedError:
        raise
    except Exception as exc:  # noqa: BLE001
        raise map_kiota_error(exc) from exc

    if isinstance(data, (bytes, bytearray)):
        raw = {
            "contentType": _sniff_content_type(bytes(data)),
            "size": len(data),
            "contentBytes": base64.b64encode(data).decode("ascii"),
        }
    else:
        raw = {"contentType": None, "size": 0, "contentBytes": None}
    return trim_hosted_content_download(raw, include_raw=include_raw)


def register(mcp, *, graph) -> None:
    @mcp.tool(name="download_hosted_content", description=download_hosted_content.__doc__ or "")
    async def _download_hosted_content(
        message_id: str,
        hosted_content_id: str,
        chat_id: str | None = None,
        team_id: str | None = None,
        channel_id: str | None = None,
        include_raw: bool = False,
    ):
        return await download_hosted_content(
            graph=graph, message_id=message_id, hosted_content_id=hosted_content_id,
            chat_id=chat_id, team_id=team_id, channel_id=channel_id, include_raw=include_raw,
        )

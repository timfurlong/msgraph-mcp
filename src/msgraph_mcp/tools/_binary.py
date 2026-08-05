"""Shared helpers for tools that return downloaded binary content.

Image bytes are returned to the client as a native MCP image block
(mcp.server.fastmcp.utilities.types.Image) so agents see the image directly
instead of a base64 text payload; save_path writes bytes to disk instead and
returns only metadata. Non-image content without save_path falls back to the
legacy base64 dict.
"""

from __future__ import annotations

import pathlib

from mcp.server.fastmcp.utilities.types import Image

_EXT_BY_CONTENT_TYPE = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/webp": ".webp",
}


def is_image(content_type: str | None) -> bool:
    return bool(content_type) and content_type in _EXT_BY_CONTENT_TYPE


def ext_for(content_type: str | None) -> str:
    return _EXT_BY_CONTENT_TYPE.get(content_type or "", ".bin")


def image_result(meta: dict, data: bytes, content_type: str) -> list:
    """Metadata dict + native MCP image block (FastMCP renders both)."""
    return [meta, Image(data=data, format=content_type.removeprefix("image/"))]


def write_bytes(save_path: str, data: bytes, *, default_name: str) -> str:
    """Write data to save_path and return the resolved path.

    save_path may be a file path, or an existing directory (default_name is
    appended). Parent directories are created as needed.
    """
    target = pathlib.Path(save_path).expanduser()
    if target.is_dir():
        target = target / default_name
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(data)
    return str(target)

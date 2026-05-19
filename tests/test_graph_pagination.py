import pytest

from outlook_mcp.graph import pagination


def test_encode_decode_roundtrip():
    url = "https://graph.microsoft.com/v1.0/me/messages?$skip=25"
    token = pagination.encode_next_link(url)
    assert token is not None
    assert token != url  # opaque
    assert "graph.microsoft.com" not in token
    assert pagination.decode_page_token(token) == url


def test_encode_none_returns_none():
    assert pagination.encode_next_link(None) is None


def test_decode_invalid_token_raises_validation_error():
    with pytest.raises(pagination.GraphValidationError):
        pagination.decode_page_token("not-base64!")


def test_decode_token_with_non_graph_url_raises():
    bad = pagination.encode_next_link("https://evil.example.com/foo")
    assert bad is not None
    with pytest.raises(pagination.GraphValidationError):
        pagination.decode_page_token(bad)


def test_validate_limit_accepts_range():
    assert pagination.validate_limit(1) == 1
    assert pagination.validate_limit(25) == 25
    assert pagination.validate_limit(100) == 100


def test_validate_limit_rejects_out_of_range():
    with pytest.raises(pagination.GraphValidationError):
        pagination.validate_limit(0)
    with pytest.raises(pagination.GraphValidationError):
        pagination.validate_limit(101)

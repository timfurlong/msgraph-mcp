from types import SimpleNamespace

from outlook_mcp.graph import errors


def test_graph_validation_error_is_runtime_error():
    err = errors.GraphValidationError("bad limit")
    assert isinstance(err, RuntimeError)
    assert "bad limit" in str(err)


def test_graph_api_error_carries_status_and_code():
    err = errors.GraphAPIError(status=404, code="ResourceNotFound", message="No such message")
    assert err.status == 404
    assert err.code == "ResourceNotFound"
    assert err.message == "No such message"
    # __str__ should be informative
    text = str(err)
    assert "404" in text
    assert "ResourceNotFound" in text
    assert "No such message" in text


def test_map_kiota_error_extracts_inner_error():
    # Simulate the shape kiota gives us: an object with .response_status_code
    # and .error.code / .error.message attributes.
    inner = SimpleNamespace(code="ErrorItemNotFound", message="Item was not found")
    kiota_err = SimpleNamespace(response_status_code=404, error=inner)

    mapped = errors.map_kiota_error(kiota_err)
    assert isinstance(mapped, errors.GraphAPIError)
    assert mapped.status == 404
    assert mapped.code == "ErrorItemNotFound"
    assert mapped.message == "Item was not found"


def test_map_kiota_error_handles_missing_inner():
    kiota_err = SimpleNamespace(response_status_code=500, error=None)
    mapped = errors.map_kiota_error(kiota_err)
    assert mapped.status == 500
    assert mapped.code == "Unknown"


def test_map_kiota_error_handles_bare_exception():
    mapped = errors.map_kiota_error(ValueError("boom"))
    assert mapped.status == 0
    assert "boom" in mapped.message

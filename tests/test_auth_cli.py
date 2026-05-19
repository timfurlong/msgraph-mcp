from unittest.mock import MagicMock

from outlook_mcp.auth import cli


def _setup_fakes(monkeypatch, *, init_flow=None, acquire_result=None):
    fake_app = MagicMock()
    fake_cache = MagicMock()
    fake_cache.has_state_changed = True
    if init_flow is not None:
        fake_app.initiate_device_flow.return_value = init_flow
    if acquire_result is not None:
        fake_app.acquire_token_by_device_flow.return_value = acquire_result
    persist_calls = []
    monkeypatch.setattr(cli, "build_app", lambda: (fake_app, fake_cache))
    monkeypatch.setattr(cli, "persist_cache", lambda c: persist_calls.append(c))
    return fake_app, fake_cache, persist_calls


def test_login_happy_path(monkeypatch, capsys):
    _, _, persist_calls = _setup_fakes(
        monkeypatch,
        init_flow={
            "message": "Visit https://aka.ms/devicelogin and enter ABC123",
            "user_code": "ABC123",
        },
        acquire_result={
            "access_token": "tok",
            "id_token_claims": {"preferred_username": "me@example.com"},
        },
    )
    exit_code = cli.run_login()
    assert exit_code == 0
    captured = capsys.readouterr()
    assert "Visit https://aka.ms/devicelogin" in captured.out
    assert "me@example.com" in captured.out
    assert len(persist_calls) == 1


def test_login_fails_when_initiate_device_flow_returns_no_user_code(monkeypatch, capsys):
    _setup_fakes(monkeypatch, init_flow={"error": "boom"})
    exit_code = cli.run_login()
    assert exit_code != 0
    captured = capsys.readouterr()
    # Stderr or stdout — implementation may print to either
    assert "boom" in (captured.out + captured.err)


def test_login_fails_when_acquire_returns_error(monkeypatch, capsys):
    _setup_fakes(
        monkeypatch,
        init_flow={"message": "Visit url and enter ABC", "user_code": "ABC"},
        acquire_result={"error": "authorization_declined", "error_description": "User declined"},
    )
    exit_code = cli.run_login()
    assert exit_code != 0
    captured = capsys.readouterr()
    assert "authorization_declined" in (captured.out + captured.err)

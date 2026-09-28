"""Shape tests for the Teams Graph helper — no live tokens, no network."""
from __future__ import annotations

from dashboard.plugin_api import missing_scope, teams_platform_enabled


def test_missing_scope_detects_graph_forbidden():
    payload = {
        "error": {
            "code": "Forbidden",
            "message": "Missing scope permissions on the request. API requires one of 'Chat.Read'.",
        }
    }
    assert missing_scope(payload) is True


def test_missing_scope_false_on_plain_error():
    assert missing_scope({"error": {"message": "Not found"}}) is False


def test_teams_platform_enabled_reads_config(tmp_path, monkeypatch):
    home = tmp_path / "hermes"
    home.mkdir()
    (home / "config.yaml").write_text(
        "platforms:\n  telegram:\n    enabled: true\n  teams:\n    enabled: true\n",
        encoding="utf-8",
    )
    monkeypatch.setenv("HERMES_HOME", str(home))
    # Re-import helper bound to env — function reads get_hermes_home each call.
    from dashboard import plugin_api as api

    monkeypatch.setattr(api, "get_hermes_home", lambda: home)
    assert api.teams_platform_enabled() is True

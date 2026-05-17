from fastapi.testclient import TestClient

import webapp.dependencies as dependencies


class SpyCleanupService:
    def __init__(self):
        self.calls = []
        self.allowed_runtime_dirs = []

    def cleanup_runtime_files(self, *, reason: str) -> None:
        self.calls.append(reason)


def test_cleanup_runs_only_on_startup_and_shutdown(monkeypatch):
    spy = SpyCleanupService()
    monkeypatch.setattr(dependencies, "get_cleanup_service", lambda: spy)

    from webapp.app import app

    with TestClient(app) as client:
        assert spy.calls == ["startup"]
        response = client.get("/")
        assert response.status_code == 200
        assert spy.calls == ["startup"]

    assert spy.calls == ["startup", "shutdown"]

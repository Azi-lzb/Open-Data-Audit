"""S5 synthetic tests never persist schema snapshots in a real user directory."""
import pytest


@pytest.fixture(autouse=True)
def isolated_s5_runtime(tmp_path, monkeypatch):
    from s1audit_configured import sources, numeric_scope

    monkeypatch.setattr(sources, "DEFAULT_SCHEMA_SNAPSHOT_DIR", tmp_path / ".schema_snapshots")
    monkeypatch.setattr(numeric_scope, "TEMPLATE_DIR", tmp_path / "templates")

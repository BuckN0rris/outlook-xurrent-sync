"""Make the project root importable for the test suite, plus a shared config fixture."""

import os
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent))

#: every setting the program reads, so a developer's real environment or .env
#: cannot leak into a test
_SETTING_PREFIXES = (
    "XURRENT_", "GRAPH_", "OPEN_", "DONE_", "WAITING_", "MENTION_", "LOOKBACK_",
    "STATE_", "LOG_", "DAILY_", "VALIDATE_", "SUBJECT_", "HEADER_", "AFTER_", "MULTI_",
)  # fmt: skip


@pytest.fixture
def app_config(monkeypatch: pytest.MonkeyPatch, tmp_path: Path):
    """A valid Config whose state, logs and token cache all live in tmp_path."""
    from outlook_xurrent.config import Config

    for key in list(os.environ):
        if key.startswith(_SETTING_PREFIXES):
            monkeypatch.delenv(key, raising=False)
    settings = {
        "XURRENT_TOKEN": "test-token",
        "XURRENT_ACCOUNT": "example-it",
        "OPEN_CATEGORY": "Nicht erledigt",
        "DONE_CATEGORY": "Erledigt",
        "STATE_FILE": str(tmp_path / "state.json"),
        "LOG_DIR": str(tmp_path / "logs"),
        "DAILY_LOG_DIR": str(tmp_path / "daily"),
        "GRAPH_TOKEN_CACHE": str(tmp_path / "msal_cache.json"),
    }
    for name, value in settings.items():
        monkeypatch.setenv(name, value)
    return Config.from_env()

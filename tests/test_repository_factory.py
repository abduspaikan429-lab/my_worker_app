from pathlib import Path
from unittest.mock import patch
import pytest

from config import settings
from core.repository_factory import get_worker_repository
from repositories.sqlite_worker_repository import SqliteWorkerRepository
from repositories.worker_repository import JsonWorkerRepository


def test_json_config_returns_json_repository():
    """1. Test config WORKER_STORAGE='json' returns JsonWorkerRepository."""
    with patch.object(settings, "WORKER_STORAGE", "json"):
        repo = get_worker_repository()
        assert isinstance(repo, JsonWorkerRepository)


def test_sqlite_config_returns_sqlite_repository(tmp_path: Path):
    """2. Test modified config WORKER_STORAGE='sqlite' returns SqliteWorkerRepository."""
    test_db_path = tmp_path / "test_factory.db"
    with patch.object(settings, "WORKER_STORAGE", "sqlite"), patch.object(
        settings, "SQLITE_DB_PATH", str(test_db_path)
    ):
        repo = get_worker_repository()
        assert isinstance(repo, SqliteWorkerRepository)
        assert repo.db_path == test_db_path


def test_explicit_storage_type_override(tmp_path: Path):
    """Test passing explicit storage_type parameter overrides settings."""
    test_db_path = tmp_path / "explicit_test.db"
    with patch.object(settings, "SQLITE_DB_PATH", str(test_db_path)):
        repo_json = get_worker_repository(storage_type="json")
        assert isinstance(repo_json, JsonWorkerRepository)

        repo_sqlite = get_worker_repository(storage_type="sqlite")
        assert isinstance(repo_sqlite, SqliteWorkerRepository)


def test_invalid_config_raises_value_error():
    """3. Test invalid config mode raises ValueError."""
    with patch.object(settings, "WORKER_STORAGE", "unsupported_db"):
        with pytest.raises(ValueError) as exc_info:
            get_worker_repository()
        assert "Unsupported WORKER_STORAGE mode" in str(exc_info.value)

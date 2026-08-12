from datetime import datetime
from pathlib import Path
import shutil
from typing import Dict, List, Union

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_BACKUP_DIR = BASE_DIR / "data" / "backups"
DEFAULT_DB_PATH = BASE_DIR / "data" / "worker.db"
DEFAULT_JSON_PATH = BASE_DIR / "data" / "master_state.json"


class BackupService:
    """Service handling system backup, restoration, listing, and cleanup."""

    def __init__(self, backup_dir: Union[str, Path] = DEFAULT_BACKUP_DIR):
        self.backup_dir = Path(backup_dir)

    def backup_database(
        self, db_path: Union[str, Path] = DEFAULT_DB_PATH
    ) -> Path:
        """Backup SQLite database to backup_dir with timestamp filename."""
        src_path = Path(db_path)
        if not src_path.exists():
            raise FileNotFoundError(f"Database file not found: {src_path}")

        self.backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        dest_filename = f"worker_{timestamp}.db"
        dest_path = self.backup_dir / dest_filename

        shutil.copy2(src_path, dest_path)
        return dest_path

    def backup_json(
        self, json_path: Union[str, Path] = DEFAULT_JSON_PATH
    ) -> Path:
        """Backup JSON state file to backup_dir with timestamp filename."""
        src_path = Path(json_path)
        if not src_path.exists():
            raise FileNotFoundError(f"JSON state file not found: {src_path}")

        self.backup_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        dest_filename = f"master_{timestamp}.json"
        dest_path = self.backup_dir / dest_filename

        shutil.copy2(src_path, dest_path)
        return dest_path

    def list_backups(self) -> Dict[str, List[Path]]:
        """List all database and JSON backups in backup_dir."""
        if not self.backup_dir.exists():
            return {"db_backups": [], "json_backups": []}

        db_backups = sorted(list(self.backup_dir.glob("worker_*.db")))
        json_backups = sorted(list(self.backup_dir.glob("master_*.json")))

        return {
            "db_backups": db_backups,
            "json_backups": json_backups,
        }

    def cleanup_old_backups(self, max_backups: int = 30) -> Dict[str, int]:
        """Delete oldest backups exceeding max_backups limit."""
        backups = self.list_backups()
        db_deleted = 0
        json_deleted = 0

        # db backups
        if len(backups["db_backups"]) > max_backups:
            to_remove = backups["db_backups"][:-max_backups]
            for f in to_remove:
                f.unlink(missing_ok=True)
                db_deleted += 1

        # json backups
        if len(backups["json_backups"]) > max_backups:
            to_remove = backups["json_backups"][:-max_backups]
            for f in to_remove:
                f.unlink(missing_ok=True)
                json_deleted += 1

        return {"db_deleted": db_deleted, "json_deleted": json_deleted}

    def restore_database(
        self,
        backup_file: Union[str, Path],
        target_db_path: Union[str, Path] = DEFAULT_DB_PATH,
    ) -> Path:
        """Restore database from backup_file, creating a pre-restore backup first."""
        backup_src = Path(backup_file)
        if not backup_src.exists():
            raise FileNotFoundError(f"Backup file not found: {backup_src}")

        target_path = Path(target_db_path)
        target_path.parent.mkdir(parents=True, exist_ok=True)

        # 1. Create pre-restore safety backup if target DB exists
        pre_restore_backup = None
        if target_path.exists():
            pre_restore_backup = self.backup_database(target_path)

        # 2. Copy backup file to target DB
        shutil.copy2(backup_src, target_path)

        return pre_restore_backup or backup_src

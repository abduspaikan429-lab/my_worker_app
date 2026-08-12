from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from services.backup_service import BackupService


def main() -> None:
    service = BackupService()

    db_path = BASE_DIR / "data" / "worker.db"
    json_path = BASE_DIR / "data" / "master_state.json"

    db_backup = service.backup_database(db_path) if db_path.exists() else None
    json_backup = service.backup_json(json_path) if json_path.exists() else None

    service.cleanup_old_backups(max_backups=30)

    db_rel = db_backup.relative_to(BASE_DIR) if db_backup else "Not Found"
    json_rel = json_backup.relative_to(BASE_DIR) if json_backup else "Not Found"

    print("\nBackup completed")
    print("\nDatabase:")
    print(db_rel)
    print("\nJSON:")
    print(json_rel)


if __name__ == "__main__":
    main()

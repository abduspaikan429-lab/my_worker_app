from pathlib import Path
import sys

BASE_DIR = Path(__file__).resolve().parent.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from services.backup_service import BackupService


def main() -> None:
    service = BackupService()
    backups = service.list_backups()

    if len(sys.argv) > 1:
        backup_input = Path(sys.argv[1])
        if not backup_input.is_absolute():
            backup_input = BASE_DIR / backup_input
    else:
        db_backups = backups["db_backups"]
        if not db_backups:
            print("Error: No database backups found in data/backups/")
            return
        backup_input = db_backups[-1]  # Latest backup

    if not backup_input.exists():
        print(f"Error: Specified backup file {backup_input} does not exist!")
        return

    target_db = BASE_DIR / "data" / "worker.db"
    pre_backup = service.restore_database(backup_input, target_db_path=target_db)

    print("\nRestore completed")
    print(f"Restored from: {backup_input.relative_to(BASE_DIR)}")
    print(f"Target Database: {target_db.relative_to(BASE_DIR)}")
    if pre_backup:
        print(f"Safety Pre-restore Backup: {pre_backup.relative_to(BASE_DIR)}")


if __name__ == "__main__":
    main()

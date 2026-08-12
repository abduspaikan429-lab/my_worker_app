from datetime import datetime
import json
from pathlib import Path
from typing import Any, Dict, List, Union

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_BACKUP_PATH = BASE_DIR / "data" / "master_state_backup.json"


def backup_workers_to_json(
    workers: List[Dict[str, Any]],
    backup_path: Union[str, Path] = DEFAULT_BACKUP_PATH,
) -> None:
    """Export workers list to a backup JSON file (data/master_state_backup.json)."""
    path_obj = Path(backup_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    now_str = datetime.now().isoformat(timespec="seconds")
    backup_data = {
        "version": datetime.now().strftime("%Y%m%d_%H%M%S_%f"),
        "updated_at": now_str,
        "rows": workers,
    }

    with open(path_obj, "w", encoding="utf-8") as f:
        json.dump(backup_data, f, ensure_ascii=False, indent=4)

from datetime import datetime
from pathlib import Path
from typing import Union

BASE_DIR = Path(__file__).resolve().parent.parent
DEFAULT_LOG_PATH = BASE_DIR / "logs" / "sync.log"


def write_sync_log(
    message: str, log_path: Union[str, Path] = DEFAULT_LOG_PATH
) -> None:
    """Write timestamped message to sync log file."""
    path_obj = Path(log_path)
    path_obj.parent.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().isoformat(timespec="seconds")
    log_line = f"[{timestamp}] {message}\n"

    with open(path_obj, "a", encoding="utf-8") as f:
        f.write(log_line)

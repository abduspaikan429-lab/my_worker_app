import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

DEFAULT_STATE_FILE = Path("data/master_state.json")


class WorkerRepository:
    """Abstract base class for Worker Repository."""

    def get_all_workers(self) -> List[Dict[str, Any]]:
        raise NotImplementedError("Subclasses must implement get_all_workers()")

    def save_workers(self, workers: List[Dict[str, Any]]) -> None:
        raise NotImplementedError("Subclasses must implement save_workers()")

    def get_worker_by_id_card(self, id_card: str) -> Optional[Dict[str, Any]]:
        raise NotImplementedError("Subclasses must implement get_worker_by_id_card()")


class JsonWorkerRepository(WorkerRepository):
    """JSON implementation of WorkerRepository using master_state.json."""

    def __init__(self, file_path: Union[str, Path] = DEFAULT_STATE_FILE):
        self.file_path = Path(file_path)

    def get_all_workers(self) -> List[Dict[str, Any]]:
        """Read workers list from JSON file. Returns empty list if file does not exist."""
        if not self.file_path.exists():
            return []
        try:
            with open(self.file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data.get("rows", [])
            elif isinstance(data, list):
                return data
            return []
        except Exception:
            return []

    def save_workers(self, workers: List[Dict[str, Any]]) -> None:
        """Save workers list to JSON file with ensure_ascii=False and indent=4."""
        self.file_path.parent.mkdir(parents=True, exist_ok=True)

        state: Dict[str, Any] = {}
        if self.file_path.exists():
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    existing = json.load(f)
                if isinstance(existing, dict):
                    state = existing
            except Exception:
                state = {}

        state["rows"] = workers

        with open(self.file_path, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=4)

    def get_worker_by_id_card(self, id_card: str) -> Optional[Dict[str, Any]]:
        """Search worker by id_card (checking both '身份证号' and 'id_card')."""
        if not id_card:
            return None

        workers = self.get_all_workers()
        for worker in workers:
            if isinstance(worker, dict):
                card = worker.get("身份证号") or worker.get("id_card")
                if card == id_card:
                    return worker
        return None

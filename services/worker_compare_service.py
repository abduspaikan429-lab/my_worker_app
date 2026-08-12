from __future__ import annotations

from typing import Any, Dict, List


class WorkerCompareService:
    """Service to compare worker records between JSON repository and SQLite repository."""

    @staticmethod
    def compare_workers(
        json_workers: List[Dict[str, Any]],
        sqlite_workers: List[Dict[str, Any]],
    ) -> Dict[str, Any]:
        """Compare JSON and SQLite workers by id_card and name."""

        def extract_card_and_name(worker: Dict[str, Any]) -> tuple[str, str]:
            card = (
                worker.get("id_card") or worker.get("身份证号") or ""
            )
            name = worker.get("name") or worker.get("姓名") or ""
            return str(card).strip(), str(name).strip()

        json_map: Dict[str, str] = {}
        for w in json_workers:
            if isinstance(w, dict):
                card, name = extract_card_and_name(w)
                if card:
                    json_map[card] = name

        sqlite_map: Dict[str, str] = {}
        for w in sqlite_workers:
            if isinstance(w, dict):
                card, name = extract_card_and_name(w)
                if card:
                    sqlite_map[card] = name

        json_count = len(json_map)
        sqlite_count = len(sqlite_map)

        json_cards = set(json_map.keys())
        sqlite_cards = set(sqlite_map.keys())

        missing_in_sqlite = sorted(list(json_cards - sqlite_cards))
        missing_in_json = sorted(list(sqlite_cards - json_cards))

        different_name = []
        common_cards = sorted(list(json_cards & sqlite_cards))
        for card in common_cards:
            j_name = json_map[card]
            s_name = sqlite_map[card]
            if j_name != s_name:
                different_name.append(
                    {
                        "id_card": card,
                        "json_name": j_name,
                        "sqlite_name": s_name,
                    }
                )

        is_pass = (
            json_count == sqlite_count
            and not missing_in_sqlite
            and not missing_in_json
            and not different_name
        )

        return {
            "json_count": json_count,
            "sqlite_count": sqlite_count,
            "missing_in_sqlite": missing_in_sqlite,
            "missing_in_json": missing_in_json,
            "different_name": different_name,
            "result": "PASS" if is_pass else "FAIL",
        }

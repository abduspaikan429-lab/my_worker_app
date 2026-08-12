from __future__ import annotations

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from db.database import DEFAULT_DB_PATH, get_connection, init_db
from repositories.worker_repository import WorkerRepository


class SqliteWorkerRepository(WorkerRepository):
    """SQLite implementation of WorkerRepository."""

    def __init__(self, db_path: Union[str, Path] = DEFAULT_DB_PATH):
        self.db_path = Path(db_path)
        init_db(self.db_path)

    def clear_all_workers(self) -> None:
        """Clear all records in workers table."""
        conn = get_connection(self.db_path)
        try:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM workers;")
            conn.commit()
        finally:
            conn.close()

    def get_all_workers(self) -> List[Dict[str, Any]]:
        """Retrieve all workers from SQLite database."""
        conn = get_connection(self.db_path)
        try:
            cursor = conn.cursor()
            query = """
                SELECT 
                    w.*,
                    c.name AS company_name,
                    t.name AS team_name
                FROM workers w
                LEFT JOIN companies c ON w.company_id = c.id
                LEFT JOIN teams t ON w.team_id = t.id
                ORDER BY w.id ASC
            """
            cursor.execute(query)
            rows = cursor.fetchall()
            workers = []
            for row in rows:
                r_dict = dict(row)
                w_dict = {
                    "id": r_dict["id"],
                    "worker_code": r_dict["worker_code"] or "",
                    "序号": r_dict["worker_code"] or "",
                    "name": r_dict["name"] or "",
                    "姓名": r_dict["name"] or "",
                    "id_card": r_dict["id_card"] or "",
                    "身份证号": r_dict["id_card"] or "",
                    "phone": r_dict["phone"] or "",
                    "手机号": r_dict["phone"] or "",
                    "bank_card": r_dict["bank_card"] or "",
                    "工资卡号": r_dict["bank_card"] or "",
                    "company_id": r_dict["company_id"],
                    "分包/所属企业": r_dict["company_name"] or "",
                    "team_id": r_dict["team_id"],
                    "班组": r_dict["team_name"] or "",
                    "job_type": r_dict["job_type"] or "",
                    "工种": r_dict["job_type"] or "",
                    "status": r_dict["status"] or "active",
                    "在场/进退场状态": r_dict["status"] or "",
                    "entry_date": r_dict["entry_date"] or "",
                    "进场日期": r_dict["entry_date"] or "",
                    "exit_date": r_dict["exit_date"] or "",
                    "退场日期": r_dict["exit_date"] or "",
                    "created_at": r_dict["created_at"] or "",
                    "updated_at": r_dict["updated_at"] or "",
                }
                for k, v in r_dict.items():
                    if k not in w_dict:
                        w_dict[k] = v
                workers.append(w_dict)
            return workers
        finally:
            conn.close()

    def get_worker_by_id_card(self, id_card: str) -> Optional[Dict[str, Any]]:
        """Search worker by id_card in SQLite database."""
        if not id_card:
            return None
        conn = get_connection(self.db_path)
        try:
            cursor = conn.cursor()
            query = """
                SELECT 
                    w.*,
                    c.name AS company_name,
                    t.name AS team_name
                FROM workers w
                LEFT JOIN companies c ON w.company_id = c.id
                LEFT JOIN teams t ON w.team_id = t.id
                WHERE w.id_card = ?
            """
            cursor.execute(query, (id_card,))
            row = cursor.fetchone()
            if not row:
                return None

            r_dict = dict(row)
            w_dict = {
                "id": r_dict["id"],
                "worker_code": r_dict["worker_code"] or "",
                "序号": r_dict["worker_code"] or "",
                "name": r_dict["name"] or "",
                "姓名": r_dict["name"] or "",
                "id_card": r_dict["id_card"] or "",
                "身份证号": r_dict["id_card"] or "",
                "phone": r_dict["phone"] or "",
                "手机号": r_dict["phone"] or "",
                "bank_card": r_dict["bank_card"] or "",
                "工资卡号": r_dict["bank_card"] or "",
                "company_id": r_dict["company_id"],
                "分包/所属企业": r_dict["company_name"] or "",
                "team_id": r_dict["team_id"],
                "班组": r_dict["team_name"] or "",
                "job_type": r_dict["job_type"] or "",
                "工种": r_dict["job_type"] or "",
                "status": r_dict["status"] or "active",
                "在场/进退场状态": r_dict["status"] or "",
                "entry_date": r_dict["entry_date"] or "",
                "进场日期": r_dict["entry_date"] or "",
                "exit_date": r_dict["exit_date"] or "",
                "退场日期": r_dict["exit_date"] or "",
                "created_at": r_dict["created_at"] or "",
                "updated_at": r_dict["updated_at"] or "",
            }
            for k, v in r_dict.items():
                if k not in w_dict:
                    w_dict[k] = v
            return w_dict
        finally:
            conn.close()

    def save_workers(self, workers: List[Dict[str, Any]]) -> None:
        """Save/upsert workers into SQLite database."""
        if not workers:
            return

        conn = get_connection(self.db_path)
        now_str = datetime.now().isoformat(timespec="seconds")
        try:
            cursor = conn.cursor()
            for worker in workers:
                if not isinstance(worker, dict):
                    continue

                name = worker.get("name") or worker.get("姓名") or ""
                id_card = worker.get("id_card") or worker.get("身份证号") or None
                if not name or not id_card:
                    continue

                worker_code = (
                    worker.get("worker_code")
                    or worker.get("序号")
                    or worker.get("人员编号")
                    or None
                )
                if worker_code is not None:
                    worker_code = str(worker_code).strip()
                    if not worker_code:
                        worker_code = None

                if worker_code is not None:
                    cursor.execute("SELECT id_card FROM workers WHERE worker_code = ?", (worker_code,))
                    existing_code_row = cursor.fetchone()
                    if existing_code_row and existing_code_row["id_card"] != id_card:
                        worker_code = None

                phone = worker.get("phone") or worker.get("手机号") or None
                if phone is not None:
                    phone = str(phone).strip() or None

                bank_card = worker.get("bank_card") or worker.get("工资卡号") or None
                if bank_card is not None:
                    bank_card = str(bank_card).strip() or None

                job_type = worker.get("job_type") or worker.get("工种") or None
                status = (
                    worker.get("status")
                    or worker.get("在场/进退场状态")
                    or "active"
                )
                entry_date = (
                    worker.get("entry_date") or worker.get("进场日期") or None
                )
                exit_date = worker.get("exit_date") or worker.get("退场日期") or None

                company_name = (
                    worker.get("company_name")
                    or worker.get("分包/所属企业")
                    or None
                )
                team_name = worker.get("team_name") or worker.get("班组") or None

                company_id = None
                if company_name:
                    cursor.execute(
                        "SELECT id FROM companies WHERE name = ?",
                        (company_name,),
                    )
                    c_row = cursor.fetchone()
                    if c_row:
                        company_id = c_row["id"]
                    else:
                        cursor.execute(
                            "INSERT INTO companies (name, created_at, updated_at) VALUES (?, ?, ?)",
                            (company_name, now_str, now_str),
                        )
                        company_id = cursor.lastrowid

                team_id = None
                if team_name and company_id:
                    cursor.execute(
                        "SELECT id FROM teams WHERE company_id = ? AND name = ?",
                        (company_id, team_name),
                    )
                    t_row = cursor.fetchone()
                    if t_row:
                        team_id = t_row["id"]
                    else:
                        cursor.execute(
                            "INSERT INTO teams (company_id, name, created_at, updated_at) VALUES (?, ?, ?, ?)",
                            (company_id, team_name, now_str, now_str),
                        )
                        team_id = cursor.lastrowid

                cursor.execute(
                    """
                    INSERT INTO workers (
                        worker_code, name, id_card, phone, bank_card,
                        company_id, team_id, job_type, status,
                        entry_date, exit_date, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(id_card) DO UPDATE SET
                        worker_code = excluded.worker_code,
                        name = excluded.name,
                        phone = excluded.phone,
                        bank_card = excluded.bank_card,
                        company_id = excluded.company_id,
                        team_id = excluded.team_id,
                        job_type = excluded.job_type,
                        status = excluded.status,
                        entry_date = excluded.entry_date,
                        exit_date = excluded.exit_date,
                        updated_at = excluded.updated_at
                    """,
                    (
                        worker_code,
                        name,
                        id_card,
                        phone,
                        bank_card,
                        company_id,
                        team_id,
                        job_type,
                        status,
                        entry_date,
                        exit_date,
                        now_str,
                        now_str,
                    ),
                )
            conn.commit()
        finally:
            conn.close()

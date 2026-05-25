"""SQLite 관리 모듈."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterable

DB_PATH = Path(__file__).parent / "sqlite.db"


def set_db_path(path: str | Path) -> None:
    """런타임에 DB 경로를 변경 (웹/세션별 임시 DB 용)."""
    global DB_PATH
    DB_PATH = Path(path)


SCHEMA = """
CREATE TABLE IF NOT EXISTS farms (
    farm_name TEXT PRIMARY KEY,
    farm_id TEXT,
    owner TEXT,
    address TEXT
);

CREATE TABLE IF NOT EXISTS cattle (
    cattle_no TEXT PRIMARY KEY,
    farm_name TEXT NOT NULL,
    breed TEXT,
    sex TEXT,
    birth_date TEXT,
    mother_no TEXT,
    status TEXT DEFAULT '사육',
    status_date TEXT,
    first_seen_doc_date TEXT,
    last_seen_doc_date TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS processed_docs (
    doc_name TEXT NOT NULL,
    farm_name TEXT NOT NULL,
    doc_date TEXT,
    processed_at TEXT DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (doc_name, farm_name)
);

CREATE INDEX IF NOT EXISTS idx_cattle_farm ON cattle(farm_name);
CREATE INDEX IF NOT EXISTS idx_cattle_status ON cattle(status);
"""


@contextmanager
def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with connect() as conn:
        conn.executescript(SCHEMA)


def is_doc_processed(doc_name: str, farm_name: str) -> bool:
    with connect() as conn:
        cur = conn.execute(
            "SELECT 1 FROM processed_docs WHERE doc_name=? AND farm_name=?",
            (doc_name, farm_name),
        )
        return cur.fetchone() is not None


def mark_doc_processed(doc_name: str, farm_name: str, doc_date: str) -> None:
    with connect() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO processed_docs (doc_name, farm_name, doc_date) "
            "VALUES (?, ?, ?)",
            (doc_name, farm_name, doc_date),
        )


def upsert_farm(farm_name: str, farm_id: str | None, owner: str | None, address: str | None) -> None:
    with connect() as conn:
        conn.execute(
            """INSERT INTO farms (farm_name, farm_id, owner, address)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(farm_name) DO UPDATE SET
                 farm_id=excluded.farm_id,
                 owner=excluded.owner,
                 address=excluded.address""",
            (farm_name, farm_id, owner, address),
        )


def upsert_cattle_batch(rows: Iterable[dict], doc_date: str) -> tuple[int, int]:
    """rows: cattle_no, farm_name, breed, sex, birth_date, mother_no.
    Returns (inserted, updated_last_seen).
    """
    inserted = 0
    updated = 0
    with connect() as conn:
        for r in rows:
            cur = conn.execute("SELECT cattle_no FROM cattle WHERE cattle_no=?", (r["cattle_no"],))
            if cur.fetchone():
                conn.execute(
                    "UPDATE cattle SET last_seen_doc_date=?, farm_name=? WHERE cattle_no=?",
                    (doc_date, r["farm_name"], r["cattle_no"]),
                )
                updated += 1
            else:
                conn.execute(
                    """INSERT INTO cattle
                       (cattle_no, farm_name, breed, sex, birth_date, mother_no,
                        status, status_date, first_seen_doc_date, last_seen_doc_date)
                       VALUES (?, ?, ?, ?, ?, ?, '사육', ?, ?, ?)""",
                    (
                        r["cattle_no"],
                        r["farm_name"],
                        r.get("breed"),
                        r.get("sex"),
                        r.get("birth_date"),
                        r.get("mother_no"),
                        doc_date,
                        doc_date,
                        doc_date,
                    ),
                )
                inserted += 1
    return inserted, updated


def get_cattle_by_farm_doc_date(farm_name: str, doc_date: str) -> list[sqlite3.Row]:
    with connect() as conn:
        cur = conn.execute(
            "SELECT * FROM cattle WHERE farm_name=? AND last_seen_doc_date=?",
            (farm_name, doc_date),
        )
        return cur.fetchall()


def get_prev_doc_date(farm_name: str, current_doc_date: str) -> str | None:
    """해당 농장의 current 이전 가장 최근 doc_date."""
    with connect() as conn:
        cur = conn.execute(
            """SELECT DISTINCT doc_date FROM processed_docs
               WHERE farm_name=? AND doc_date < ?
               ORDER BY doc_date DESC LIMIT 1""",
            (farm_name, current_doc_date),
        )
        row = cur.fetchone()
        return row["doc_date"] if row else None


def get_cattle_active_at(farm_name: str, doc_date: str) -> list[sqlite3.Row]:
    """해당 농장에서 doc_date 시점에 last_seen_doc_date <= doc_date 이고 사육 상태인 소들."""
    with connect() as conn:
        cur = conn.execute(
            """SELECT * FROM cattle
               WHERE farm_name=? AND last_seen_doc_date=? """,
            (farm_name, doc_date),
        )
        return cur.fetchall()


def update_cattle_status(cattle_no: str, status: str, status_date: str | None) -> None:
    with connect() as conn:
        conn.execute(
            "UPDATE cattle SET status=?, status_date=?, updated_at=CURRENT_TIMESTAMP WHERE cattle_no=?",
            (status, status_date, cattle_no),
        )


def get_active_cattle_all() -> list[sqlite3.Row]:
    """현재 도축/폐사가 아닌 모든 농장 사육 중인 소."""
    with connect() as conn:
        cur = conn.execute(
            """SELECT * FROM cattle
               WHERE status NOT IN ('도축', '폐사')
               ORDER BY farm_name, cattle_no"""
        )
        return cur.fetchall()


def get_latest_doc_date(farm_name: str) -> str | None:
    with connect() as conn:
        cur = conn.execute(
            "SELECT MAX(doc_date) AS d FROM processed_docs WHERE farm_name=?",
            (farm_name,),
        )
        row = cur.fetchone()
        return row["d"] if row and row["d"] else None

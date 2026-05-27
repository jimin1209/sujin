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
    farm_id TEXT,            -- 농장식별번호 (farmUniqueNo)
    farm_no TEXT,            -- 농장번호 (farmNo, API 매칭용)
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
    acquisition_date TEXT,   -- 매입일 (우리 농장으로 들어온 날, API 사육지 이력 기준)
    first_seen_doc_date TEXT,
    last_seen_doc_date TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS processed_docs (
    doc_name TEXT NOT NULL,
    farm_name TEXT NOT NULL,
    doc_date TEXT,
    cattle_count INTEGER DEFAULT 0,
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


def _ensure_column(conn: sqlite3.Connection, table: str, column: str, col_type: str) -> None:
    cols = {row[1] for row in conn.execute(f"PRAGMA table_info({table})")}
    if column not in cols:
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_type}")


def init_db() -> None:
    with connect() as conn:
        conn.executescript(SCHEMA)
        # 마이그레이션 — 기존 DB 파일에 새 컬럼이 없으면 추가
        _ensure_column(conn, "farms", "farm_no", "TEXT")
        _ensure_column(conn, "cattle", "acquisition_date", "TEXT")
        _ensure_column(conn, "processed_docs", "cattle_count", "INTEGER DEFAULT 0")


def is_doc_processed(doc_name: str, farm_name: str) -> bool:
    with connect() as conn:
        cur = conn.execute(
            "SELECT 1 FROM processed_docs WHERE doc_name=? AND farm_name=?",
            (doc_name, farm_name),
        )
        return cur.fetchone() is not None


def mark_doc_processed(doc_name: str, farm_name: str, doc_date: str, cattle_count: int = 0) -> None:
    with connect() as conn:
        conn.execute(
            """INSERT INTO processed_docs (doc_name, farm_name, doc_date, cattle_count)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(doc_name, farm_name) DO UPDATE SET
                 doc_date=excluded.doc_date,
                 cattle_count=excluded.cattle_count""",
            (doc_name, farm_name, doc_date, cattle_count),
        )


def get_monthly_total_counts(year: int) -> dict[int, int]:
    """그 연도 1~12월의 월사육두수 (모든 농장 합산).

    각 농장별로 그 월의 최신 doc_date 의 cattle_count 를 잡고,
    doc 가 없는 월은 직전 월 값을 캐리오버. 그 후 농장 간 합산.
    """
    with connect() as conn:
        farms = [r["farm_name"] for r in conn.execute("SELECT DISTINCT farm_name FROM processed_docs")]
        result: dict[int, int] = {m: 0 for m in range(1, 13)}
        for farm in farms:
            cur = conn.execute(
                """SELECT substr(doc_date, 6, 2) AS mm, cattle_count, doc_date
                   FROM processed_docs
                   WHERE farm_name = ? AND substr(doc_date, 1, 4) = ?
                   ORDER BY doc_date""",
                (farm, str(year)),
            )
            # 농장별 월→cattle_count (같은 월에 여러 doc 면 가장 최근값)
            farm_monthly: dict[int, int] = {}
            for r in cur:
                farm_monthly[int(r["mm"])] = r["cattle_count"] or 0
            # 정방향 캐리오버
            last_val = 0
            for m in range(1, 13):
                if m in farm_monthly:
                    last_val = farm_monthly[m]
                result[m] += last_val
        return result


def upsert_farm(farm_name: str, farm_id: str | None, owner: str | None, address: str | None) -> None:
    with connect() as conn:
        conn.execute(
            """INSERT INTO farms (farm_name, farm_id, owner, address)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(farm_name) DO UPDATE SET
                 farm_id=COALESCE(excluded.farm_id, farms.farm_id),
                 owner=COALESCE(excluded.owner, farms.owner),
                 address=COALESCE(excluded.address, farms.address)""",
            (farm_name, farm_id, owner, address),
        )


def set_farm_no(farm_name: str, farm_no: str) -> None:
    with connect() as conn:
        conn.execute("UPDATE farms SET farm_no=? WHERE farm_name=?", (farm_no, farm_name))


def get_farm(farm_name: str) -> sqlite3.Row | None:
    with connect() as conn:
        cur = conn.execute("SELECT * FROM farms WHERE farm_name=?", (farm_name,))
        return cur.fetchone()


def get_all_farms() -> list[sqlite3.Row]:
    with connect() as conn:
        cur = conn.execute("SELECT * FROM farms")
        return cur.fetchall()


def update_cattle_acquisition(cattle_no: str, acquisition_date: str | None) -> None:
    with connect() as conn:
        conn.execute(
            "UPDATE cattle SET acquisition_date=? WHERE cattle_no=?",
            (acquisition_date, cattle_no),
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
    """현재 도축/폐사/양수도가 아닌 모든 농장 사육 중인 소."""
    with connect() as conn:
        cur = conn.execute(
            """SELECT * FROM cattle
               WHERE status NOT IN ('도축', '폐사', '양수도')
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

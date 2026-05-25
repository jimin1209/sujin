"""농장 상세 엑셀 (xls) 파서."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


FILENAME_RE = re.compile(
    r"^\s*(?P<farm>.+?)\s*-\s*(?P<yy>\d{2})[._-](?P<mm>\d{2})[._-](?P<dd>\d{2})\s*기준\s*\.(?:xls|xlsx)\s*$",
    re.IGNORECASE,
)


@dataclass
class FileMeta:
    farm_name: str
    doc_date: str  # YYYY-MM-DD
    doc_name: str


def parse_filename(path: Path) -> FileMeta | None:
    name = unicodedata.normalize("NFC", path.name)
    m = FILENAME_RE.match(name)
    if not m:
        return None
    yy, mm, dd = m["yy"], m["mm"], m["dd"]
    year = 2000 + int(yy)
    return FileMeta(
        farm_name=m["farm"].strip(),
        doc_date=f"{year:04d}-{int(mm):02d}-{int(dd):02d}",
        doc_name=name,
    )


def _normalize_cattle_no(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    digits = re.sub(r"\D", "", s)
    if len(digits) < 12:
        return None
    return digits[:12]


def _format_birth(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    m = re.match(r"^(\d{2})\.(\d{2})\.(\d{2})$", s)
    if m:
        yy, mm, dd = m.groups()
        return f"20{yy}-{mm}-{dd}"
    return s


def parse_farm_excel(path: Path) -> tuple[dict, list[dict]]:
    """
    Returns (farm_info, cattle_rows).
    farm_info: { farm_name, farm_id, owner, address }
    cattle_rows: [{ cattle_no, breed, sex, birth_date, mother_no }]
    """
    engine = "openpyxl" if path.suffix.lower() == ".xlsx" else "xlrd"
    df = pd.read_excel(path, sheet_name=0, header=None, engine=engine)

    farm_info = {"farm_name": None, "farm_id": None, "owner": None, "address": None}

    for i in range(min(10, len(df))):
        cell = df.iat[i, 0]
        if not isinstance(cell, str):
            continue
        m = re.search(r"농장명\s*:\s*([^\(\s]+).*농장식별번호\s*:\s*(\d+)", cell)
        if m:
            farm_info["farm_name"] = m.group(1).strip()
            farm_info["farm_id"] = m.group(2).strip()
        m = re.search(r"농장 주소\s*:\s*\(법정동\)\s*([^\n]+)", cell)
        if m:
            farm_info["address"] = m.group(1).strip()
        m = re.search(r"농장경영자명\s*:\s*([^,]+)", cell)
        if m:
            farm_info["owner"] = m.group(1).strip()

    # 데이터 시작 행 찾기 (개체식별번호 컬럼이 있는 헤더 이후)
    header_row = None
    for i in range(len(df)):
        cell = df.iat[i, 1] if df.shape[1] > 1 else None
        if isinstance(cell, str) and "개체식별번호" in cell:
            header_row = i
            break

    rows: list[dict] = []
    if header_row is None:
        return farm_info, rows

    start = header_row + 1
    for i in range(start, len(df)):
        cattle_no = _normalize_cattle_no(df.iat[i, 1])
        if not cattle_no:
            continue
        breed = df.iat[i, 4] if df.shape[1] > 4 else None
        sex = df.iat[i, 5] if df.shape[1] > 5 else None
        birth = df.iat[i, 6] if df.shape[1] > 6 else None
        mother = _normalize_cattle_no(df.iat[i, 10]) if df.shape[1] > 10 else None

        rows.append(
            {
                "cattle_no": cattle_no,
                "breed": str(breed).strip() if isinstance(breed, str) else None,
                "sex": str(sex).strip() if isinstance(sex, str) else None,
                "birth_date": _format_birth(birth),
                "mother_no": mother,
            }
        )

    return farm_info, rows

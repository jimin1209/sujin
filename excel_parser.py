"""농장 상세 파일 (xls / xlsx / pdf) 파서."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


FILENAME_RE = re.compile(
    r"^\s*(?P<farm>.+?)\s*-\s*(?P<yy>\d{2})[._-](?P<mm>\d{2})[._-](?P<dd>\d{2})\s*기준\s*\.(?:xls|xlsx|pdf)\s*$",
    re.IGNORECASE,
)

# PDF 한 줄 데이터 행:
#   일련번호 12자리식별 종류 성별 YY.MM.DD 개월령 12자리모개체 상태 YY.MM.DD 단축번호
PDF_ROW_RE = re.compile(
    r"^\s*(?P<seq>\d+)\s+(?P<no>\d{12})\s+(?P<breed>\S+)\s+(?P<sex>\S+)\s+"
    r"(?P<birth>\d{2}\.\d{2}\.\d{2})\s+(?P<months>\d+)\s+(?P<mother>\d{12})\s+"
    r"(?P<status>\S+)\s+(?P<reg>\d{2}\.\d{2}\.\d{2})\s+(?P<short>\S+)\s*$"
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


def parse_farm_pdf(path: Path) -> tuple[dict, list[dict]]:
    """PDF 형태 농장 사육현황 파서 (xls 와 동일한 리턴 형태)."""
    import pdfplumber  # 지연 임포트

    farm_info = {"farm_name": None, "farm_id": None, "owner": None, "address": None}
    rows: list[dict] = []

    with pdfplumber.open(path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            if page_idx == 0:
                m = re.search(r"농장명\s*:\s*([^\(\s]+).*?농장식별번호\s*:\s*(\d+)", text)
                if m:
                    farm_info["farm_name"] = m.group(1).strip()
                    farm_info["farm_id"] = m.group(2).strip()
                m = re.search(r"농장\s*주소\s*:\s*\(법정동\)\s*([^\n]+)", text)
                if m:
                    farm_info["address"] = m.group(1).strip()
                m = re.search(r"농장경영자명\s*:\s*([^,]+)", text)
                if m:
                    farm_info["owner"] = m.group(1).strip()

            for line in text.splitlines():
                m = PDF_ROW_RE.match(line)
                if not m:
                    continue
                rows.append(
                    {
                        "cattle_no": m["no"],
                        "breed": m["breed"],
                        "sex": m["sex"],
                        "birth_date": _format_birth(m["birth"]),
                        "mother_no": m["mother"],
                    }
                )

    return farm_info, rows


def parse_farm_file(path: Path) -> tuple[dict, list[dict]]:
    """확장자 보고 적절한 파서로 디스패치."""
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return parse_farm_pdf(path)
    return parse_farm_excel(path)

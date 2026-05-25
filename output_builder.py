"""사육소 계산 출력 엑셀 생성기.

기존 양식과 동일:
  - 매입 연도별 시트 (예: '2024년 송아지개체')
  - 컬럼: NO. / 개체식별번호 / 성별 / 출생일자 / 매입일 / 매입월령 / 월별 누적 카운터
  - 월별 컬럼은 업로드된 doc_date 의 최소/최대 범위에서 동적 생성
  - 셀 값 = 매입월부터 1, 2, 3 ... (도축/폐사/양수도된 월까지)
"""
from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import date, datetime
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


HEADER_FILL = PatternFill("solid", fgColor="FFE699")
END_STATUSES = {"도축", "폐사", "양수도"}


@dataclass
class CowRow:
    cattle_no: str
    sex: str | None
    birth_date: str | None
    acquisition_date: str | None
    status: str
    status_date: str | None


def _parse_ymd(s: str | None) -> date | None:
    if not s:
        return None
    try:
        return datetime.strptime(s[:10], "%Y-%m-%d").date()
    except ValueError:
        return None


def _format_cattle_no(no: str) -> str:
    """12자리 → '002 1850 5612 1' 처럼 3-4-4-1 공백 포맷."""
    n = "".join(ch for ch in no if ch.isdigit())
    if len(n) == 12:
        return f"{n[:3]} {n[3:7]} {n[7:11]} {n[11]}"
    return no


def _months_diff(a: date, b: date) -> int:
    """a → b 까지의 월 차이 (a 가 더 이전). 1 부터 시작 (a 포함)."""
    return (b.year - a.year) * 12 + (b.month - a.month) + 1


def _generate_month_columns(start: date, end: date) -> list[tuple[int, int, str]]:
    """(year, month, label) 리스트. start ~ end 까지 매월."""
    result = []
    y, m = start.year, start.month
    while (y, m) <= (end.year, end.month):
        if m == 1:
            label = f"{y % 100}년 1월"
        else:
            label = f"{m}월"
        result.append((y, m, label))
        m += 1
        if m == 13:
            m = 1
            y += 1
    return result


def build_workbook(
    cows: Iterable[CowRow],
    *,
    doc_date_range: tuple[date, date],
) -> bytes:
    start, end = doc_date_range
    month_cols = _generate_month_columns(start, end)

    cows_list = list(cows)
    # 매입 연도별 그룹
    by_year: dict[int, list[CowRow]] = {}
    for c in cows_list:
        ad = _parse_ymd(c.acquisition_date)
        year = ad.year if ad else (_parse_ymd(c.birth_date).year if _parse_ymd(c.birth_date) else 0)
        by_year.setdefault(year, []).append(c)

    wb = Workbook()
    wb.remove(wb.active)

    base_headers = ["NO.", "개체식별번호", "성별", "출생일자", "매입일", "매입월령"]
    month_labels = [lbl for _, _, lbl in month_cols]
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="888888")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for year in sorted(by_year.keys(), reverse=True):
        sheet_name = f"{year}년 송아지개체" if year else "매입일 미확인"
        ws = wb.create_sheet(title=sheet_name[:31])

        headers = base_headers + month_labels + ["상태", "사유일자"]
        ws.append(headers)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = HEADER_FILL
            cell.alignment = header_align
            cell.border = border

        rows = sorted(
            by_year[year],
            key=lambda c: (_parse_ymd(c.acquisition_date) or date.min, c.cattle_no),
        )

        for i, c in enumerate(rows, start=1):
            birth = _parse_ymd(c.birth_date)
            acq = _parse_ymd(c.acquisition_date)
            acq_mo = _months_diff(birth, acq) if (birth and acq) else None
            end_dt = _parse_ymd(c.status_date) if c.status in END_STATUSES else None

            row = [
                i,
                _format_cattle_no(c.cattle_no),
                c.sex or "",
                birth,
                acq,
                acq_mo,
            ]

            # 월별 카운터
            counter = 0
            for y, m, _ in month_cols:
                cell_month = date(y, m, 1)
                cell_month_end = date(y + (1 if m == 12 else 0), 1 if m == 12 else m + 1, 1)
                is_acquired = (acq is not None and acq < cell_month_end)
                is_alive = (end_dt is None or cell_month < date(end_dt.year, end_dt.month, 1) or
                            (end_dt.year == y and end_dt.month == m))
                if is_acquired:
                    counter += 1
                    row.append(counter if is_alive else None)
                else:
                    row.append(None)

            row.append(c.status)
            row.append(_parse_ymd(c.status_date))

            ws.append(row)

        for col_idx in range(1, len(headers) + 1):
            ws.cell(row=ws.max_row if ws.max_row else 1, column=col_idx)

        # 컬럼 폭
        ws.column_dimensions["A"].width = 5
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 6
        ws.column_dimensions["D"].width = 12
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 9
        for i, _ in enumerate(month_cols, start=7):
            ws.column_dimensions[get_column_letter(i)].width = 8
        ws.column_dimensions[get_column_letter(7 + len(month_cols))].width = 8
        ws.column_dimensions[get_column_letter(8 + len(month_cols))].width = 12

        ws.freeze_panes = "G2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

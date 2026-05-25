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


def _last_day_of_month(y: int, m: int) -> date:
    if m == 12:
        return date(y + 1, 1, 1).fromordinal(date(y + 1, 1, 1).toordinal() - 1)
    return date(y, m + 1, 1).fromordinal(date(y, m + 1, 1).toordinal() - 1)


def _months_between(a: date, b: date) -> int:
    """a 부터 b 까지 월 차이 (정수). a=birth, b=now → 만 개월령."""
    months = (b.year - a.year) * 12 + (b.month - a.month)
    if b.day < a.day:
        months -= 1
    return months


def _build_summary_sheet(wb: Workbook, cows: list[CowRow], year: int) -> None:
    """월별 통계 요약 시트.

    표1) 입식 통계 (그 월에 매입된 개체, 매입월령 기준)
        - 6~13개월 육성우
        - 5개월 이하 어린소
        - 합계
    표2) 월사육두수 (그 월말 시점 사육 중인 전체 수)
    표3) 사육두수 분석
        - 사육수-5개월이하 = 월사육두수 - 5개월이하
        - 14개월이상 성축
        - 13개월 육성우 (0.5 가중치)
        - 총 사육두수 합계 (가중치 적용)
        - 월 50두 초과두수
    우측 하단: 총사육두수 / 월평균두수 / 비과세 / 기준초과두수
    """
    ws = wb.create_sheet(title=f"{year}년 요약")
    months = list(range(1, 13))
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center")
    yellow = PatternFill("solid", fgColor="FFE699")

    # 각 cow 의 acquisition_date / birth_date / status_end 파싱
    parsed: list[dict] = []
    for c in cows:
        birth = _parse_ymd(c.birth_date)
        acq = _parse_ymd(c.acquisition_date)
        end = _parse_ymd(c.status_date) if c.status in END_STATUSES else None
        if not birth or not acq:
            continue
        parsed.append({
            "birth": birth, "acq": acq, "end": end,
            "acq_mo": _months_diff(birth, acq),
        })

    def alive_in_month(p: dict, y: int, m: int) -> bool:
        month_end = _last_day_of_month(y, m)
        if p["acq"] > month_end:
            return False
        if p["end"] and p["end"] < date(y, m, 1):
            return False
        return True

    # 표1 - 입식 통계
    ws.cell(row=1, column=1, value="").font = header_font
    ws.cell(row=1, column=2, value="1월")
    headers_top = ["구분"] + [f"{m}월" for m in months] + ["합계"]
    for i, h in enumerate(headers_top, start=1):
        cell = ws.cell(row=1, column=i, value=h)
        cell.font = header_font
        cell.alignment = header_align
        cell.fill = yellow

    def write_row(row_idx: int, label: str, values: list[int | float], highlight: bool = False):
        cell = ws.cell(row=row_idx, column=1, value=label)
        cell.font = header_font
        if highlight:
            cell.fill = yellow
        total = 0
        for j, v in enumerate(values, start=2):
            c = ws.cell(row=row_idx, column=j, value=v if v else None)
            if highlight:
                c.fill = yellow
            total += v
        c = ws.cell(row=row_idx, column=2 + len(values), value=total)
        c.font = header_font
        if highlight:
            c.fill = yellow

    # 표1 데이터: 그 월에 매입된 개체를 매입월령으로 분류
    young_in = []  # 5개월 이하
    grow_in = []   # 6~13개월
    for m in months:
        y = year
        c_young = sum(
            1 for p in parsed
            if p["acq"].year == y and p["acq"].month == m and p["acq_mo"] <= 5
        )
        c_grow = sum(
            1 for p in parsed
            if p["acq"].year == y and p["acq"].month == m and 6 <= p["acq_mo"] <= 13
        )
        young_in.append(c_young)
        grow_in.append(c_grow)

    write_row(2, "6-13개월 육성우", grow_in)
    write_row(3, "5개월이하 어린소", young_in)
    write_row(4, "6-13개월 전체합계",
              [a + b for a, b in zip(young_in, grow_in)], highlight=True)

    # 표2 - 월사육두수
    ws.cell(row=6, column=1, value="구분").font = header_font
    ws.cell(row=6, column=1).fill = yellow
    for i, m in enumerate(months, start=2):
        c = ws.cell(row=6, column=i, value=f"{m}월")
        c.font = header_font
        c.fill = yellow
    ws.cell(row=6, column=14, value="합계").font = header_font
    ws.cell(row=6, column=14).fill = yellow

    monthly_total = []
    monthly_under5 = []
    monthly_14plus = []
    monthly_6to13 = []
    for m in months:
        live = [p for p in parsed if alive_in_month(p, year, m)]
        ages = [_months_between(p["birth"], _last_day_of_month(year, m)) for p in live]
        monthly_total.append(len(live))
        monthly_under5.append(sum(1 for a in ages if a <= 5))
        monthly_14plus.append(sum(1 for a in ages if a >= 14))
        monthly_6to13.append(sum(1 for a in ages if 6 <= a <= 13))

    write_row(7, "월사육두수", monthly_total, highlight=True)

    # 표3 - 사육두수 분석
    ws.cell(row=9, column=1, value="구분").font = header_font
    ws.cell(row=9, column=1).fill = yellow
    for i, m in enumerate(months, start=2):
        c = ws.cell(row=9, column=i, value=f"{m}월")
        c.font = header_font
        c.fill = yellow
    ws.cell(row=9, column=14, value="합계").font = header_font
    ws.cell(row=9, column=14).fill = yellow

    write_row(10, "사육수-5개월이하",
              [t - u for t, u in zip(monthly_total, monthly_under5)], highlight=True)
    write_row(11, "14개월이상성축", monthly_14plus)
    write_row(12, "13개월 육성우(0.5)", monthly_6to13)
    weighted = [a + b * 0.5 for a, b in zip(monthly_14plus, monthly_6to13)]
    write_row(13, "총 사육두수 합계", weighted, highlight=True)
    over50 = [max(0, w - 50) for w in weighted]
    write_row(14, "월 50두 초과두수", over50)

    # 우측 하단 요약
    total_sum = sum(weighted)
    avg = total_sum / 12 if total_sum else 0
    ws.cell(row=16, column=13, value="총사육두수").font = header_font
    ws.cell(row=16, column=14, value=round(total_sum, 1))
    ws.cell(row=17, column=13, value="월평균두수").font = header_font
    ws.cell(row=17, column=14, value=round(avg, 1))
    ws.cell(row=18, column=13, value="비과세").font = header_font
    ws.cell(row=18, column=14, value=50)
    ws.cell(row=19, column=13, value="기준초과두수").font = header_font
    ws.cell(row=19, column=13).fill = yellow
    ws.cell(row=19, column=14, value=round(max(0, avg - 50), 1))
    ws.cell(row=19, column=14).fill = yellow

    ws.column_dimensions["A"].width = 18
    for i in range(2, 15):
        ws.column_dimensions[get_column_letter(i)].width = 8


def build_workbook(
    cows: Iterable[CowRow],
    *,
    doc_date_range: tuple[date, date],
) -> bytes:
    start, end = doc_date_range
    month_cols = _generate_month_columns(start, end)

    cows_list = list(cows)
    # 어린소만 시트에 포함 (매입월령 ≤ 13). 성축 매입은 제외.
    # 시트 분리 기준 = 출생 연도.
    by_year: dict[int, list[CowRow]] = {}
    for c in cows_list:
        birth = _parse_ymd(c.birth_date)
        acq = _parse_ymd(c.acquisition_date)
        if not birth or not acq:
            continue
        acq_mo = _months_diff(birth, acq)
        if acq_mo > 13:
            continue  # 성축 매입은 송아지 시트에 안 들어감
        by_year.setdefault(birth.year, []).append(c)

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
            acq_mo = _months_diff(birth, acq)
            end_dt = _parse_ymd(c.status_date) if c.status in END_STATUSES else None

            row = [
                i,
                _format_cattle_no(c.cattle_no),
                c.sex or "",
                birth,
                acq,
                acq_mo,
            ]

            # 월별 카운터: 매입월 = 1, 다음월 = 2, ... 도축/폐사/양수도 된 월까지
            counter = 0
            for y, m, _ in month_cols:
                cell_month = date(y, m, 1)
                cell_month_end = date(y + (1 if m == 12 else 0), 1 if m == 12 else m + 1, 1)
                is_acquired = acq < cell_month_end
                is_alive = (
                    end_dt is None
                    or cell_month < date(end_dt.year, end_dt.month, 1)
                    or (end_dt.year == y and end_dt.month == m)
                )
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

    # 요약 시트 — doc_date 범위의 최신 연도 기준
    _build_summary_sheet(wb, cows_list, end.year)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

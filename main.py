"""안심농장 사육수 계산 메인 파이프라인.

순서:
  1) input/ 폴더의 엑셀 파일 수집 (날짜순)
  2) 각 파일마다:
     - 이미 처리된 (doc_name, farm_name) 조합이면 스킵
     - 농장 정보, 개체 리스트 파싱해 DB upsert
     - 같은 농장의 이전 처리 doc과 비교해 "사라진 개체" 추출
     - 각 개체를 API로 조회해 status/status_date 업데이트
     - processed_docs 마킹
  3) 출력 엑셀에 현재 연도 시트 생성 및 활성 개체 작성
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import db
from api_client import fetch_farm_no_for_unique, fetch_trace
from excel_parser import parse_farm_file, parse_filename
from output_builder import CowRow, build_workbook

PROJECT_ROOT = Path(__file__).parent
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
OUTPUT_FILE = OUTPUT_DIR / "사육소계산(현재사용중).xlsx"


def collect_input_files() -> list[Path]:
    files = []
    candidates = list(INPUT_DIR.glob("*.xls")) + list(INPUT_DIR.glob("*.xlsx")) + list(INPUT_DIR.glob("*.pdf"))
    for p in sorted(candidates):
        meta = parse_filename(p)
        if meta is None:
            print(f"  [skip] 파일명 형식 불일치: {p.name}")
            continue
        files.append(p)
    # doc_date 오름차순 정렬
    files.sort(key=lambda p: parse_filename(p).doc_date)
    return files


def get_cattle_nos_in_doc(farm_name: str, doc_date: str) -> set[str]:
    rows = db.get_cattle_by_farm_doc_date(farm_name, doc_date)
    return {r["cattle_no"] for r in rows}


def process_file(path: Path, *, skip_api: bool = False) -> None:
    meta = parse_filename(path)
    if meta is None:
        return

    if db.is_doc_processed(meta.doc_name, meta.farm_name):
        print(f"[skip] 이미 처리됨: {meta.doc_name}")
        return

    print(f"\n[처리] {meta.doc_name} (농장: {meta.farm_name}, 기준일: {meta.doc_date})")

    farm_info, cattle_rows = parse_farm_file(path)
    farm_name = farm_info["farm_name"] or meta.farm_name
    db.upsert_farm(farm_name, farm_info["farm_id"], farm_info["owner"], farm_info["address"])

    for r in cattle_rows:
        r["farm_name"] = farm_name

    # 농장의 farm_no 학습 (API)
    our_farm_no = None
    if not skip_api and farm_info["farm_id"] and cattle_rows:
        f = db.get_farm(farm_name)
        if f and f["farm_no"]:
            our_farm_no = f["farm_no"]
        else:
            our_farm_no = fetch_farm_no_for_unique(cattle_rows[0]["cattle_no"], farm_info["farm_id"])
            if our_farm_no:
                db.set_farm_no(farm_name, our_farm_no)

    prev_doc_date = db.get_prev_doc_date(farm_name, meta.doc_date)
    inserted, updated = db.upsert_cattle_batch(cattle_rows, meta.doc_date)
    print(f"  개체: 신규 {inserted}, 갱신 {updated} (총 {len(cattle_rows)})")

    if prev_doc_date:
        prev_set = get_cattle_nos_in_doc(farm_name, prev_doc_date)
        curr_set = {r["cattle_no"] for r in cattle_rows}
        missing = prev_set - curr_set
        print(f"  전월({prev_doc_date}) 대비 사라진 개체: {len(missing)}")
        for no in sorted(missing):
            if skip_api:
                print(f"    [dry-run] {no}")
                continue
            info = fetch_trace(no, our_farm_no=our_farm_no)
            db.update_cattle_status(no, info.status, info.status_date)
            if info.acquisition_date:
                db.update_cattle_acquisition(no, info.acquisition_date)
            print(f"    {no} → {info.status} ({info.status_date or '-'})")
    else:
        print("  (이전 처리 기록 없음 — 비교 생략)")

    db.mark_doc_processed(meta.doc_name, meta.farm_name, meta.doc_date, len(cattle_rows))


def write_output_excel() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)
    with db.connect() as conn:
        all_cattle = conn.execute(
            "SELECT * FROM cattle ORDER BY acquisition_date, cattle_no"
        ).fetchall()
        dates = conn.execute(
            "SELECT MIN(doc_date) AS lo, MAX(doc_date) AS hi FROM processed_docs"
        ).fetchone()

    if dates and dates["lo"] and dates["hi"]:
        lo = datetime.strptime(dates["lo"], "%Y-%m-%d").date()
        hi = datetime.strptime(dates["hi"], "%Y-%m-%d").date()
    else:
        today = datetime.now().date()
        lo = hi = today

    cow_rows = [
        CowRow(
            cattle_no=r["cattle_no"],
            sex=r["sex"],
            birth_date=r["birth_date"],
            acquisition_date=r["acquisition_date"] or r["first_seen_doc_date"],
            status=r["status"],
            status_date=r["status_date"],
        )
        for r in all_cattle
    ]
    out_path = OUTPUT_DIR / "사육소 계산(현재사용중인표).xlsx"
    monthly_totals = db.get_monthly_total_counts(hi.year)
    out_path.write_bytes(build_workbook(cow_rows, doc_date_range=(lo, hi), monthly_totals=monthly_totals))
    print(f"\n[저장] {out_path} ({len(cow_rows)}건)")


def main(skip_api: bool = False) -> None:
    db.init_db()
    files = collect_input_files()
    print(f"input 파일 {len(files)}개")
    for p in files:
        process_file(p, skip_api=skip_api)
    write_output_excel()
    print("\n완료.")


if __name__ == "__main__":
    skip_api = "--skip-api" in sys.argv
    main(skip_api=skip_api)

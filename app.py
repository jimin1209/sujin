"""안심농장 사육수 계산 — Streamlit 웹 UI.

비개발자용 화면:
  1) 엑셀 파일 여러 개 업로드 (월별 누적)
  2) [처리 시작] 클릭
  3) 결과 표 확인 + 출력 엑셀 다운로드
"""
from __future__ import annotations

import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook

# Streamlit Cloud 의 secrets 에 API 키가 있으면 환경변수로 주입 (없으면 .env 사용)
try:
    if "LIVESTOCK_API_KEY" in st.secrets:
        os.environ["LIVESTOCK_API_KEY"] = st.secrets["LIVESTOCK_API_KEY"]
except Exception:
    pass

import db  # noqa: E402
from api_client import fetch_status  # noqa: E402
from excel_parser import parse_farm_excel, parse_filename  # noqa: E402

st.set_page_config(page_title="수진쨩노 안심농장 사육수 계산", page_icon="🐄", layout="wide")
st.title("🐄 수진쨩노 안심농장 사육수 계산")
st.caption("월별 농장 엑셀을 올리면 누적 집계 후 현재 사육 중인 개체 목록을 엑셀로 받을 수 있어요.")


def _ensure_session_db() -> None:
    if "db_path" not in st.session_state:
        tmp = tempfile.NamedTemporaryFile(prefix="sujin_", suffix=".db", delete=False)
        tmp.close()
        st.session_state.db_path = tmp.name
    db.set_db_path(st.session_state.db_path)
    db.init_db()


def _reset_session_db() -> None:
    if "db_path" in st.session_state:
        try:
            Path(st.session_state.db_path).unlink(missing_ok=True)
        except Exception:
            pass
        del st.session_state.db_path


def _save_uploaded(upload) -> Path:
    """업로드된 파일을 임시 경로에 저장 (xls는 파일 경로로 읽어야 해서)."""
    tmpdir = Path(tempfile.gettempdir()) / "sujin_uploads"
    tmpdir.mkdir(exist_ok=True)
    path = tmpdir / upload.name
    path.write_bytes(upload.getbuffer())
    return path


def _build_output_xlsx(rows: list) -> bytes:
    year = datetime.now().year
    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}년"
    headers = ["농장명", "개체식별번호", "소의종류", "성별", "출생일자", "모 개체번호",
               "최초확인일", "최종확인일", "상태"]
    ws.append(headers)
    for r in rows:
        ws.append([
            r["farm_name"], r["cattle_no"], r["breed"], r["sex"],
            r["birth_date"], r["mother_no"],
            r["first_seen_doc_date"], r["last_seen_doc_date"], r["status"],
        ])
    widths = [10, 18, 10, 6, 12, 18, 12, 12, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ────────────────────────────────────────────────────────────────────
# Sidebar
# ────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 옵션")
    use_api = st.toggle(
        "이력조회 API 호출",
        value=True,
        help="끄면 사라진 개체의 상태(도축 등) 갱신을 생략합니다. 빠른 테스트용.",
    )
    if not os.environ.get("LIVESTOCK_API_KEY"):
        st.warning("API 키가 설정되지 않았습니다. 관리자에게 문의하세요.", icon="⚠️")

    st.divider()
    if st.button("🔄 세션 초기화 (업로드 기록 삭제)", use_container_width=True):
        _reset_session_db()
        st.success("초기화 완료. 페이지를 새로고침 하세요.")


# ────────────────────────────────────────────────────────────────────
# 메인
# ────────────────────────────────────────────────────────────────────
_ensure_session_db()

st.subheader("1. 농장 엑셀 업로드")
st.markdown(
    "- 파일명 형식: `1농장-25.01.31기준.xls` (농장명-YY.MM.DD기준.xls)\n"
    "- 여러 개 한 번에 올려도 됩니다. 같은 농장의 여러 월을 함께 올리면 자동으로 비교합니다."
)
uploads = st.file_uploader(
    "xls 파일 선택",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if uploads:
    rows_preview = []
    for u in uploads:
        meta = parse_filename(Path(u.name))
        rows_preview.append({
            "파일명": u.name,
            "농장": meta.farm_name if meta else "❌ 형식 오류",
            "기준일": meta.doc_date if meta else "-",
        })
    st.dataframe(pd.DataFrame(rows_preview), use_container_width=True, hide_index=True)

    if st.button("▶️ 처리 시작", type="primary", use_container_width=True):
        # doc_date 순으로 정렬 (이전월 비교가 가능하도록)
        ordered = []
        for u in uploads:
            meta = parse_filename(Path(u.name))
            if meta is None:
                st.warning(f"건너뜀(파일명 형식 오류): {u.name}")
                continue
            ordered.append((meta, u))
        ordered.sort(key=lambda x: x[0].doc_date)

        progress = st.progress(0.0, text="시작...")
        log_box = st.expander("📋 처리 로그", expanded=False)
        log_lines: list[str] = []

        def log(msg: str) -> None:
            log_lines.append(msg)
            log_box.code("\n".join(log_lines[-200:]), language=None)

        api_calls = 0
        api_done = 0

        # 사라진 개체 수 사전 계산 (진행률 표시용)
        for i, (meta, u) in enumerate(ordered):
            if db.is_doc_processed(meta.doc_name, meta.farm_name):
                continue
            prev = db.get_prev_doc_date(meta.farm_name, meta.doc_date)
            if prev:
                prev_set = {r["cattle_no"] for r in db.get_cattle_by_farm_doc_date(meta.farm_name, prev)}
                # 임시 파싱해서 사라진 수 산출
                path = _save_uploaded(u)
                _, cattle_rows = parse_farm_excel(path)
                curr_set = {r["cattle_no"] for r in cattle_rows}
                api_calls += len(prev_set - curr_set) if use_api else 0

        total_steps = len(ordered) + api_calls
        step = 0

        for idx, (meta, u) in enumerate(ordered):
            if db.is_doc_processed(meta.doc_name, meta.farm_name):
                log(f"[skip] 이미 처리됨: {meta.doc_name}")
                step += 1
                progress.progress(min(step / max(total_steps, 1), 1.0),
                                  text=f"({idx+1}/{len(ordered)}) {meta.doc_name}")
                continue

            path = _save_uploaded(u)
            log(f"[처리] {meta.doc_name}")

            farm_info, cattle_rows = parse_farm_excel(path)
            farm_name = farm_info["farm_name"] or meta.farm_name
            db.upsert_farm(farm_name, farm_info["farm_id"], farm_info["owner"], farm_info["address"])
            for r in cattle_rows:
                r["farm_name"] = farm_name

            prev_doc_date = db.get_prev_doc_date(farm_name, meta.doc_date)
            inserted, updated = db.upsert_cattle_batch(cattle_rows, meta.doc_date)
            log(f"  개체: 신규 {inserted}, 갱신 {updated} (총 {len(cattle_rows)})")

            if prev_doc_date:
                prev_set = {r["cattle_no"] for r in db.get_cattle_by_farm_doc_date(farm_name, prev_doc_date)}
                curr_set = {r["cattle_no"] for r in cattle_rows}
                missing = sorted(prev_set - curr_set)
                log(f"  전월({prev_doc_date}) 대비 사라진 개체: {len(missing)}")
                if use_api:
                    for no in missing:
                        info = fetch_status(no)
                        db.update_cattle_status(no, info.status, info.status_date)
                        api_done += 1
                        step += 1
                        progress.progress(min(step / max(total_steps, 1), 1.0),
                                          text=f"API 조회 {api_done}/{api_calls}")
                        log(f"    {no} → {info.status} ({info.status_date or '-'})")
            else:
                log("  (이전 처리 기록 없음 — 비교 생략)")

            db.mark_doc_processed(meta.doc_name, meta.farm_name, meta.doc_date)
            step += 1
            progress.progress(min(step / max(total_steps, 1), 1.0),
                              text=f"({idx+1}/{len(ordered)}) 완료")

        progress.progress(1.0, text="완료")
        st.success("처리 완료")
        st.session_state["processed"] = True


# ────────────────────────────────────────────────────────────────────
# 결과
# ────────────────────────────────────────────────────────────────────
if st.session_state.get("processed"):
    st.divider()
    st.subheader("2. 결과")

    active = db.get_active_cattle_all()
    active_df = pd.DataFrame([dict(r) for r in active])

    col1, col2, col3 = st.columns(3)
    col1.metric("활성 개체 (사육 중)", len(active))
    col2.metric("농장 수", active_df["farm_name"].nunique() if not active_df.empty else 0)

    # 농장별 통계
    if not active_df.empty:
        by_farm = active_df.groupby("farm_name").size().reset_index(name="활성 개체수")
        col3.metric("최대 농장", f"{by_farm.iloc[by_farm['활성 개체수'].idxmax()]['farm_name']}")

        st.markdown("**농장별 활성 개체**")
        st.dataframe(by_farm, use_container_width=True, hide_index=True)

    if not active_df.empty:
        st.markdown("**전체 활성 개체 목록**")
        display = active_df[[
            "farm_name", "cattle_no", "breed", "sex", "birth_date",
            "first_seen_doc_date", "last_seen_doc_date", "status",
        ]].rename(columns={
            "farm_name": "농장",
            "cattle_no": "개체식별번호",
            "breed": "소의종류",
            "sex": "성별",
            "birth_date": "출생일자",
            "first_seen_doc_date": "최초확인일",
            "last_seen_doc_date": "최종확인일",
            "status": "상태",
        })
        st.dataframe(display, use_container_width=True, hide_index=True, height=400)

    xlsx_bytes = _build_output_xlsx(active)
    st.download_button(
        "📥 사육소계산(현재사용중).xlsx 다운로드",
        data=xlsx_bytes,
        file_name="사육소계산(현재사용중).xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

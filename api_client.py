"""축산물품질평가원 축산물통합이력정보 API 클라이언트.

문서: https://www.data.go.kr/data/15058923/openapi.do
엔드포인트: /openapi-data/service/user/animalTrace/traceNoSearch
optionNo=2 (사육지 이력) 한 번 호출로 도축/폐사/양수도 모두 판별.
"""
from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Optional

import requests
from dotenv import load_dotenv

load_dotenv()

API_KEY = os.getenv("LIVESTOCK_API_KEY", "")
BASE_URL = "http://data.ekape.or.kr/openapi-data/service/user/animalTrace/traceNoSearch"


@dataclass
class TraceInfo:
    status: str  # 사육 / 도축 / 폐사 / 양수도 / 미조회
    status_date: Optional[str]


def _norm_date(raw: Optional[str]) -> Optional[str]:
    if not raw:
        return None
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) >= 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return raw


def fetch_status(cattle_no: str) -> TraceInfo:
    """개체식별번호의 최종 상태를 1회 API 호출로 판정.

    optionNo=2 의 사육지 이력 마지막 row 의 regType:
      - 도축출하  → 도축
      - 폐사     → 폐사
      - 양수/양도 → 양수도
      - 전산등록  → 사육
    """
    if not API_KEY:
        raise RuntimeError("LIVESTOCK_API_KEY 환경변수가 비어있습니다 (.env / Streamlit Secrets 확인)")

    try:
        r = requests.get(
            BASE_URL,
            params={"serviceKey": API_KEY, "traceNo": cattle_no, "optionNo": "2"},
            timeout=15,
        )
        r.raise_for_status()
    except requests.RequestException:
        return TraceInfo(status="미조회", status_date=None)

    try:
        root = ET.fromstring(r.text)
    except ET.ParseError:
        return TraceInfo(status="미조회", status_date=None)

    result_code = root.findtext(".//resultCode")
    if result_code and result_code != "00":
        return TraceInfo(status="미조회", status_date=None)

    items = root.findall(".//item")
    if not items:
        return TraceInfo(status="미조회", status_date=None)

    latest = max(items, key=lambda it: it.findtext("regYmd") or "")
    reg_type = (latest.findtext("regType") or "").strip()
    date = _norm_date(latest.findtext("regYmd"))

    if reg_type == "도축출하":
        return TraceInfo(status="도축", status_date=date)
    if reg_type == "폐사":
        return TraceInfo(status="폐사", status_date=date)
    if reg_type in ("양수", "양도"):
        return TraceInfo(status="양수도", status_date=date)
    return TraceInfo(status="사육", status_date=None)

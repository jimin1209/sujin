"""축산물품질평가원 축산물통합이력정보 API 클라이언트.

문서: https://www.data.go.kr/data/15058923/openapi.do
엔드포인트: /openapi-data/service/user/animalTrace/traceNoSearch

옵션:
  optionNo=1 → 개체정보 (farmNo, farmUniqueNo 등)
  optionNo=2 → 사육지 이력 (regYmd, regType, farmNo, farmerNm)
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
    acquisition_date: Optional[str] = None  # 우리 농장으로 들어온 날 (양수/전산등록)


def _norm_date(raw: Optional[str]) -> Optional[str]:
    if not raw:
        return None
    digits = "".join(ch for ch in raw if ch.isdigit())
    if len(digits) >= 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return raw


def _request(cattle_no: str, option_no: str) -> Optional[ET.Element]:
    if not API_KEY:
        raise RuntimeError("LIVESTOCK_API_KEY 환경변수가 비어있습니다 (.env / Streamlit Secrets 확인)")
    try:
        r = requests.get(
            BASE_URL,
            params={"serviceKey": API_KEY, "traceNo": cattle_no, "optionNo": option_no},
            timeout=15,
        )
        r.raise_for_status()
    except requests.RequestException:
        return None
    try:
        root = ET.fromstring(r.text)
    except ET.ParseError:
        return None
    code = root.findtext(".//resultCode")
    if code and code != "00":
        return None
    return root


def fetch_farm_no_for_unique(cattle_no: str, farm_unique_no: str) -> Optional[str]:
    """주어진 개체의 optionNo=1 으로 farm_unique_no 와 매칭되는 farmNo 를 찾음.

    같은 농장식별번호(farmUniqueNo) 를 농장번호(farmNo) 로 매핑.
    """
    root = _request(cattle_no, "1")
    if root is None:
        return None
    for it in root.findall(".//item"):
        if (it.findtext("farmUniqueNo") or "").strip() == str(farm_unique_no):
            return (it.findtext("farmNo") or "").strip() or None
    return None


def fetch_trace(cattle_no: str, our_farm_no: Optional[str] = None) -> TraceInfo:
    """사육지 이력으로 최종 상태 + 우리 농장 매입일 한 번에 판정.

    our_farm_no 가 주어지면, 그 farm_no 의 가장 빠른 (전산등록|양수) 행 = 매입일.
    """
    root = _request(cattle_no, "2")
    if root is None:
        return TraceInfo(status="미조회", status_date=None)

    items = root.findall(".//item")
    if not items:
        return TraceInfo(status="미조회", status_date=None)

    # 1) 최종 상태 판정
    latest = max(items, key=lambda it: it.findtext("regYmd") or "")
    reg_type = (latest.findtext("regType") or "").strip()
    status_date = _norm_date(latest.findtext("regYmd"))

    if reg_type == "도축출하":
        status, sdate = "도축", status_date
    elif reg_type == "폐사":
        status, sdate = "폐사", status_date
    elif reg_type in ("양수", "양도"):
        # 우리 농장으로의 양수면 사육 (아직 우리 농장에 있음)
        latest_farm = (latest.findtext("farmNo") or "").strip()
        if our_farm_no and latest_farm == str(our_farm_no):
            status, sdate = "사육", None
        else:
            status, sdate = "양수도", status_date
    else:
        status, sdate = "사육", None

    # 2) 매입일 — 우리 농장 farmNo 와 일치하는 가장 빠른 row 의 regYmd
    acquisition = None
    if our_farm_no:
        ours = sorted(
            [
                it for it in items
                if (it.findtext("farmNo") or "").strip() == str(our_farm_no)
                and (it.findtext("regType") or "") in ("전산등록", "양수")
            ],
            key=lambda it: it.findtext("regYmd") or "",
        )
        if ours:
            acquisition = _norm_date(ours[0].findtext("regYmd"))

    return TraceInfo(status=status, status_date=sdate, acquisition_date=acquisition)


# 하위 호환 alias
def fetch_status(cattle_no: str) -> TraceInfo:
    return fetch_trace(cattle_no)

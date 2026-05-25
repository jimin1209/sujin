"""축산물품질평가원 쇠고기 이력정보 API 클라이언트.

문서: https://www.data.go.kr/data/15056898/openapi.do
엔드포인트: /openapi-data/service/user/mtrace/breeding/cattle
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
BASE_URL = "http://data.ekape.or.kr/openapi-data/service/user/mtrace/breeding/cattle"


@dataclass
class TraceInfo:
    status: str  # 사육 / 도축 / 폐사 / 미조회
    status_date: Optional[str]


def _first_text(root: ET.Element, tag: str) -> Optional[str]:
    el = root.find(f".//{tag}")
    if el is None or el.text is None:
        return None
    val = el.text.strip()
    return val or None


def fetch_status(cattle_no: str) -> TraceInfo:
    """주어진 개체식별번호의 최종 상태를 조회.

    butcheryYmd가 있으면 도축, 없으면 사육 처리. API 호출 실패 시 '미조회'.
    """
    if not API_KEY:
        raise RuntimeError("LIVESTOCK_API_KEY 환경변수가 비어있습니다 (.env 확인)")

    try:
        r = requests.get(
            BASE_URL,
            params={"serviceKey": API_KEY, "cattleNo": cattle_no},
            timeout=15,
        )
        r.raise_for_status()
    except requests.RequestException:
        return TraceInfo(status="미조회", status_date=None)

    try:
        root = ET.fromstring(r.text)
    except ET.ParseError:
        return TraceInfo(status="미조회", status_date=None)

    result_code = _first_text(root, "resultCode")
    if result_code and result_code != "00":
        return TraceInfo(status="미조회", status_date=None)

    butchery_ymd = _first_text(root, "butcheryYmd")
    if butchery_ymd:
        return TraceInfo(status="도축", status_date=butchery_ymd)

    return TraceInfo(status="사육", status_date=None)

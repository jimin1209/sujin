"""쇠고기 이력 API 응답 영구 캐시.

cache/cattle_cache.json 에 {cattle_no: {status, status_date}} 저장.
로컬에서 쓰면 파일에만 저장. 웹에서는 commit_to_github() 으로 깃에 푸시.
"""
from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any, Optional

import requests

CACHE_REPO_PATH = "cache/cattle_cache.json"  # 깃 레포 기준 상대경로
CACHE_LOCAL_PATH = Path(__file__).parent / CACHE_REPO_PATH


def load() -> dict[str, dict[str, Any]]:
    if not CACHE_LOCAL_PATH.exists():
        return {}
    try:
        return json.loads(CACHE_LOCAL_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_local(data: dict[str, dict[str, Any]]) -> None:
    CACHE_LOCAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    text = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True)
    CACHE_LOCAL_PATH.write_text(text + "\n", encoding="utf-8")


def commit_to_github(
    data: dict[str, dict[str, Any]],
    *,
    repo: str,
    token: str,
    branch: str = "main",
    path: str = CACHE_REPO_PATH,
    message: Optional[str] = None,
) -> dict[str, Any]:
    """GitHub Contents API 로 캐시 파일을 커밋/업데이트.

    repo 예: 'jimin1209/sujin'
    token: repo Contents Read/Write 권한이 있는 fine-grained PAT
    """
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    url = f"https://api.github.com/repos/{repo}/contents/{path}"

    # 기존 파일 sha 조회 (없으면 신규)
    r = requests.get(url, params={"ref": branch}, headers=headers, timeout=20)
    sha = r.json().get("sha") if r.status_code == 200 else None

    content_str = json.dumps(data, indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    content_b64 = base64.b64encode(content_str.encode("utf-8")).decode("ascii")

    body: dict[str, Any] = {
        "message": message or f"chore: update cattle cache ({len(data)} entries)",
        "content": content_b64,
        "branch": branch,
    }
    if sha:
        body["sha"] = sha

    r = requests.put(url, json=body, headers=headers, timeout=30)
    r.raise_for_status()
    return r.json()

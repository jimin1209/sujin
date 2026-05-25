# 안심농장 사육수 계산

월별 농장 엑셀(예: `1농장-25.01.31기준.xls`)을 누적해서, 현재 사육 중인 개체 목록을 엑셀로 받는 도구입니다.
사라진 개체는 축산물품질평가원 쇠고기 이력 API로 도축 여부를 자동 확인합니다.

## 사용 방법 (웹)

비개발자: Streamlit Cloud에 올린 URL을 열고 파일을 업로드 → 다운로드 받으면 끝.

## 로컬에서 실행 (개발자)

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
cp .env.example .env  # 그리고 LIVESTOCK_API_KEY 채우기
streamlit run app.py
```

브라우저가 자동으로 열리고 `http://localhost:8501` 에서 사용 가능.

## Streamlit Cloud 배포

1. 이 코드를 GitHub 레포(예: `jimin1209/sujin`)에 push
2. https://share.streamlit.io 접속 → GitHub 로그인
3. **New app** → 레포 / 브랜치 / `app.py` 선택
4. **Advanced settings → Secrets** 에 아래 추가:
   ```toml
   LIVESTOCK_API_KEY = "여기에 공공데이터포털 서비스키"
   ```
5. **Deploy** → 고정 URL 발급. 비개발자에게 이 URL만 공유.

## 파일 구조

```
.
├── app.py              # Streamlit 웹 UI
├── main.py             # CLI 파이프라인 (배치 실행용)
├── db.py               # SQLite 관리
├── excel_parser.py     # 농장 엑셀 파서
├── api_client.py       # 이력조회 API
├── requirements.txt
├── input/              # (CLI 모드용) 농장 엑셀 위치
└── output/             # (CLI 모드용) 출력 엑셀
```

## CLI 모드

```bash
# input/ 에 *.xls 파일을 두고
python3 main.py            # API 호출 포함
python3 main.py --skip-api # API 호출 없이 dry-run
```

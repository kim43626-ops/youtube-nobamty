# 유튜브 노밤티 분석 보고서 배포 안내

이 폴더는 Streamlit Community Cloud 배포용입니다. `data/youtube_data.xlsx`에 기본 엑셀 데이터가 포함되어 앱 실행 시 자동으로 불러옵니다. 다른 파일을 사용하려면 앱에서 업로드할 수 있습니다. 자동 분석은 엑셀의 `1. 신규` 시트 구조를 기준으로 합니다. API 키, 로컬 데이터베이스, 배치파일과 테스트 파일은 포함하지 않았습니다.

## 배포

1. 이 폴더의 파일을 GitHub 저장소의 **최상위**에 올립니다. `packages.txt`는 저장소 최상위에 있어야 한글 글꼴을 설치합니다. 저장소를 만들고 업로드하는 작업은 이 폴더 생성에 포함되어 있지 않습니다.
2. https://share.streamlit.io/ 에 로그인하고 새 앱에서 저장소와 브랜치를 선택합니다.
3. Main file path에 `main_app.py`를 입력합니다. Python 버전은 로컬과 맞는 3.12를 선택합니다.
4. 키를 서버에서 관리할 경우 Advanced settings > Secrets에 아래 형식으로 입력합니다. 실제 키는 GitHub 파일이나 이 설명서에 적지 마세요.

```toml
YOUTUBE_API_KEY = "발급받은_API_키"
```

5. 앱의 공개·비공개 접근 설정을 확인하고 배포합니다. 키가 설정되지 않으면 엑셀 분석만 동작합니다. 앱 방문자의 YouTube 조회도 서버에 설정된 키의 할당량을 사용합니다.

## 보안 및 데이터

- `.env`, `.streamlit/secrets.toml`, API 키는 GitHub에 올리지 마세요. `.gitignore`가 일부 파일을 제외하지만 이미 추적된 파일에는 적용되지 않으므로 업로드 전에 확인해야 합니다.
- `data/youtube_data.xlsx`는 요청에 따라 저장소에 포함됩니다. 저장소 접근자와 배포된 앱의 허용된 방문자는 엑셀 데이터 또는 집계 결과에 접근할 수 있으므로 회사 데이터의 공유 범위를 확인하고 필요한 경우 비공개 저장소와 비공개 앱을 사용하세요.
- Streamlit Community Cloud의 로컬 파일은 영구 보관이 보장되지 않습니다. 이 배포본은 데이터를 서버에 영구 보관하는 용도가 아닙니다.
- 공개 YouTube API 검색은 전체 인기 순위나 실제 검색량이 아닙니다. API 키는 화면에서 가려지며 보고서에는 기록하지 않습니다.
- 보고서는 Windows의 맑은 고딕 대신 Linux의 나눔고딕을 사용합니다. Word 주간·월간 보고서는 디자인 유지를 위해 페이지 이미지로 구성됩니다.

## 로컬 테스트

Python 3.10 이상에서 `pip install -r requirements.txt` 후 `streamlit run main_app.py`를 실행합니다. Windows의 경우 설치된 맑은 고딕을 사용합니다. Linux에서는 `packages.txt`의 `fonts-nanum`을 설치해야 합니다.

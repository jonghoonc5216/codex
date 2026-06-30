# L-사전규격

나라장터 사전규격공개 자동 수집 도구입니다.

## 빠른 시작

1. `config.json`을 열고 `api_key`에 공공데이터포털 인증키를 입력합니다.
2. `run_g2b_prespec_update.bat`을 실행합니다.
3. 생성된 `조달청_사전규격_자동수집.xlsx` 파일을 확인합니다.
4. 바탕화면 바로가기/로그온 자동 실행을 등록하려면 PowerShell에서 `install_startup_and_shortcut.ps1`을 실행합니다.

## 포함 내용

- `g2b_prespec_updater.py`: 나라장터 사전규격 API 수집 및 엑셀 생성
- `config.json`: API 키가 비어 있는 설정 템플릿
- `run_g2b_prespec_update.bat`: 수집 실행 배치파일
- `edit_config.bat`: 설정 파일 편집 배치파일
- `install_startup_and_shortcut.ps1`: 바탕화면 바로가기 및 로그온 자동 실행 등록
- `g2b_prespec_deeplink_extension/`: 브라우저 상세조회 보조 확장 파일

## 공유 제외 파일

개인 API 키, 수집 기록, 상태 파일, 로그, 생성된 엑셀 파일, Chrome 프로필은 GitHub에 올리지 않습니다.

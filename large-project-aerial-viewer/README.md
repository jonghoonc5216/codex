# 조감도작성

DXF 기반 조감도 작성, 경사분석, Google Earth 및 Blender 연동 프로그램입니다.

좌표가 들어간 DXF 도면을 불러와 3D로 확인하고, 간단한 선형/해치 편집을 한 뒤 Google Earth Pro에서 실제 현장 위치에 대입할 수 있습니다. Blender가 설치되어 있으면 현재 도면을 지형 메쉬, CAD 선형, 해치면, 이미지 평면으로 변환해 Blender에서 바로 열 수 있습니다.

## 빠른 실행

### exe로 실행

GitHub에서 `large-project-aerial-viewer/dist/조감도작성.exe`를 내려받아 실행합니다.

실행하면 로컬 서버가 켜지고 기본 브라우저에서 다음 주소가 열립니다.

```text
http://127.0.0.1:8765
```

### Python으로 실행

```powershell
cd C:\Users\saman\Documents\Codex\large-project-aerial-viewer
python launch_viewer.py
```

또는 서버만 직접 실행할 수 있습니다.

```powershell
python server.py --host 127.0.0.1 --port 8765
```

## exe 다시 만들기

노트북에서 최신 코드로 exe를 다시 만들려면 다음 파일을 실행합니다.

```text
조감도작성_EXE_만들기.bat
```

완료되면 다음 파일이 생성됩니다.

```text
dist\조감도작성.exe
```

GitHub Actions에서도 `Build 조감도작성 Windows EXE` 워크플로를 실행하면 `조감도작성-windows-exe` 아티팩트로 exe를 받을 수 있습니다.

## 주요 기능

- DXF 열기
- 빈 화면에서 시작
- GRS80 계열 좌표계 자동 감지
- DXF 선, 해치, 문자 표시
- DXF 이미지 표시 및 Google Earth 반영
- 현재 작업 도면의 `합본1.jpg` 항공사진 보강 표시
- 마우스 3D 이동, 회전, 기울기, 확대/축소
- 간단 편집 모드
- 선, 면, 해치 선택 및 이동
- 해치 색상 및 투명도 변경
- 작업 저장 및 불러오기
- 수정 DXF 내보내기
- Google Earth Pro로 바로 열기
- 경사분석
- Blender로 현재 도면 열기

## 경사분석

DXF 등고선 Z값을 기반으로 경사도를 계산합니다.

기본 구간:

```text
0-5, 6-10, 11-16, 17, 18-25, 26-30, 31도 이상
```

셀 크기는 기본 2m이고 1m까지 설정할 수 있습니다. 큰 도면은 앱이 멈추지 않도록 자동으로 셀 크기를 키워 계산합니다. 경사분석 결과는 화면에 색상면으로 표시되고, Google Earth로 열 때도 폴리곤으로 포함됩니다.

## Google Earth 연동

`Google Earth로 열기` 버튼을 누르면 현재 도면 상태를 KMZ로 만들고 설치된 Google Earth Pro에서 바로 엽니다.

Google Earth Pro 기본 탐색 경로:

```text
C:\Program Files\Google\Google Earth Pro\client\googleearth.exe
C:\Program Files (x86)\Google\Google Earth Pro\client\googleearth.exe
```

Google Earth에 기존 KMZ가 이미 열려 있으면 자동 갱신되지 않을 수 있습니다. 기존 항목을 삭제하거나 체크 해제한 뒤 다시 열어 주세요.

## Blender 연동

`Blender로 열기` 버튼을 누르면 현재 도면을 Blender용 3D 장면으로 변환합니다.

현재 기본 탐색 경로:

```text
D:\Ai 프로그래밍\Blender\blender.exe
C:\Program Files\Blender Foundation\Blender\blender.exe
```

조감도작성에서 좌표계를 해석하고, Blender에는 도면 중심 기준 상대좌표로 넘깁니다. 그래서 Blender 장면 안에서는 지형, 선형, 해치, 이미지가 서로 맞게 열립니다. 원래 좌표계 정보는 Blender 스크립트 안에 EPSG, 원점 좌표, WGS84 중심값으로 함께 저장됩니다.

Blender 자체 GIS 플러그인은 필수는 아닙니다. 필요하면 나중에 BlenderGIS 같은 플러그인을 추가해 고급 GIS/DEM 작업을 확장할 수 있습니다.

## 주의사항

- DWG는 직접 열기보다 DXF로 변환해서 사용합니다.
- 경사분석 정확도는 DXF 등고선 Z값 품질에 영향을 받습니다.
- Google Earth 연동은 Google Earth Pro 설치가 필요합니다.
- Blender 연동은 Blender 설치가 필요합니다.
- 참조 이미지는 실제 파일이 존재해야 화면과 Google Earth/Blender에 표시됩니다.
- `cache`, `exports`, `data/uploads`, `__pycache__`, 로그 파일은 GitHub 업로드에서 제외합니다.

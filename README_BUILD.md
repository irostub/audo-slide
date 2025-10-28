# AutoSlide EXE 빌드 가이드

## 필수 요구사항

- Python 3.8 이상
- 가상환경 활성화
- 의존성 패키지 설치

## 빌드 방법

### 1. 가상환경 활성화 및 의존성 설치

```bash
# 가상환경 활성화
source .venv/bin/activate  # macOS/Linux
# 또는
.venv\Scripts\activate     # Windows

# PyInstaller 설치 (아직 안 했다면)
pip install pyinstaller
```

### 2. EXE 빌드 실행

```bash
# 방법 1: spec 파일 사용 (권장)
pyinstaller autoslide.spec

# 방법 2: 직접 명령어
pyinstaller --onefile --console --name AutoSlide main.py
```

### 3. 빌드 결과 확인

```
dist/
  └── AutoSlide          # macOS/Linux
  └── AutoSlide.exe      # Windows
```

## 사용 방법

### 실행 전 준비
1. `template.pptx` 파일을 exe와 같은 폴더에 배치
2. `data.xlsx` 파일을 exe와 같은 폴더에 배치

### 실행
```bash
# macOS/Linux
./dist/AutoSlide

# Windows
dist\AutoSlide.exe
```

또는 더블클릭으로 실행

### 결과
- `output.pptx` 파일이 exe와 같은 폴더에 생성됨

## 설정 변경

`main.py`의 상단 설정값을 수정하여 파일명 변경 가능:

```python
TEMPLATE_FILE = 'template.pptx'  # 템플릿 PPT 파일명
EXCEL_FILE = 'data.xlsx'          # 엑셀 데이터 파일명
OUTPUT_FILE = 'output.pptx'       # 출력 PPT 파일명
PREFERRED_SHEET = 'export'         # 우선 사용할 시트명
```

## 문제 해결

### 빌드 오류 시
```bash
# 캐시 정리 후 재빌드
rm -rf build dist __pycache__
pyinstaller --clean autoslide.spec
```

### 실행 오류 시
- 콘솔 창에서 실행하여 오류 메시지 확인
- `template.pptx`와 `data.xlsx` 파일이 올바른 위치에 있는지 확인
- 엑셀 파일이 다른 프로그램에서 열려있지 않은지 확인

## 배포 시 주의사항

배포용 폴더 구조:
```
AutoSlide/
  ├── AutoSlide.exe      # 실행 파일
  ├── template.pptx      # 템플릿 (예제)
  ├── data.xlsx          # 엑셀 데이터 (예제)
  └── README.txt         # 사용 설명서
```

## 고급 설정

### GUI 모드로 빌드 (콘솔 창 숨김)
`autoslide.spec` 파일에서 `console=False`로 변경

### 아이콘 설정
1. `.ico` 파일 준비 (Windows) 또는 `.icns` 파일 (macOS)
2. `autoslide.spec`에서 `icon='path/to/icon.ico'` 설정

### 단일 파일 vs 폴더 배포
- 단일 파일 (현재 설정): 느리지만 배포 간편
- 폴더 배포: `--onedir` 옵션 사용, 빠르지만 여러 파일

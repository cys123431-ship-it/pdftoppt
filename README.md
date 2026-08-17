# PDF Converter (v1.4.0)

`PDF -> PPTX / DOCX / PNG / JPG`, `PDF 병합`, `PDF 분할`, `폴더 일괄 변환`을 지원하는 Windows GUI 프로그램입니다.

## 주요 기능

- 변환: `PDF -> PPTX`, `PDF -> DOCX`, `PDF -> PNG`, `PDF -> JPG`
- PDF 편집: 여러 PDF 병합, 페이지 단위 분할
- 페이지 범위 지정: `1-3,5,8-10` 형식 지원
- 파일 큐 + 드래그앤드롭: 여러 PDF를 큐에 추가해 작업 가능
- 암호 PDF 지원:
  - 입력 PDF 비밀번호(열기)
  - 출력 PDF 비밀번호(병합/분할 결과 저장 시)
- 출력 충돌 정책:
  - `Overwrite`
  - `Skip Existing`
  - `Auto Rename`
- 품질 옵션:
  - `Render DPI` (이미지/PPT 렌더링 품질)
  - `JPG quality`
- 일괄 변환 실패 로그 CSV 저장
- 작업 취소 버튼 지원
- 한국어 / English UI

## v1.4.0 안정성 개선

- 덮어쓰기 작업을 임시 파일에 먼저 완성한 뒤 교체하도록 변경해 변환 실패 시 기존 출력 파일 보호
- PDF 병합 결과를 입력 PDF 위에 덮어써 원본을 잃을 수 있던 경로 차단
- PNG/JPG 및 PDF 분할 작업을 트랜잭션 방식으로 처리해 취소 시 일부 결과만 남는 문제 방지
- 서로 다른 페이지 비율이 섞인 PDF를 PPTX로 변환할 때 이미지 비율을 유지하고 가운데 배치
- DOCX 변환을 별도 프로세스로 실행해 변환 중 취소 가능
- 병합 모드의 입력 상태를 파일 큐 하나로 통일해 `큐 비우기` 후 이전 선택이 다시 살아나는 문제 수정
- 현재 작업에 실제로 필요한 DPI/JPG 옵션만 검증
- 변환 결과 메시지 한국어 표시 개선
- 실패 로그 CSV의 수식 주입 위험 완화
- 의존성 버전 고정 및 Windows CI 자동 테스트 추가

## 지원 작업 목록

- 단일 PDF 변환
  - `PDF -> PPTX`
  - `PDF -> DOCX`
  - `PDF -> PNG/JPG`
- 다중 PDF 병합 (`Merge PDFs`)
- 단일 PDF 분할 (`Split PDF`)
- 폴더 일괄 변환 (`Batch Convert Folder`)
  - 출력 형식: `PPTX`, `DOCX`, `PNG`, `JPG`

## 설치 및 실행

### 1) 실행 파일 사용 (권장)

Releases 페이지에서 최신 `PDFConverter.exe`를 다운로드해 실행하세요.

- Releases: https://github.com/cys123431-ship-it/pdftoppt/releases

### 2) 소스 코드 실행

```bash
git clone https://github.com/cys123431-ship-it/pdftoppt.git
cd pdftoppt
pip install -r requirements.txt
python main.py
```

## 테스트

```bash
python -m unittest discover -s tests -v
```

## 빌드 (Windows EXE)

```bash
pyinstaller --noconfirm --clean --noconsole --onefile --name PDFConverter main.py
```

빌드 결과물: `dist/PDFConverter.exe`

## 자동 릴리스

GitHub Actions 워크플로우가 설정되어 있습니다.

- Pull Request: Windows에서 의존성 설치 후 단위 테스트 실행
- `release:`로 시작하는 커밋이 `main`에 반영되면:
  - `VERSION` 파일 기준 태그 생성
  - Windows EXE 빌드
  - GitHub Release 생성 또는 갱신
  - `PDFConverter.exe` 에셋 업로드
- `workflow_dispatch`를 이용한 수동 릴리스도 지원

## 기술 스택

- Python 3.11 (release build)
- Tkinter
- tkinterdnd2
- PyMuPDF (fitz)
- python-pptx
- pdf2docx
- PyInstaller

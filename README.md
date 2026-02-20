# PDF Converter (v1.2.0)

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

릴리스 페이지에서 최신 `PDFConverter.exe`를 다운로드해 실행하세요.

- Releases: https://github.com/cys123431-ship-it/pdftoppt/releases

### 2) 소스 코드 실행

```bash
git clone https://github.com/cys123431-ship-it/pdftoppt.git
cd pdftoppt
pip install -r requirements.txt
python main.py
```

## 빌드 (Windows EXE)

```bash
pyinstaller --noconfirm --clean --noconsole --onefile --name PDFConverter main.py
```

빌드 결과물: `dist/PDFConverter.exe`

## 자동 릴리스

GitHub Actions 워크플로우가 설정되어 있습니다.

- 파일: `.github/workflows/windows-release.yml`
- 태그 `v*` 푸시 시:
  - Windows에서 EXE 빌드
  - GitHub Release 생성
  - `PDFConverter.exe` 에셋 업로드

## 기술 스택

- Python
- Tkinter
- tkinterdnd2
- PyMuPDF (fitz)
- python-pptx
- pdf2docx
- PyInstaller

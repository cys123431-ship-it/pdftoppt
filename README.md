# PDF Converter (v1.5.0)

Windows용 PDF 변환/편집 GUI 프로그램입니다. 일반 변환뿐 아니라 **무료 로컬 OCR**, PDF 미리보기와 페이지 편집, 설정 저장까지 지원합니다.

## 주요 기능

### PDF 변환

- `PDF -> PPTX`
- `PDF -> DOCX`
- `PDF -> PNG`
- `PDF -> JPG`
- 페이지 범위 지정: `1-3,5,8-10`

### 무료 로컬 OCR

릴리스 `PDFConverter.exe`에는 Tesseract OCR 런타임과 한국어/영어 언어 데이터가 포함되어 있어 별도 API 키나 인터넷 연결 없이 OCR을 수행합니다.

- `OCR -> Searchable PDF`
  - 스캔 PDF에 검색 가능한 텍스트 레이어 생성
- `OCR -> TXT`
  - PDF의 텍스트를 UTF-8 텍스트 파일로 추출
- `OCR -> DOCX`
  - OCR 결과를 편집 가능한 Word 문서로 저장
- OCR 언어
  - `kor+eng` 한국어 + 영어
  - `kor` 한국어
  - `eng` 영어
- OCR 모드
  - `auto`: 이미 충분한 텍스트가 있는 페이지는 원본 텍스트를 사용하고 스캔/이미지 페이지에만 OCR 적용
  - `force`: 모든 선택 페이지를 OCR 처리
- OCR DPI: 72~600
- OCR 작업 취소 지원
- 모든 OCR 처리는 로컬 PC에서 수행

### PDF 미리보기 / 페이지 편집

- 선택 페이지 미리보기
- 페이지 순서 변경
  - 위/아래 버튼
  - 목록에서 드래그 정렬
- 페이지 삭제
- 페이지 좌/우 90도 회전
- 편집 결과를 새 PDF로 저장
- 기존 출력 파일 보호를 위한 임시 파일 + 원자적 교체 방식 사용

### PDF 편집 / 일괄 작업

- 여러 PDF 병합
- PDF 페이지 단위 분할
- 폴더 일괄 변환
  - `PPTX`, `DOCX`, `PNG`, `JPG`
- 파일 큐 + 드래그앤드롭
- 입력 PDF 비밀번호
- 출력 PDF 비밀번호
  - 병합 / 분할 / OCR PDF / 페이지 편집
- 출력 충돌 정책
  - `Overwrite`
  - `Skip Existing`
  - `Auto Rename`
- 일괄 변환 실패 로그 CSV 저장

### 편의 기능

- 한국어 / English UI
- 완료 후 출력 폴더 자동 열기
- 설정 자동 저장
  - 언어
  - 충돌 정책
  - 렌더 DPI
  - JPG 품질
  - OCR 언어/모드/DPI
  - 배치 대상
  - 실패 로그 여부
  - 출력 폴더 자동 열기 여부
  - 최근 출력 폴더
- 비밀번호는 설정 파일에 저장하지 않음
- 설정 위치: `%APPDATA%\PDFConverter\settings.json`

## v1.5.0 추가 기능

- Tesseract 기반 오프라인 OCR 추가
- 검색 가능한 OCR PDF / TXT / DOCX 출력 추가
- 한국어, 영어, 한국어+영어 OCR 지원
- 자동 OCR 감지 및 강제 OCR 모드 추가
- OCR 중 취소 지원
- PDF 미리보기 및 페이지 순서 변경/삭제/회전 추가
- 페이지 드래그 정렬 추가
- 앱 설정 영구 저장 추가
- 변환 완료 후 출력 폴더 자동 열기 추가
- OCR 관련 단위 테스트, 페이지 편집 테스트, 설정 저장 테스트 추가
- Windows 릴리스 EXE에 Tesseract 런타임과 `eng`, `kor`, `osd` 언어 데이터를 함께 번들

## v1.4.0 안정성 개선

- 덮어쓰기 작업을 임시 파일에 먼저 완성한 뒤 교체하도록 변경해 변환 실패 시 기존 출력 파일 보호
- PDF 병합 결과를 입력 PDF 위에 덮어써 원본을 잃을 수 있던 경로 차단
- PNG/JPG 및 PDF 분할 작업을 트랜잭션 방식으로 처리해 취소 시 일부 결과만 남는 문제 방지
- 서로 다른 페이지 비율이 섞인 PDF를 PPTX로 변환할 때 이미지 비율 유지
- DOCX 변환을 별도 프로세스로 실행해 변환 중 취소 가능
- 병합 모드의 입력 상태를 파일 큐 하나로 통일
- 실패 로그 CSV의 수식 주입 위험 완화
- 의존성 버전 고정 및 Windows CI 자동 테스트 추가

## 설치 및 실행

### 1) 릴리스 EXE 사용

Releases 페이지에서 최신 `PDFConverter.exe`를 받아 실행합니다.

- 릴리스 EXE는 OCR 엔진을 포함하므로 Tesseract를 별도로 설치할 필요가 없습니다.

### 2) 소스 코드 실행

```bash
git clone https://github.com/cys123431-ship-it/pdftoppt.git
cd pdftoppt
pip install -r requirements.txt
python main.py
```

소스에서 OCR 기능을 사용할 경우 Tesseract 5를 별도로 설치해야 합니다. 환경 변수 `TESSERACT_CMD`로 `tesseract.exe` 경로를 직접 지정할 수도 있습니다.

## 테스트

```bash
python -m unittest discover -s tests -v
```

## Windows 릴리스 빌드

GitHub Actions의 `.github/workflows/windows-release.yml`이 공식 릴리스 빌드를 수행합니다.

릴리스 빌드 과정에서:

1. Python 의존성 설치
2. Tesseract 5 설치
3. `tessdata_fast` 4.1.0의 `eng`, `kor`, `osd` 모델 준비
4. 필요한 Tesseract 런타임/DLL/언어 데이터만 번들 디렉터리에 복사
5. PyInstaller one-file EXE 생성
6. GitHub Release에 `PDFConverter.exe`와 `THIRD_PARTY_NOTICES.md` 업로드

## 자동 릴리스

- Pull Request: Windows에서 단위 테스트 실행
- `release:`로 시작하는 커밋이 `main`에 반영되면:
  - `VERSION` 파일 기준 태그 생성
  - Windows EXE 빌드
  - GitHub Release 생성 또는 갱신
  - 릴리스 에셋 업로드
- `workflow_dispatch` 수동 릴리스 지원

## 기술 스택

- Python 3.11
- Tkinter / tkinterdnd2
- PyMuPDF
- python-pptx
- pdf2docx
- python-docx
- Tesseract OCR 5
- tessdata_fast (`eng`, `kor`, `osd`)
- PyInstaller

## 라이선스 / 제3자 구성요소

OCR 런타임 및 언어 데이터 등 제3자 구성요소 정보는 `THIRD_PARTY_NOTICES.md`를 참고하세요.

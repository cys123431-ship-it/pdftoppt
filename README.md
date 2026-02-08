# PDF to PPT Converter (PDF-PPT 변환기)

간단하고 강력한 PDF to PowerPoint 변환 도구입니다. 각 PDF 페이지를 고해상도 이미지로 변환하여 파워포인트 슬라이드에 삽입합니다.

## 특징 (Features)

*   **100% 원본 유지**: PDF 페이지를 이미지로 변환하여 폰트 깨짐이나 레이아웃 변경 없이 원본 그대로 PPT에 담깁니다.
*   **간편한 사용**: 직관적인 GUI로 파일만 선택하면 끝!
*   **독립 실행**: Python이 설치되지 않은 PC에서도 `.exe` 파일 하나로 실행 가능합니다.

## 설치 및 실행 방법 (Installation & Usage)

### 실행 파일 사용 (Recommended)
1.  [Releases](https://github.com/cys123431-ship-it/pdftoppt/releases) 페이지에서 최신 `PDFtoPPTConverter.exe`를 다운로드합니다. (또는 `dist` 폴더 확인)
2.  다운로드한 파일을 실행합니다.
3.  "Select PDF" 버튼을 눌러 PDF를 선택하고 변환을 시작합니다.

### 소스 코드에서 실행 (For Developers)

이 프로젝트는 Python 3.x 기반입니다.

1.  저장소 클론:
    ```bash
    git clone https://github.com/cys123431-ship-it/pdftoppt.git
    cd pdftoppt
    ```

2.  의존성 설치:
    ```bash
    pip install -r requirements.txt
    ```

3.  앱 실행:
    ```bash
    python main.py
    ```

## 빌드 방법 (Building from Source)

PyInstaller를 사용하여 실행 파일을 만들 수 있습니다.

```bash
pip install pyinstaller
pyinstaller --noconsole --onefile --name "PDFtoPPTConverter" main.py
```

`dist` 폴더에 실행 파일이 생성됩니다.

## 기술 스택 (Tech Stack)

*   [Python](https://www.python.org/)
*   [Tkinter](https://docs.python.org/3/library/tkinter.html): GUI
*   [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/): PDF Rendering
*   [python-pptx](https://python-pptx.readthedocs.io/): PPT Generation

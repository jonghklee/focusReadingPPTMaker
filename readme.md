# 포커스 리딩 키워드 훈련 PPT 생성기

이 프로그램은 *포커스 리딩*이라는 책의 '키워드 포착 훈련'을 편리하게 할 수 있도록 파워포인트(PPT)를 만들어주는 소프트웨어입니다. 이 README 파일은 설치, 빌드, 사용 방법에 대한 안내를 제공합니다.

## 주요 기능

- 키워드 포착 훈련을 위한 파워포인트 슬라이드를 자동으로 생성합니다.
- 사용하기 쉬운 인터페이스.
- 주로 윈도우에서 호환됩니다.

## 사전 준비 사항

- Python 3.x
- pip (Python 패키지 관리자)

## 설치 방법

1. 저장소를 로컬 머신에 클론합니다:

   ```bash
   git clone https://github.com/yourusername/focus-reading-ppt-generator.git
   ```

2. 프로젝트 디렉토리로 이동합니다:

   ```bash
   cd focus-reading-ppt-generator
   ```

3. 가상 환경을 만듭니다:

   ```bash
   python -m venv venv
   ```

4. 가상 환경을 활성화합니다:


   ```bash
   venv\Scripts\activate
   ```

5. 필요한 Python 패키지를 설치합니다:

   ```bash
   pip install -r requirements.txt
   ```

## 빌드 방법

### 윈도우에서 빌드

이 소프트웨어를 윈도우에서 빌드하려면 pyinstaller를 사용해야 합니다.

1. pyinstaller를 설치합니다:

   ```bash
   pip install pyinstaller
   ```

2. 소프트웨어를 빌드합니다:

   ```bash
   pyinstaller --onefile main.py
   ```

필요하다면 UPX를 사용하여 빌드된 파일의 크기를 줄일 수 있습니다. UPX는 고성능 실행 파일 압축기입니다. UPX에 대한 자세한 내용과 설치 방법은 [UPX GitHub 페이지](https://github.com/upx/upx)에서 확인할 수 있습니다.

### 맥OS에서 빌드

이 소프트웨어는 윈도우만을 지원하고 있기에 맥북에서 안정적으로 작동하려면 개량이 필요합니다. 맥OS에서 빌드하려면 pyinstaller 혹은 py2app을 사용하세요.

## 사용 방법

1. `dist/main.exe` 파일을 실행합니다.

2. 소프트웨어의 지시에 따라 키워드 포착 훈련을 위한 PPT를 생성합니다.

## 지원

문제가 발생하거나 개선 사항이 있으면 이슈 트래커에 보고해 주세요.

## 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 LICENSE 파일을 참조하세요.
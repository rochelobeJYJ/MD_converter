# 📄 통합 문서 → Markdown 변환기 (Document to Markdown Converter)

이 프로그램은 `opendataloader-pdf` 라이브러리와 Microsoft의 최신 `markitdown` 엔진을 모두 지원하여, PDF뿐만 아니라 Word, Excel, PPT, 이미지, 오디오 파일 등 일상적인 모든 포맷을 AI가 읽기 좋은 최적의 Markdown 파일로 변환해 주는 통합 로컬 데스크탑 애플리케이션입니다.

## 🌟 주요 기능
* **듀얼 변환 엔진 선택**:
  * **OpenDataLoader (PDF만)**: 논문, 서적 등 복잡한 다단, 표, 목차 구조의 PDF 파일을 무손실 이미지 추출과 함께 전문적으로 파싱합니다.
  * **MarkItDown (다양한 포맷)**: `.docx`, `.xlsx`, `.pptx`, `.jpg`, `.png`, `.mp3`, `.wav`, `.html`, `.csv`, `.zip` 등 Microsoft MarkItDown 통합 엔진으로 거의 모든 문서/미디어 포맷을 즉시 Markdown으로 변환합니다.
* **간단한 Drag & Drop**: 변환할 파일들을 목록에 끌어다 놓는 것만으로 추가 가능. 엔진 전환 시 호환되지 않는 확장자는 자동으로 걸러냅니다.
* **다국어 지원 (KOR/ENG)**: 하단의 라디오 버튼으로 한국어와 영어 UI를 실시간으로 전환하여 글로벌 환경에서도 편리하게 사용할 수 있습니다.
* **최적화된 아이콘 및 UI**: "모든 문서를 MD로 변환한다"는 직관적인 아이콘과, 직관적인 세로형 엔진 선택 버튼, 눈이 편안한 그레이톤 테마를 제공합니다.
* **멀티스레딩 & 일괄 처리**: 변환 도중 프로그램이 멈추지 않으며, 여러 개의 문서를 동시에 안정적으로 처리합니다.
* **백그라운드 무소음 실행**: 변환 작업 중 눈에 거슬리는 검은색 콘솔(터미널) 창이 깜빡이지 않도록 내부 프로세스가 최적화되어 있습니다.
* **스마트 목록 관리**: 여러 작업을 반복할 때 편리하도록, 변환이 완료된 후 새로운 파일을 드래그 앤 드롭 하면 기존 목록을 자동으로 비워줍니다.
* **저장 폴더 자동 열기**: 모든 변환이 무사히 끝나면 변환 결과물이 저장된 폴더를 윈도우 탐색기로 즉시 띄워주는 편의 기능을 내장했습니다.

<img width="600" height="583" alt="image" src="https://github.com/user-attachments/assets/fa3bea04-b120-40a0-ae5a-f4f4386453dd" />

<img width="600" height="583" alt="image" src="https://github.com/user-attachments/assets/ec799762-9c25-481c-9a81-9ca831c59ee4" />




## ⚙️ 시스템 요구사항
이 프로그램은 완전한 오프라인 환경에서 동작합니다. 단, 다음과 같은 런타임 환경이 필요합니다.
* **OS**: Windows 10/11
* **Java**: **Java 11 이상** (`Java 17` 권장, OpenJDK 등)  
  *(OpenDataLoader의 레이아웃 분석기 작동을 위해 필수)*
* **Python**: Python 3.10 이상 (Anaconda 환경 지원)

## 📦 의존성 패키지 설치
Anaconda Prompt 등 개발된 터미널 환경에서 아래 명령어를 실행하여 필수 패키지를 설치하십시오.
```bash
pip install opendataloader-pdf tkinterdnd2 pillow markitdown
```

## 🚀 실행 및 빌드 방법
### 직접 실행하기
수동으로 터미널에서 실행하려면 다음 명령어를 사용하세요.
```bash
# 콘솔 창 없이 백그라운드 UI만 실행 (사용자 권장)
pythonw start.pyw

# 디버깅/에러로그 확인용 콘솔과 함께 실행
python app.py
```

### 독립 실행형 빌드 (exe 제작)
프로젝트에 동봉된 `build_dist.bat` 파일을 실행하면, `build_helper.py`가 작동하면서 **사용자 PC에 Python 환경이 없어도 동작하는 독립형 `.exe` 파일**과 `앱 폴더.zip` 배포본이 `dist` 폴더 안에 자동으로 완성됩니다. (자동으로 conda 환경을 만들고 지우며 빌드합니다.)

## 💡 사용 팁
1. **각 엔진의 장점 극대화**: 고화질 그래프 이미지가 포함된 학술 논문 PDF는 `OpenDataLoader` 엔진을 권장하며, 일반적인 문서(Word, Excel 등) 및 사진 파일의 메타/OCR 요약은 `MarkItDown` 엔진을 사용하시는 것이 가장 빠르고 효율적입니다.
2. **저장 경로 옵션**: 기본적으로 원본 문서 크기 파일이 있던 위치에 동일한 이름(`파일명.md`)으로 결과물이 저장되지만, 필요 시 앱 UI 상에서 "폴더 변경" 버튼을 눌러 결과물만 보관하는 고정 폴더(예: 바탕화면 내 임의의 폴더)를 지정해 둘 수 있습니다.

## ⚠️ 라이선스 & 주의사항
* 본 프로그램의 파싱 엔진은 `opendataloader-project`와 Microsoft의 `markitdown` 코어 오픈소스에 의존하고 있습니다.
* 문서에 보안 및 파일 암호(비밀번호)가 걸려있을 경우 미리 잠금을 제거한 후 진행해야 변환 실패를 피할 수 있습니다.

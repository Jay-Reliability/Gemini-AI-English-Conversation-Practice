# 🎓 Gemini AI Enlgish Conversation Tutor V1.0

기존 Windows PC(PyQt5) 환경에서 실행되던 `01_Gemini_API_Eng_Study.py`의 영어회화 튜터링 및 실시간 문장 교정(Error Correction) 기능을 그대로 유지하면서, 다중 사용자 회원 관리 및 학습 이력 복습/엑셀 내보내기 기능을 추가하여 **Streamlit Web** 환경에서 실행되도록 개발한 웹 애플리케이션입니다.

---

## 📂 프로젝트 파일 구성

| 파일명 | 설명 |
| :--- | :--- |
| `English_Study_app.py` | Streamlit 기반 웹 애플리케이션 메인 UI 및 이벤트 라우팅 스크립트 |
| `collector.py` | 회원 관리(신규등록/로그인/관리자 초기화), Gemini API 통신, 사용자별 엑셀 로그 수집·필터링·줄바꿈 저장·삭제 모듈 |
| `requirement.txt` | GitHub 및 로컬 환경 패키지 설치 명세서 (`requirements.txt` 동시 제공) |
| `README.md` | 프로젝트 소개, 설치/실행 방법 및 기능 명세 문서 |
| `Gemini_English.xlsx` | 사용자별 영어회화 대화 이력 및 문장 교정 결과 저장 엑셀 파일 |

---

## ✨ 주요 기능

### 1. 상단 7대 메뉴 구성
프로그램 실행 시 상단에 아래 7개 메뉴 버튼이 고정 배치됩니다.
- **`신규등록`** :
  - 신규 회원이 `ID`와 `Password(2회 확인)`를 입력하여 사용 권한을 부여받습니다.
  - 이미 등록된 ID를 입력할 경우 `"회원 등록이 이미 되었습니다"` 메시지 창을 표시합니다.
  - **ID/PW 분실 시 초기화 기능**: 기본 생성된 관리자 계정(`ID: admin` / `PW: jhlee_english01`)을 통해 분실 회원의 ID 목록을 조회하거나 비밀번호/계정을 초기화할 수 있습니다.
- **`Log In`** :
  - 등록된 `ID`와 `Password`가 일치하면 프로그램을 활성화합니다.
  - 일치하지 않을 경우 즉시 **Error Message 창**을 표시합니다.
- **`Gemini 영어공부`** :
  - `01_Gemini_API_Eng_Study.py`와 동일한 Tone & Manner 및 기능으로 동작합니다.
  - **Font Size 조절(8~50px)** 및 실시간 UI 반영, **Conversation Topic** 설정 기능을 제공합니다.
  - 사용자 문장(`Me`, 파란색)에 문법/표현 교정이 필요한 경우 즉시 `[Error Correction]`(빨간색)으로 최적의 표현(Best Expression)을 제시하고, 자연스러운 회화 응답(`Gemini`, 초록색)을 이어갑니다.
  - 모든 대화 결과는 `Gemini_English.xlsx` 포맷의 **첫 번째 열에 `User`(로그인 ID)를 추가**하여 실시간 저장합니다.
- **`영어공부 복습`** :
  - 로그인한 본인의 `User` ID에 해당하는 영어회화 로그만 선별하여 출력합니다 (타인의 ID 로그는 출력하지 않음).
  - **제목 행(`User | Topic | Date | Who | Conversation History | Error Correction`)은 상단에 고정**되며, 내용이 길어질 경우 우측 스크롤바를 통해 상하 스크롤이 가능합니다.
- **`엑셀 저장`** :
  - 로그인한 `User` ID의 영어회화 로그(`Me`, `Gemini`)만 추출하여 원하는 폴더와 파일명으로 엑셀 파일(`.xlsx`)을 저장하거나 브라우저로 다운로드할 수 있습니다.
  - 문장 길이가 셀 크기를 초과할 경우 셀 내부 줄바꿈(`\n`, Excel 내 `Shift+Enter`/`Alt+Enter` 및 `wrap_text=True`)을 자동 적용하여 글자가 셀 크기 내에 깔끔하게 표시됩니다.
- **`모든 데이터 삭제`** :
  - 현재 로그인한 ID의 영어회화 로그를 모두 삭제할지 묻는 **삭제 확인 창**을 표시하고, `삭제` 버튼을 누르면 해당 ID의 데이터만 모두 삭제합니다.
- **`Log Out`** :
  - 현재 세션을 안전하게 종료하고 `Log In` 전 단계 화면으로 돌아갑니다.

---

## 📊 데이터 출력 및 엑셀 저장 포맷

`영어공부 복습` 및 `엑셀 저장` 시 아래와 같이 첫 번째 열에 `User`가 포함된 6개 열 구조로 출력 및 저장됩니다.

| User | Topic | Date | Who | Conversation History | Error Correction |
| :--- | :--- | :--- | :--- | :--- | :--- |
| Jay | Dinner | 2026-10-04 20:21:10 | Me | What do you say watching movie together after work? | How about watching a movie together after work? |
| Jay | Hobby | 2026-10-04 20:41:04 | Me | I am 57 year old and I'm very interested in keeping my helath. so, I usually enjoy hiking in my free time on weekend. | I am 57 years old and I'm very interested in staying healthy, so I usually enjoy hiking in my free time on weekends. |
| Jay | Hobby | 2026-10-04 20:41:11 | Gemini | That's a fantastic hobby! Hiking is such a great way to stay active and enjoy the outdoors. Do you have a favorite mountain or trail you like to hike? | |

---

## 🚀 설치 및 실행 방법

### 1. 필수 패키지 설치
```bash
pip install -r requirement.txt
```

### 2. Google Gemini API Key 설정
프로젝트 루트 폴더의 `.env` 파일에 아래와 같이 API 키를 설정하거나, 웹 실행 후 좌측 사이드바의 **환경 설정** 입력창에 직접 입력할 수 있습니다.
```env
GOOGLE_API_KEY=your_google_gemini_api_key_here
```
*(Streamlit Community Cloud 배포 시에는 App Settings -> **Secrets**에 `GOOGLE_API_KEY = "your_api_key"`를 등록하세요.)*

### 3. Streamlit 웹 프로그램 실행
```bash
streamlit run English_Study_app.py
```

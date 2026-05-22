# Excel to VPR Converter (Flask Web App)

이 프로젝트는 자동차 NVH(Noise, Vibration, Harshness) 테스트 데이터를 포함한 개별 Excel 파일을 업로드하여, 사전에 정의된 양식의 HTML 리포트로 자동 변환하고 관리하는 Flask 기반의 웹 애플리케이션입니다.

## 🚀 주요 기능

### 1단계: 동적 파일 업로드 및 검증
- **Multi-Day 지원**: `+`/`-` 버튼을 통해 여러 날짜(Day 1, Day 2...)의 엑셀 파일을 동시에 업로드할 수 있습니다.
- **실시간 유효성 확인**: 파일을 업로드하는 즉시 서버에서 읽기 가능 여부를 판단하여 **OK/NG** 상태를 표시합니다.

### 2단계: 정밀 엑셀 파싱 및 리포트 초안 생성
- **세부 진행률 표시**: 대용량 엑셀 로딩 시간을 고려하여, 실시간 분석 단계 및 에너지바 애니메이션(0~100%)을 제공합니다.
- **자동 데이터 추출**: `Example_Plain txt.txt` 양식에 따라 의뢰번호, 프로젝트명, 평가자, 차량 정보, UNIQUE 타이어 사양, 온도 통계(Min~Max) 등을 자동 파싱합니다.
- **A4 디자인**: 리포트 편집 화면은 실제 A4 종이 질감의 레이아웃을 제공하며, 창 크기에 따라 반응형으로 조절됩니다.

### 3단계: 리포트 관리 및 DB 저장
- **최종 HTML 변환**: 수동 입력값을 포함하여 깨끗한 텍스트 형태의 HTML 문서를 생성합니다.
- **데이터베이스 통합**: 모든 리포트는 SQLite DB에 저장되어 언제든 다시 불러올 수 있습니다.
- **재편집 기능**: 저장된 리포트를 목록에서 선택하여 확인하고, 필요 시 다시 수정 모드로 진입할 수 있습니다.

---

## 🛠 기술 스택
- **Backend**: Python 3.12+ / Flask / Flask-SQLAlchemy (SQLite)
- **Excel Engine**: openpyxl
- **Frontend**: HTML5, CSS3, Vanilla JavaScript (SSE 기반 비동기 통신)

---

## 💻 실행 방법

### 1. 필수 라이브러리 설치
터미널에서 아래 명령어를 실행하여 필요한 패키지를 설치합니다.
```bash
pip install flask pandas openpyxl flask-sqlalchemy
```

### 2. 서버 실행
프로젝트 폴더(`JUNWON`) 내에서 서버를 구동합니다.
```bash
python app.py
```

### 3. 웹페이지 접속
브라우저를 열고 아래 주소로 접속합니다.
- **URL**: [http://127.0.0.1:5000](http://127.0.0.1:5000)

---

## 📂 프로젝트 구조
```text
JUNWON/
├── app.py              # Flask 백엔드 서버 및 엑셀 파싱 로직
├── templates/
│   └── index.html      # 메인 웹 레이아웃
├── static/
│   ├── css/style.css   # 리포트 디자인 및 애니메이션 스타일
│   └── js/script.js    # 동적 UI 및 SSE 통신 클라이언트 로직
├── uploads/            # 엑셀 파일 임시 업로드 폴더
├── instance/           # SQLite DB 저장 폴더
├── GEMINI.md           # AI 에이전트 프로젝트 지침서
└── README.md           # 프로젝트 실행 가이드 (현재 파일)
```

## 📋 참고 문서
- `EXCEL TO VPR PROMPT.md`: 상세 단계별 구현 요구사항
- `Example_Plain txt.txt`: 리포트 텍스트 양식 가이드
- `Design.Plain txt.md`: 시각적 디자인 가이드
- `info_git.md`: Git Ignore 관리 규칙

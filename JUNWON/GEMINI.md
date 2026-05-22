# Excel to VPR Converter 프로젝트 지침

이 프로젝트는 개별 Excel 파일을 업로드하여 정해진 양식의 HTML 리포트로 변환하고 DB에 저장하는 Flask 기반 웹 애플리케이션입니다.

## 기술 스택
- **Backend**: Python Flask, Flask-SQLAlchemy (SQLite)
- **Excel Parsing**: openpyxl (Pandas는 검증용으로만 사용)
- **Frontend**: Vanilla JS, CSS (Modular 구조: static/templates)

## 프로젝트 구조 및 핵심 경로
- 루트 디렉토리: `D:\새 폴더\NVH-Vibe-Coding\JUNWON` (현재 이 폴더에서 모든 작업 진행)
- `app.py`: 메인 서버 로직 및 SSE(Server-Sent Events) 기반 파싱 엔진
- `static/js/script.js`: 동적 UI 및 리포트 편집 로직
- `templates/index.html`: 메인 레이아웃

## 주요 구현 규칙
1. **실시간 파싱 상태 (SSE)**: 엑셀 로딩이 오래 걸리므로 `app.py`의 `/extract` 라우트는 실시간으로 진행률(%)과 상세 단계 메시지를 전송해야 함.
2. **A4 종이 디자인**: 리포트 편집 및 보기 화면은 `Design.Plain txt.md`에 따라 A4 비율의 흰색 종이 질감과 그림자 효과를 유지해야 함.
3. **데이터 매핑**: `Example_Plain txt.txt`의 양식을 100% 준수해야 하며, UNIQUE 정보 추출(H17~Q20 범위) 및 온도 통계 계산 로직이 포함되어 있음.
4. **DB 저장**: `VPRReport` 모델은 최종 HTML뿐만 아니라 재편집을 위해 `raw_data`(JSON) 필드를 반드시 포함해야 함.
5. **로그 정리**: `openpyxl` 라이브러리의 `DrawingML` 관련 `UserWarning`은 데이터 추출에 영향을 주지 않으므로 `warnings` 모듈을 사용해 무시(suppress) 처리함.
6. **UI 렌더링 주의**: DOM 조작 시(예: 업로드 영역 숨김 처리) 대상 DOM(`manual-input-section`)이 부모 요소(`upload-section-main`) 내부에 잘못 중첩(nesting)되어 함께 숨겨지지 않도록 HTML 구조의 완전한 분리를 유지할 것.

## 실행 방법
1. `pip install flask pandas openpyxl flask-sqlalchemy` 설치
2. `python app.py` 실행
3. `http://127.0.0.1:5000` 접속

## 현재 진행 상태
- 1~3단계 기본 기능 구현 완료.
- 리포트 목록 보기, 최종 HTML 저장, 저장된 리포트 불러오기 및 수정 기능 활성화됨.

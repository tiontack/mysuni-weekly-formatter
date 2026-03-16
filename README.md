# mySUNI Weekly Formatter

mySUNI Weekly 양식 정리 도구입니다.

기능:

- `docx` 업로드 후 3단 구조 재정렬
- 웹페이지 내 직접 입력/수정
- 회사 용어집/영문 이름 + Gemini 기반 영문 번역
- Word 내보내기 / 인쇄 / PDF 저장
- 통합 문서 생성 / 아카이브 저장

## 로컬 실행

```bash
cd "/Users/shinhyeontaek/Library/Mobile Documents/com~apple~CloudDocs/mysuni-weekly-formatter"
node server.js
```

브라우저에서 `http://127.0.0.1:8002` 을 엽니다.

## 환경 변수

- 로컬: `.env.local`
- 배포: 서버 환경변수 `GEMINI_API_KEY`

예시는 [.env.example](/Users/shinhyeontaek/Library/Mobile%20Documents/com~apple~CloudDocs/mysuni-weekly-formatter/.env.example) 참고

## 배포

정적 GitHub Pages만으로는 AI 번역을 안전하게 운영할 수 없습니다.

안전한 방식:

- 프론트엔드 + `/api/translate` 서버 엔드포인트를 함께 배포
- `GEMINI_API_KEY`는 서버 환경변수로만 보관
- 현재 프로젝트에는 로컬 서버용 [server.js](/Users/shinhyeontaek/Library/Mobile%20Documents/com~apple~CloudDocs/mysuni-weekly-formatter/server.js) 와 서버리스용 [api/translate.js](/Users/shinhyeontaek/Library/Mobile%20Documents/com~apple~CloudDocs/mysuni-weekly-formatter/api/translate.js) 가 포함되어 있음

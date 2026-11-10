# Claude-Gemini 연동

## 개요

Claude Code에서 Gemini를 불러 대용량 자료 독해, 구글 검색 기반 조사 같은 일을 맡기는 환경을 만든다. ChatGPT도 Codex CLI로 연결해 교차 검토에 쓴다(`.claude/skills/chatgpt/`). Gemini는 회사 계정(kwon.kyungmin@sll.co.kr)의 Google AI Studio API 키를 쓴다. 연동 방식은 `.claude/skills/gemini/`의 스크립트로 Gemini API를 직접 부르는 것이다.

- 계정 구성: Claude 개인 계정, Gemini 회사 계정, ChatGPT 남편 계정(회사 자료는 보내지 않음)
- Gemini CLI는 2026-06-18부터 구글 계정 로그인과 무료 API 키를 받지 않아 쓰지 않는다(설치는 되어 있음, v0.65.0).

## 현재 상태

마지막 확인일 2026-10-23 / 완료(사용자 종료 처리). `/gemini` 스킬로 일반 질문과 구글 검색 조사가 작동함(무료 등급). 무료 등급이라 공개 자료로만 쓴다. 파일 첨부 기능은 아직 시험하지 않았다. `/chatgpt` 스킬도 작동함(남편 계정이라 회사 자료는 보내지 않음).

## 다음 행동

## 아직 결정하지 않은 것

## 진행 기록

- 2026-10-21: Gemini CLI 회사 계정 로그인 시도 → 개인용 지원 종료로 실패. AI Studio 접속·API 키 발급 가능(무료 등급) 확인. API 직접 호출 방식으로 연동 시작.
- 2026-10-21: `.claude/skills/gemini` 스킬 작성, API 키 환경변수 등록 후 시험 통과. 무료 등급 검색 할당량은 gemini-2.5-flash에만 있어 검색 시 자동 전환하도록 설정.
- 2026-10-21: Codex CLI 0.163.0 설치, 남편 ChatGPT 계정 로그인, `.claude/skills/chatgpt` 스킬 작성·시험 통과. 기존 Codex 앱 설정(모델 고정, 브라우저 플러그인)은 불러오지 않도록 함.
- 2026-10-21: 7월 Codex 데스크톱 앱 잔여 파일(실행 파일·런타임·플러그인·설정, 약 2.9GB) 휴지통으로 이동. 예전 대화 기록 DB와 로그인은 남김. Windows 계정 CodexSandboxOffline/Online은 남아 있음(삭제는 본인이 판단).
- 2026-10-23: 업무 종료(사용자). 스킬은 작동 상태로 둠. 쓸모 판단·IT 문의와 미결 2건(유료 전환, 1/28 키 삭제)은 닫음

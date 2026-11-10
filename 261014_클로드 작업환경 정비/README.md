# 클로드 작업환경 정비

## 개요
- 시작일: 2026-10-14
- 내용: 퍼블리 아티클 「손은 편해졌는데 머리는 더 바빠졌다」(김민석)의 방식(업무 폴더, CLAUDE.md, /worklog, 상태판)을 Windows 환경의 SLL Claude 폴더에 적용
- 결정 사항 (사용자 선택)
  - 개인 지침 다섯 줄은 전역(`C:\Users\Jnote\.claude\CLAUDE.md`)으로 옮기지 않고 이 폴더 CLAUDE.md에 둔다
  - 상태판은 `_state.md`로 따로 두지 않고 각 업무 README.md 위쪽 섹션으로 둔다 (진행 기록이 `_log.md` 역할)
  - 업무일지는 최상위 별도 폴더 없이 각 업무 README에 기록한다
  - `/today` 스킬을 만든다

## 현재 상태
마지막 확인일 2026-11-10 / 보스 요청으로 클라우드 환경 전환 중. 작업 폴더를 `D:\SLL Claude`(git 저장소)로 옮겼고, GitHub 업로드까지 끝났고 Claude 연결이 남았다.

## 다음 행동
- (사용자) Claude에 GitHub 연결 → 클라우드 환경에서 저장소 선택해 시험 세션 실행
- (사용자) PC에서도 Claude 작업 폴더를 `D:\SLL Claude`로 바꿔 열기. G드라이브 폴더는 백업으로 두고 더 고치지 않음
- (사용자) 다음 작업을 마친 뒤 `/worklog` 실행 → README 기록 방식이 마음에 드는지 확인 후 SKILL.md 다듬기

## 아직 결정하지 않은 것

## 진행 기록
- 2026-10-14: 변경 전 CLAUDE.md를 `자료/261014 CLAUDE.md 변경 전.md`로 보관. CLAUDE.md 정리(일하는 방식 섹션 분리, 깨진 목록 기호 정리, `.claude/` 예외, README 상태판 형식·갱신 규칙 추가), 스킬 2개 생성(`.claude/skills/worklog/SKILL.md`, `.claude/skills/today/SKILL.md`), 기존 업무 README 3개에 상태 섹션 추가(기존 진행 기록은 그대로 둠)
- 2026-10-16: 새 대화에서 /today 첫 실행, 결과를 보고 사용자가 미결 3건 결정. '담당실적 확인 필요' 미결은 해당 업무 완료로 정리
- 2026-10-28: /today 결과를 체크·완료 표시가 되는 로컬 HTML 표로 작성(`결과/261028 오늘 할 일.html`, 체크는 브라우저 저장, 「완료 목록 복사」로 README 반영 요청)
- 2026-11-10: 출퇴근 중 이어서 작업하도록 클라우드 환경 전환 시작. 폴더 전체를 G드라이브에서 `D:\SLL Claude`로 복사(구글 시트 바로가기 1개 제외), git 저장소 생성·첫 커밋, `.gitignore` 추가. 회사 GitHub 계정에 비공개 저장소 sslkkm/sll-claude 생성(사용자)
- 2026-11-10: 사용자가 PowerShell에서 첫 업로드(push) 완료, GitHub main과 로컬 일치 확인

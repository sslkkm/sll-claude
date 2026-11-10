---
name: chatgpt
description: Codex CLI로 ChatGPT에 질문해 다른 시각의 답이나 교차 검토 의견을 받는다. Claude가 쓴 초안이나 아이디어를 다른 모델에게 한 번 더 검토받고 싶을 때 쓴다.
---

# ChatGPT 호출

Codex CLI에 로그인된 ChatGPT 계정으로 질문한다. 이 계정은 **남편 명의 계정**이다.

## 보내도 되는 자료

- 보낸 내용은 남편 계정의 기록과 사용 한도에 남는다. **회사 자료, 개인정보, 미공개 정보, 사적인 내용은 보내지 않는다.** 공개 정보와 일반적인 질문, 일반화한 아이디어 검토에만 쓴다.
- 회사 자료를 검토받고 싶으면 고유명사·수치·내부 정보를 빼고 일반화한 버전을 보내되, 보내기 전에 보낼 내용을 사용자에게 보여 주고 확인받는다.

## 사용법

```
python .claude/skills/chatgpt/chatgpt.py "질문"
python .claude/skills/chatgpt/chatgpt.py --image 이미지경로 "질문"
python .claude/skills/chatgpt/chatgpt.py --model 모델이름 "질문"
```

- 빈 임시 폴더에서 읽기 전용으로 실행하므로 Codex가 이 PC의 파일을 읽거나 고치지 못한다. 자료를 보여 주려면 질문 안에 필요한 내용만 넣는다.
- 답이 오기까지 수십 초 걸릴 수 있다.
- 로그인이 풀리면 사용자에게 PowerShell에서 `codex login`을 다시 실행하도록 안내한다. 로그인은 사용자가 직접 한다.

## 결과 다루기

- ChatGPT 답은 참고 의견으로 다룬다. 사실관계는 따로 확인한 뒤 쓴다.
- 사용자에게 전할 때는 ChatGPT가 낸 의견이라는 것을 밝힌다.

"""Codex CLI(ChatGPT 계정 로그인)로 ChatGPT에 질문하고 답을 출력한다.

사용 예:
  python chatgpt.py "질문"
  python chatgpt.py --image 시안.png "이 이미지에 대해 의견을 줘"
  python chatgpt.py --model gpt-5 "질문"

Codex가 파일을 읽거나 고치지 못하도록 빈 임시 폴더에서 읽기 전용으로 실행하고,
세션 기록은 디스크에 남기지 않는다.
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser()
    p.add_argument("prompt")
    p.add_argument("--model")
    p.add_argument("--image", action="append", default=[], help="첨부 이미지(여러 번 가능)")
    a = p.parse_args()

    codex = shutil.which("codex")
    if not codex:
        sys.exit("codex 명령을 찾을 수 없습니다. npm install -g @openai/codex 로 설치하세요.")

    with tempfile.TemporaryDirectory() as work:
        out = os.path.join(work, "answer.txt")
        # 사용자 config.toml(모델 고정, 브라우저·MCP 플러그인)은 불러오지 않는다. 로그인은 그대로 쓴다
        cmd = [codex, "exec", "--skip-git-repo-check", "--ephemeral", "--ignore-user-config",
               "-s", "read-only", "-C", work, "-o", out, "--color", "never"]
        if a.model:
            cmd += ["-m", a.model]
        for img in a.image:
            cmd += ["-i", os.path.abspath(img)]
        cmd.append("-")  # 질문은 stdin으로 넘겨 한글·따옴표 깨짐을 피한다
        r = subprocess.run(cmd, input=a.prompt, capture_output=True,
                           text=True, encoding="utf-8", errors="replace")
        if os.path.exists(out):
            with open(out, encoding="utf-8") as f:
                print(f.read().strip())
        if r.returncode != 0:
            sys.exit(f"Codex 오류(종료 코드 {r.returncode}):\n{r.stderr.strip()[-2000:]}")


if __name__ == "__main__":
    main()

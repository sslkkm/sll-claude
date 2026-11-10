"""Gemini API를 호출해 결과를 출력한다.

사용 예:
  python gemini.py "질문"
  python gemini.py --search "최근 OTT 시장 동향 요약"
  python gemini.py --file 자료.pdf "핵심 내용 요약"
  python gemini.py --prompt-file 지시서.md --file 자료.pdf --out 결과.md
  python gemini.py --list-models

API 키는 환경변수 GEMINI_API_KEY에서 읽는다. 현재 프로세스에 없으면
Windows 사용자 환경변수(레지스트리)에서 읽는다.
"""

import argparse
import base64
import json
import mimetypes
import os
import sys
import time
import urllib.error
import urllib.request

API = "https://generativelanguage.googleapis.com/v1beta"
DEFAULT_MODEL = "gemini-flash-latest"
# 무료 등급에서 구글 검색 근거는 2.5 Flash만 할당량이 있다(2026-10-21 확인)
SEARCH_MODEL = "gemini-2.5-flash"


def get_key():
    key = os.environ.get("GEMINI_API_KEY")
    if not key and sys.platform == "win32":
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment") as k:
                key, _ = winreg.QueryValueEx(k, "GEMINI_API_KEY")
        except OSError:
            pass
    if not key:
        sys.exit("GEMINI_API_KEY가 설정되어 있지 않습니다.")
    return key


def request(method, path, body=None):
    req = urllib.request.Request(
        f"{API}/{path}",
        data=json.dumps(body).encode() if body else None,
        headers={"x-goog-api-key": get_key(), "Content-Type": "application/json"},
        method=method,
    )
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=300) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            # 503은 일시적 과부하라 잠시 뒤 다시 시도한다
            if e.code == 503 and attempt < 2:
                time.sleep(10)
                continue
            sys.exit(f"API 오류 {e.code}: {e.read().decode(errors='replace')}")


def file_part(path):
    mime = mimetypes.guess_type(path)[0] or "text/plain"
    with open(path, "rb") as f:
        data = f.read()
    if mime.startswith("text/"):
        return {"text": f"[파일: {os.path.basename(path)}]\n{data.decode('utf-8', errors='replace')}"}
    return {"inline_data": {"mime_type": mime, "data": base64.b64encode(data).decode()}}


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    p = argparse.ArgumentParser()
    p.add_argument("prompt", nargs="?")
    p.add_argument("--model")
    p.add_argument("--search", action="store_true", help="구글 검색 근거 사용")
    p.add_argument("--file", action="append", default=[], help="첨부 파일(여러 번 가능)")
    p.add_argument("--prompt-file", help="질문(작업 지시서)을 파일에서 읽음")
    p.add_argument("--out", help="결과 전문을 이 파일에 저장하고 화면에는 요약만 출력")
    p.add_argument("--preview", type=int, default=300, help="--out 사용 시 화면에 보일 앞부분 글자 수(0이면 없음)")
    p.add_argument("--list-models", action="store_true")
    a = p.parse_args()
    if a.prompt_file:
        # PowerShell이 저장한 파일은 BOM이 붙어 있어 utf-8-sig로 읽는다
        with open(a.prompt_file, encoding="utf-8-sig") as f:
            a.prompt = f.read() + (f"\n\n{a.prompt}" if a.prompt else "")

    if a.list_models:
        for m in request("GET", "models?pageSize=200")["models"]:
            if "generateContent" in m.get("supportedGenerationMethods", []):
                print(m["name"].removeprefix("models/"), "-", m.get("displayName", ""))
        return
    if not a.prompt:
        p.error("질문을 입력하세요.")

    body = {"contents": [{"parts": [file_part(f) for f in a.file] + [{"text": a.prompt}]}]}
    if a.search:
        body["tools"] = [{"google_search": {}}]
    model = a.model or (SEARCH_MODEL if a.search else DEFAULT_MODEL)
    res = request("POST", f"models/{model}:generateContent", body)

    cand = (res.get("candidates") or [{}])[0]
    text = "".join(part.get("text", "") for part in cand.get("content", {}).get("parts", []))
    sources = cand.get("groundingMetadata", {}).get("groundingChunks", [])
    if sources:
        text += "\n\n출처:\n" + "\n".join(
            f"- {s.get('web', {}).get('title', '')} {s.get('web', {}).get('uri', '')}" for s in sources
        )

    if not a.out:
        print(text)
        return

    # 결과 전문은 파일로만 남기고, 화면에는 확인에 필요한 요약 정보만 낸다
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write(text + "\n")
    usage = res.get("usageMetadata", {})
    reason = cand.get("finishReason", "")
    print(f"저장: {a.out}")
    print(f"모델 {model} / {len(text):,}자 {text.count(chr(10)) + 1}줄 / 출처 {len(sources)}개"
          f" / 토큰 입력 {usage.get('promptTokenCount', '?')} 출력 {usage.get('candidatesTokenCount', '?')}")
    if reason and reason != "STOP":
        print(f"주의: 응답이 정상 종료되지 않음 ({reason}). 잘렸거나 차단됐을 수 있음")
    if a.preview:
        print("---")
        print(text[:a.preview] + (" …" if len(text) > a.preview else ""))


if __name__ == "__main__":
    main()

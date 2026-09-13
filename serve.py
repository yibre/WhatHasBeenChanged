"""
serve.py — 화면 파일만 내려주는 로컬 서버.

이 서버는 업로드 경로가 없다. 엑셀 파일을 받지도, 저장하지도 않는다.
브라우저가 index.html, app.js, worker.js, diff_engine.py를 받아간 뒤
비교는 전부 사용자 컴퓨터의 브라우저 안에서 일어난다.

    python serve.py            # http://127.0.0.1:8000 에서 열기
    python serve.py --port 9000 --no-browser
"""

from __future__ import annotations

import argparse
import http.server
import json
import os
import socketserver
import threading
import webbrowser

ROOT = os.path.dirname(os.path.abspath(__file__))


class Handler(http.server.SimpleHTTPRequestHandler):
    extensions_map = {
        **http.server.SimpleHTTPRequestHandler.extensions_map,
        ".js": "text/javascript",
        ".mjs": "text/javascript",
        ".py": "text/plain; charset=utf-8",
        ".json": "application/json",
        ".css": "text/css",
        ".html": "text/html; charset=utf-8",
    }

    def __init__(self, *a, **kw):
        super().__init__(*a, directory=ROOT, **kw)

    def end_headers(self):
        # 개발 중 캐시 때문에 옛 코드가 남는 일을 막는다
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self):
        # 폐쇄망용: wheels/ 폴더에 넣어둔 whl 파일 목록을 알려준다
        if self.path.split("?")[0] == "/wheels/list.json":
            d = os.path.join(ROOT, "wheels")
            names = sorted(f for f in os.listdir(d) if f.endswith(".whl")) \
                if os.path.isdir(d) else []
            body = json.dumps(names).encode()
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()

    def do_POST(self):  # 업로드 경로가 없다는 것을 코드로 못 박아 둔다
        self.send_error(405, "This server only serves files. Nothing is uploaded.")

    def log_message(self, fmt, *args):
        first = str(args[0]) if args else ""
        if first.startswith(("GET", "HEAD")):
            print(f"  {first}")


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


def main():
    p = argparse.ArgumentParser(description="엑셀 변경점 추적기 실행")
    p.add_argument("--port", type=int, default=8000)
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--no-browser", action="store_true")
    a = p.parse_args()

    missing = [f for f in ("index.html", "app.js", "worker.js", "styles.css", "diff_engine.py")
               if not os.path.exists(os.path.join(ROOT, f))]
    if missing:
        raise SystemExit(f"다음 파일이 없습니다: {', '.join(missing)}")

    url = f"http://{a.host}:{a.port}/"
    with Server((a.host, a.port), Handler) as httpd:
        print(f"엑셀 변경점 추적기 → {url}")
        print("파일은 브라우저 안에서만 열립니다. 이 서버로 전송되지 않습니다.")
        print("종료하려면 Ctrl+C\n")
        if not a.no_browser:
            threading.Timer(0.6, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n종료했습니다.")


if __name__ == "__main__":
    main()

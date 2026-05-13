from __future__ import annotations

import socket
import os
import sys
import threading
import time
import webbrowser

import uvicorn

from app.main import app


HOST = "127.0.0.1"
DEFAULT_PORT = 8007
MAX_PORT_TRIES = 20


def find_available_port(start_port: int = DEFAULT_PORT) -> int:
    for port in range(start_port, start_port + MAX_PORT_TRIES):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.2)
            if sock.connect_ex((HOST, port)) != 0:
                return port
    raise RuntimeError(f"{HOST}:{start_port}-{start_port + MAX_PORT_TRIES - 1} 범위에서 사용 가능한 포트를 찾지 못했습니다.")


def wait_until_ready(port: int, timeout_seconds: float = 12.0) -> bool:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.settimeout(0.3)
            if sock.connect_ex((HOST, port)) == 0:
                return True
        time.sleep(0.15)
    return False


def main() -> int:
    port = find_available_port()
    url = f"http://{HOST}:{port}/"

    config = uvicorn.Config(
        app,
        host=HOST,
        port=port,
        log_level="warning",
        access_log=False,
    )
    server = uvicorn.Server(config)
    thread = threading.Thread(target=server.run, daemon=True)
    thread.start()

    if wait_until_ready(port):
        print(f"근무표 생성기 실행 중: {url}")
        print("이 창을 닫거나 Ctrl+C를 누르면 서버가 종료됩니다.")
        if os.environ.get("WORK_SCHEDULER_NO_BROWSER", "").strip() != "1":
            webbrowser.open(url)
    else:
        print("서버 시작을 확인하지 못했습니다. 잠시 후 브라우저에서 다시 접속해 보세요.")
        print(url)

    try:
        while thread.is_alive():
            time.sleep(0.5)
    except KeyboardInterrupt:
        print("\n종료 중...")
        server.should_exit = True
        thread.join(timeout=5)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

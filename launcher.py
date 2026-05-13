from __future__ import annotations

import os
import socket
import sys
import multiprocessing
from pathlib import Path
import threading
import time
import traceback
import webbrowser

import uvicorn


HOST = "127.0.0.1"
DEFAULT_PORT = 8007
MAX_PORT_TRIES = 20


def debug_log(message: str) -> None:
    if os.environ.get("WORK_SCHEDULER_DEBUG", "").strip() != "1":
        return
    if getattr(sys, "frozen", False):
        log_path = Path(sys.executable).resolve().parent / "launcher-debug.log"
    else:
        log_path = Path(__file__).resolve().parent / "launcher-debug.log"
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} {message}\n")


def dump_stacks_after_delay(delay_seconds: float) -> None:
    if os.environ.get("WORK_SCHEDULER_DEBUG", "").strip() != "1":
        return
    time.sleep(delay_seconds)
    debug_log("stack-dump:start")
    for thread_id, frame in sys._current_frames().items():
        debug_log(f"thread:{thread_id}")
        for line in traceback.format_stack(frame):
            debug_log(line.rstrip())
    debug_log("stack-dump:end")


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


def open_browser_when_ready(port: int, url: str) -> None:
    if wait_until_ready(port):
        print(f"근무표 생성기 실행 중: {url}", flush=True)
        print("이 창을 닫거나 Ctrl+C를 누르면 서버가 종료됩니다.", flush=True)
        if os.environ.get("WORK_SCHEDULER_NO_BROWSER", "").strip() != "1":
            webbrowser.open(url)
    else:
        print("서버 시작을 확인하지 못했습니다. 잠시 후 브라우저에서 다시 접속해 보세요.", flush=True)
        print(url, flush=True)


def main() -> int:
    multiprocessing.freeze_support()
    debug_log("main:start")
    debug_log("import app.main:start")
    from app.main import app

    debug_log("import app.main:done")
    port = find_available_port()
    url = f"http://{HOST}:{port}/"
    debug_log(f"port:selected:{port}")

    thread = threading.Thread(target=open_browser_when_ready, args=(port, url), daemon=True)
    thread.start()
    debug_log("browser-thread:started")
    threading.Thread(target=dump_stacks_after_delay, args=(15,), daemon=True).start()

    debug_log("uvicorn:start")
    uvicorn.run(
        app,
        host=HOST,
        port=port,
        log_level="warning",
        access_log=False,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

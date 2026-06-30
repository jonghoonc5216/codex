from __future__ import annotations

import socket
import subprocess
import sys
import time
import webbrowser
from pathlib import Path


ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
HOST = "127.0.0.1"
PORT = 8765
URL = f"http://{HOST}:{PORT}"
SERVER_MODE_ARG = "--aerial-viewer-server"


def is_server_running() -> bool:
    try:
        with socket.create_connection((HOST, PORT), timeout=0.4):
            return True
    except OSError:
        return False


def start_server() -> None:
    creationflags = 0
    if hasattr(subprocess, "CREATE_NEW_PROCESS_GROUP"):
        creationflags |= subprocess.CREATE_NEW_PROCESS_GROUP
    if hasattr(subprocess, "DETACHED_PROCESS"):
        creationflags |= subprocess.DETACHED_PROCESS
    if hasattr(subprocess, "CREATE_NO_WINDOW"):
        creationflags |= subprocess.CREATE_NO_WINDOW

    stdout = (ROOT / "server.log").open("ab")
    stderr = (ROOT / "server.err.log").open("ab")
    if getattr(sys, "frozen", False):
        command = [sys.executable, SERVER_MODE_ARG]
    else:
        command = [sys.executable.replace("pythonw.exe", "python.exe"), "server.py", "--host", HOST, "--port", str(PORT)]

    subprocess.Popen(
        command,
        cwd=str(ROOT),
        stdin=subprocess.DEVNULL,
        stdout=stdout,
        stderr=stderr,
        creationflags=creationflags,
        close_fds=True,
    )


def wait_for_server() -> None:
    for _ in range(25):
        if is_server_running():
            return
        time.sleep(0.2)


def main() -> None:
    if SERVER_MODE_ARG in sys.argv:
        import server

        sys.argv = [sys.argv[0], "--host", HOST, "--port", str(PORT)]
        server.main()
        return

    if not is_server_running():
        start_server()
        wait_for_server()
    webbrowser.open(URL)


if __name__ == "__main__":
    main()

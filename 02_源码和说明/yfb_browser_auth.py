from __future__ import annotations

import json
import os
import socket
import subprocess
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Callable
from urllib.request import urlopen

import websocket


LOGIN_URL = "https://qiye.qianlima.com/new_qd_yfbsite/#/infoCenter/search"
AUTH_COOKIE_NAME = "Admin-Token"


@dataclass(frozen=True)
class BrowserInfo:
    name: str
    path: Path
    key: str


def _candidate_paths() -> list[BrowserInfo]:
    program_files = Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
    program_files_x86 = Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
    local_app_data = Path(os.environ.get("LOCALAPPDATA", ""))
    candidates = (

        ("Google Chrome", "chrome", program_files / "Google/Chrome/Application/chrome.exe"),
        ("Google Chrome", "chrome", program_files_x86 / "Google/Chrome/Application/chrome.exe"),
        ("Google Chrome", "chrome", local_app_data / "Google/Chrome/Application/chrome.exe"),
        ("Brave", "brave", program_files / "BraveSoftware/Brave-Browser/Application/brave.exe"),
        ("Brave", "brave", program_files_x86 / "BraveSoftware/Brave-Browser/Application/brave.exe"),
        ("Brave", "brave", local_app_data / "BraveSoftware/Brave-Browser/Application/brave.exe"),
        ("Microsoft Edge", "edge", program_files / "Microsoft/Edge/Application/msedge.exe"),
        ("Microsoft Edge", "edge", program_files_x86 / "Microsoft/Edge/Application/msedge.exe"),
        ("Microsoft Edge", "edge", local_app_data / "Microsoft/Edge/Application/msedge.exe"),
    )
    return [BrowserInfo(name, path, key) for name, key, path in candidates]


def detect_browsers() -> list[BrowserInfo]:
    found: list[BrowserInfo] = []
    seen: set[str] = set()
    for browser in _candidate_paths():
        normalized = str(browser.path).lower()
        if browser.path.is_file() and normalized not in seen:
            found.append(browser)
            seen.add(normalized)
    return found


def browser_from_path(path: str | Path) -> BrowserInfo:
    executable = Path(path).resolve()
    stem = executable.stem.lower()
    if "edge" in stem:
        name, key = "Microsoft Edge", "edge"
    elif "brave" in stem:
        name, key = "Brave", "brave"
    elif "chrome" in stem or "chromium" in stem:
        name, key = "Google Chrome / Chromium", "chrome"
    else:
        name, key = executable.stem, "custom"
    return BrowserInfo(name, executable, key)


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def _read_json(url: str, timeout: float = 1.0) -> dict:
    with urlopen(url, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def _wait_for_debugger(port: int, process: subprocess.Popen, timeout: float = 15.0) -> str:
    deadline = time.monotonic() + timeout
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        if process.poll() is not None:
            raise RuntimeError("浏览器启动后立即退出，请确认所选程序是 Chromium 内核浏览器。")
        try:
            version = _read_json(f"http://127.0.0.1:{port}/json/version")
            ws_url = str(version.get("webSocketDebuggerUrl") or "")
            if ws_url:
                return ws_url
        except Exception as exc:
            last_error = exc
        time.sleep(0.2)
    raise RuntimeError(f"浏览器调试接口启动超时：{last_error or '未知原因'}")


def _cdp_call(ws_url: str, method: str, params: dict | None = None) -> dict:
    connection = websocket.create_connection(
        ws_url,
        timeout=5,
        origin="http://localhost",
        suppress_origin=True,
    )
    try:
        connection.send(json.dumps({"id": 1, "method": method, "params": params or {}}))
        while True:
            message = json.loads(connection.recv())
            if message.get("id") != 1:
                continue
            if "error" in message:
                raise RuntimeError(str(message["error"].get("message") or message["error"]))
            return message.get("result") or {}
    finally:
        connection.close()


def _cookie_header(cookies: list[dict]) -> str:
    domain_cookies = []
    names: set[str] = set()
    for cookie in cookies:
        domain = str(cookie.get("domain") or "").lower().lstrip(".")
        if not (domain == "qianlima.com" or domain.endswith(".qianlima.com")):
            continue
        name = str(cookie.get("name") or "")
        value = str(cookie.get("value") or "")
        if not name or name in names:
            continue
        names.add(name)
        domain_cookies.append(f"{name}={value}")
    return "; ".join(domain_cookies)


def has_auth_cookie(cookie_header: str) -> bool:
    names = {part.split("=", 1)[0].strip() for part in cookie_header.split(";") if "=" in part}
    return AUTH_COOKIE_NAME in names


class BrowserSession:
    def __init__(self, browser: BrowserInfo, data_root: Path, headless: bool) -> None:
        if not browser.path.is_file():
            raise FileNotFoundError(f"浏览器程序不存在：{browser.path}")
        self.browser = browser
        self.profile_dir = data_root / browser.key
        self.profile_dir.mkdir(parents=True, exist_ok=True)
        self.port = _free_port()
        command = [
            str(browser.path),
            f"--remote-debugging-port={self.port}",
            "--remote-allow-origins=*",
            f"--user-data-dir={self.profile_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--disable-background-mode",
        ]
        if headless:
            command.extend(("--headless=new", "--disable-gpu"))
        command.append(LOGIN_URL)
        self.process = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        try:
            self.ws_url = _wait_for_debugger(self.port, self.process)
        except Exception:
            if self.process.poll() is None:
                self.process.terminate()
            raise

    def cookie_header(self) -> str:
        result = _cdp_call(self.ws_url, "Storage.getCookies")
        return _cookie_header(result.get("cookies") or [])

    def close(self) -> None:
        try:
            _cdp_call(self.ws_url, "Browser.close")
        except Exception:
            if self.process.poll() is None:
                self.process.terminate()
        try:
            self.process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            self.process.kill()


def read_saved_cookie(browser: BrowserInfo, data_root: Path) -> str:
    session = BrowserSession(browser, data_root, headless=True)
    try:
        cookie = session.cookie_header()
        return cookie if has_auth_cookie(cookie) else ""
    finally:
        session.close()


def interactive_login(
    browser: BrowserInfo,
    data_root: Path,
    timeout: int = 600,
    status_callback: Callable[[str], None] | None = None,
) -> str:
    session = BrowserSession(browser, data_root, headless=False)
    deadline = time.monotonic() + timeout
    try:
        if status_callback:
            status_callback("浏览器已打开，请在页面中完成乙方宝登录。")
        while time.monotonic() < deadline:
            if session.process.poll() is not None:
                raise RuntimeError("登录窗口已关闭，但尚未检测到有效登录状态。")
            cookie = session.cookie_header()
            if has_auth_cookie(cookie):
                if status_callback:
                    status_callback("已检测到登录状态，正在保存。")
                return cookie
            time.sleep(1)
        raise TimeoutError("等待登录超时，请重新点击“登录乙方宝”。")
    finally:
        session.close()

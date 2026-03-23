"""
사용자의 실제 브라우저(Chrome/Edge)에 연결하는 모듈.
기존 로그인 세션을 재사용하여 별도 로그인 불필요.
"""

import os
import sys
import time
import subprocess
import platform
from pathlib import Path
from typing import Optional

from playwright.sync_api import sync_playwright, Browser, BrowserContext

# 브라우저 실행 파일 경로
CHROME_PATHS = {
    "Windows": [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ],
    "Darwin": [  # macOS
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Chromium.app/Contents/MacOS/Chromium",
    ],
    "Linux": [
        "/usr/bin/google-chrome",
        "/usr/bin/chromium-browser",
        "/usr/bin/chromium",
    ],
}

EDGE_PATHS = {
    "Windows": [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ],
    "Darwin": [
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ],
    "Linux": [
        "/usr/bin/microsoft-edge",
    ],
}

def find_browser(prefer: str = "chrome") -> Optional[str]:
    """설치된 브라우저 실행파일 경로 탐색"""
    system = platform.system()
    candidates = []
    if prefer == "chrome":
        candidates = CHROME_PATHS.get(system, []) + EDGE_PATHS.get(system, [])
    elif prefer == "edge":
        candidates = EDGE_PATHS.get(system, []) + CHROME_PATHS.get(system, [])

    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def get_user_data_dir(browser: str = "chrome") -> str:
    """브라우저 기본 프로파일 경로 (기존 로그인 세션 포함)"""
    system = platform.system()
    if system == "Windows":
        base = os.environ.get("LOCALAPPDATA", "")
        if browser == "chrome":
            return os.path.join(base, "Google", "Chrome", "User Data")
        elif browser == "edge":
            return os.path.join(base, "Microsoft", "Edge", "User Data")
    elif system == "Darwin":
        home = Path.home()
        if browser == "chrome":
            return str(home / "Library/Application Support/Google/Chrome")
        elif browser == "edge":
            return str(home / "Library/Application Support/Microsoft Edge")
    elif system == "Linux":
        home = Path.home()
        if browser == "chrome":
            return str(home / ".config/google-chrome")
        elif browser == "edge":
            return str(home / ".config/microsoft-edge")
    return ""


class BrowserConnector:
    """
    연결 방식 3가지:
      1) CDP (기존 Chrome을 디버그 모드로 연결) - 가장 안정적
      2) UserDataDir (기존 프로파일 복사해서 사용) - 로그인 세션 재사용
      3) Playwright 자체 브라우저 (fallback)
    """

    def __init__(self, cfg: dict):
        self.cfg = cfg.get("browser", {})
        self.mode = self.cfg.get("mode", "auto")       # auto / cdp / profile / playwright
        self.prefer = self.cfg.get("prefer", "chrome") # chrome / edge
        self.debug_port = self.cfg.get("debug_port", 9222)
        self._pw = None
        self._browser = None
        self._context = None

    def connect(self, playwright) -> BrowserContext:
        self._pw = playwright

        if self.mode == "cdp" or (self.mode == "auto"):
            ctx = self._try_cdp(playwright)
            if ctx:
                return ctx

        if self.mode == "profile" or (self.mode == "auto"):
            ctx = self._try_profile(playwright)
            if ctx:
                return ctx

        # fallback: Playwright 자체 Chromium
        print("⚠ 기존 브라우저 연결 실패 → Playwright 내장 Chromium 사용")
        return self._playwright_fallback(playwright)

    # 방식 1: CDP
    def _try_cdp(self, playwright) -> Optional[BrowserContext]:
        """
        사용자가 Chrome을 디버그 모드로 열어 놓으면 거기에 붙음.
        사용자 안내: chrome.exe --remote-debugging-port=9222 --no-first-run
        프로그램이 자동으로 Chrome을 디버그 모드로 열어줄 수도 있음.
        """
        # 먼저 이미 열려 있는지 확인
        import urllib.request
        try:
            urllib.request.urlopen(f"http://localhost:{self.debug_port}/json", timeout=2)
            print(f"✅ 기존 Chrome 디버그 포트({self.debug_port}) 연결 성공")
            browser = playwright.chromium.connect_over_cdp(
                f"http://localhost:{self.debug_port}"
            )
            return browser.contexts[0] if browser.contexts else browser.new_context()
        except Exception:
            pass

        # Chrome을 디버그 모드로 자동 실행 시도
        exe = find_browser(self.prefer)
        if not exe:
            return None

        print(f"🔄 Chrome 디버그 모드로 실행 중... ({exe})")
        try:
            subprocess.Popen([
                exe,
                f"--remote-debugging-port={self.debug_port}",
                "--no-first-run",
                "--no-default-browser-check",
            ])
            time.sleep(3)  # 브라우저 뜨길 기다림

            browser = playwright.chromium.connect_over_cdp(
                f"http://localhost:{self.debug_port}"
            )
            print("✅ Chrome 디버그 모드 연결 성공 (기존 로그인 세션 사용)")
            return browser.contexts[0] if browser.contexts else browser.new_context()
        except Exception as e:
            print(f"  CDP 연결 실패: {e}")
            return None

    # 방식 2: 프로파일 복사
    def _try_profile(self, playwright) -> Optional[BrowserContext]:
        """
        기존 Chrome 프로파일을 임시 복사해서 사용.
        로그인 세션(쿠키)이 그대로 복사되어 재로그인 불필요.
        """
        user_data = self.cfg.get("user_data_dir") or get_user_data_dir(self.prefer)
        if not user_data or not os.path.exists(user_data):
            return None

        exe = find_browser(self.prefer)
        if not exe:
            return None

        try:
            print(f"🔄 기존 Chrome 프로파일로 연결 중...")
            # 주의: Chrome이 이미 열려 있으면 프로파일 잠금 충돌 가능
            # → 별도 임시 프로파일 디렉토리에 복사
            import shutil
            import tempfile
            tmp_dir = tempfile.mkdtemp(prefix="map_reg_")
            profile_src = os.path.join(user_data, "Default")
            profile_dst = os.path.join(tmp_dir, "Default")

            # 쿠키, 로컬스토리지만 복사 (빠름)
            os.makedirs(profile_dst, exist_ok=True)
            for item in ["Cookies", "Local Storage", "Web Data", "Login Data"]:
                src = os.path.join(profile_src, item)
                if os.path.exists(src):
                    dst = os.path.join(profile_dst, item)
                    if os.path.isdir(src):
                        shutil.copytree(src, dst)
                    else:
                        shutil.copy2(src, dst)

            context = playwright.chromium.launch_persistent_context(
                user_data_dir=tmp_dir,
                executable_path=exe,
                headless=False,
                args=["--no-first-run", "--no-default-browser-check"],
            )
            print("✅ 기존 Chrome 프로파일 연결 성공")
            return context
        except Exception as e:
            print(f"  프로파일 연결 실패: {e}")
            return None

    # 방식 3: Playwright Chromium (fallback)
    def _playwright_fallback(self, playwright) -> BrowserContext:
        headless = self.cfg.get("headless", False)
        browser = playwright.chromium.launch(
            headless=headless,
            args=["--start-maximized"],
        )
        return browser.new_context(viewport={"width": 1400, "height": 900})

    def close(self):
        try:
            if self._context:
                self._context.close()
            if self._browser:
                self._browser.close()
        except Exception:
            pass

#!/usr/bin/env python3
"""map-bookmarker GUI launcher"""
import sys, os

# PyInstaller EXE 중복 실행 방지 — 반드시 최상위에서 호출
if getattr(sys, 'frozen', False):
    import multiprocessing
    multiprocessing.freeze_support()
    if multiprocessing.parent_process() is not None:
        sys.exit(0)

import glob, subprocess

# ── 브라우저 경로 통일 ──
# EXE 환경에서 playwright가 찾는 경로와 설치 경로를 일치시킴
_BROWSERS_PATH = os.path.join(os.path.expanduser("~"), "AppData", "Local", "ms-playwright")
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _BROWSERS_PATH


def _find_chromium():
    """Playwright Chromium 실행파일이 있는지 확인"""
    for pattern in [
        os.path.join(_BROWSERS_PATH, "chromium-*", "chrome-win*", "chrome.exe"),
        os.path.join(_BROWSERS_PATH, "chromium-*", "chrome-linux", "chrome"),
    ]:
        if glob.glob(pattern):
            return True
    return False


def ensure_playwright():
    """Playwright Chromium이 없으면 자동 설치"""
    if _find_chromium():
        return

    print("Chromium 브라우저가 없습니다. 설치합니다... (1~2분)")
    try:
        from playwright._impl._driver import compute_driver_executable
        driver_info = compute_driver_executable()

        if isinstance(driver_info, tuple):
            node_exe, cli_js = driver_info
            cmd = [str(node_exe), str(cli_js), "install", "chromium"]
        else:
            cmd = [str(driver_info), "install", "chromium"]

        # 환경변수에 경로 전달
        env = os.environ.copy()
        env["PLAYWRIGHT_BROWSERS_PATH"] = _BROWSERS_PATH

        print(f"설치 경로: {_BROWSERS_PATH}")
        result = subprocess.run(cmd, env=env, timeout=300)

        if result.returncode == 0:
            print("Chromium 설치 완료!")
        else:
            print("자동 설치 실패. 수동: python -m playwright install chromium")
    except Exception as e:
        print(f"설치 오류: {e}")
        if not getattr(sys, 'frozen', False):
            try:
                subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], timeout=300)
            except Exception:
                pass
        print("수동 설치 필요: python -m playwright install chromium")


if __name__ == "__main__":
    ensure_playwright()
    from src.gui import App
    App().run()

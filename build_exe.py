"""
build_exe.py
─────────────
PyInstaller로 단일 실행파일(.exe / macOS app) 생성
의존성 전부 포함되어 사용자는 Python 설치 불필요

실행:
  python build_exe.py
"""

import os
import sys
import shutil
import subprocess
import platform

APP_NAME = "map-favorite-registrar"
ENTRY   = "src/main.py"
ICON    = "assets/icon.ico"   # 없으면 기본 아이콘 사용

def build():
    system = platform.system()
    print(f"🔨 {system} 빌드 시작...")

    # PyInstaller 설치 확인
    try:
        import PyInstaller
    except ImportError:
        print("PyInstaller 설치 중...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])

    # Playwright 브라우저 바이너리 경로 확인
    # Playwright를 포함하면 용량이 너무 커지므로
    # → 브라우저는 사용자 시스템 것을 쓰는 방식(CDP/Profile)이 기본
    # → fallback용 Playwright Chromium만 선택적으로 포함
    playwright_driver = _find_playwright_driver()

    args = [
        "pyinstaller",
        "--onefile",            # 단일 exe
        "--name", APP_NAME,
        "--clean",
        "--noconfirm",
    ]

    if system == "Windows":
        args += ["--console"]   # 진행상황 보이도록 콘솔 창 유지
        if os.path.exists(ICON):
            args += ["--icon", ICON]
    elif system == "Darwin":
        args += ["--windowed"]  # macOS는 앱 번들

    # 설정 파일 포함
    args += ["--add-data", f"config{os.pathsep}config"]

    # Playwright driver 포함 (fallback용)
    if playwright_driver:
        args += ["--add-binary", f"{playwright_driver}{os.pathsep}playwright/driver"]

    args.append(ENTRY)

    print(f"실행: {' '.join(args)}\n")
    subprocess.check_call(args)

    # 결과물 정리
    dist_dir = os.path.join("dist", APP_NAME)
    if system == "Windows":
        dist_dir += ".exe"

    if os.path.exists(dist_dir):
        print(f"\n✅ 빌드 완료: {dist_dir}")
        print(f"   파일크기: {os.path.getsize(dist_dir) / 1024 / 1024:.1f} MB")
    else:
        print(f"\n✅ 빌드 완료: dist/ 폴더 확인")

    _write_usage_guide(system)


def _find_playwright_driver() -> str:
    """Playwright 드라이버 바이너리 경로"""
    try:
        import playwright
        driver_dir = os.path.join(os.path.dirname(playwright.__file__), "driver")
        if os.path.exists(driver_dir):
            return driver_dir
    except Exception:
        pass
    return ""


def _write_usage_guide(system: str):
    """배포용 사용 가이드 텍스트 생성"""
    exe_name = f"{APP_NAME}.exe" if system == "Windows" else APP_NAME
    guide = f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  map-favorite-registrar 사용 가이드
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

1. config/config.yaml 파일을 텍스트 편집기로 엽니다

2. 아래 3가지만 수정합니다:

   input:
     file: "data/내파일.xlsx"   ← 내 엑셀 파일 경로

   columns:
     name:    "이름컬럼"        ← 내 파일의 이름 컬럼명
     address: "주소컬럼"        ← 내 파일의 주소 컬럼명

3. Chrome이 설치되어 있으면 자동으로 연결됩니다
   (기존 카카오맵/네이버지도 로그인 세션 자동 사용)

   로그인이 안 되어 있다면 config.yaml에 계정 정보 입력:
   kakao:
     id: "내_카카오_이메일"
     password: "비밀번호"

4. {exe_name} 실행 (더블클릭 또는 명령 프롬프트에서 실행)

   옵션:
     --dry-run          : 실제 등록 없이 미리보기
     --kakao-only       : 카카오맵만
     --naver-only       : 네이버지도만
     --limit 10         : 처음 10개만 테스트

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""
    with open("dist/사용_가이드.txt", "w", encoding="utf-8") as f:
        f.write(guide)
    print(guide)


if __name__ == "__main__":
    build()

"""
map-bookmarker - 지도 즐겨찾기 자동 등록
https://github.com/HarimxChoi/map-bookmarker
"""

import sys
import os

# EXE 환경에서 Playwright 브라우저 경로 설정 (import 전에 반드시 설정)
if getattr(sys, 'frozen', False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        os.path.expanduser("~"), ".playwright-browsers"
    )

import json
import time
import logging
import argparse
from pathlib import Path
from typing import Optional

import yaml
import openpyxl
import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# 로거 설정
def setup_logger(log_file: str) -> logging.Logger:
    Path(log_file).parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("map-reg")
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    sh = logging.StreamHandler(sys.stdout)
    # Windows cp949 콘솔에서 UTF-8 이모지 깨짐 방지
    if hasattr(sh.stream, 'reconfigure'):
        try:
            sh.stream.reconfigure(encoding='utf-8')
        except Exception:
            pass
    sh.setFormatter(fmt)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(sh)
    logger.addHandler(fh)
    return logger

# 설정 로드
def load_config(path: str) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)

# 데이터 로드 & 필터링
def load_data(cfg: dict) -> list[dict]:
    fi = cfg["input"]
    path = fi["file"]

    if path.endswith(".csv"):
        df = pd.read_csv(path, encoding="utf-8-sig")
    else:
        df = pd.read_excel(path, sheet_name=fi.get("sheet", 0),
                           header=fi.get("header_row", 1) - 1)

    # 필터 적용
    for f in cfg.get("filters", []):
        col = f["column"]
        if col not in df.columns:
            print(f"  ⚠ 필터 컬럼 '{col}' 없음, 건너뜀")
            continue
        if "not_contains" in f:
            mask = ~df[col].astype(str).str.contains("|".join(f["not_contains"]), na=False)
            df = df[mask]
        if "contains" in f:
            mask = df[col].astype(str).str.contains("|".join(f["contains"]), na=False)
            df = df[mask]
        if "min" in f:
            df = df[pd.to_numeric(df[col], errors="coerce") >= f["min"]]
        if "max" in f:
            df = df[pd.to_numeric(df[col], errors="coerce") <= f["max"]]

    cols = cfg["columns"]
    results = []
    for _, row in df.iterrows():
        def get(key):
            c = cols.get(key)
            return str(row[c]).strip() if c and c in df.columns and pd.notna(row[c]) else ""

        name_fmt = cfg.get("bookmark_name", "{이름}")
        memo_fmt = cfg.get("bookmark_memo", "")

        # {컬럼명} 치환
        name = name_fmt
        memo = memo_fmt
        for col in df.columns:
            val = str(row[col]) if pd.notna(row[col]) else ""
            name = name.replace(f"{{{col}}}", val)
            memo = memo.replace(f"{{{col}}}", val)

        addr = get("address")
        if not addr:
            continue

        # 동/호수/층수 자동 추출하여 즐겨찾기명에 추가
        if cfg.get("append_unit_info", False):
            import re
            unit_parts = []

            # 쉼표 뒤 상세주소 부분 분리
            after_comma = ""
            if "," in addr:
                after_comma = addr.split(",", 1)[1].strip()

            # "101-501" 패턴 (쉼표 뒤) → 101동 501호로 해석
            dash_match = re.search(r'(\d+)-(\d+)', after_comma) if after_comma else None
            if dash_match:
                unit_parts.append(f"{dash_match.group(1)}동")
                unit_parts.append(f"{dash_match.group(2)}호")
            else:
                # "106동" 패턴
                dong_match = re.search(r'(\d+동)', addr)
                if dong_match:
                    unit_parts.append(dong_match.group(1))
                # "1102호" 패턴
                ho_match = re.search(r'(\d+호)', addr)
                if ho_match:
                    unit_parts.append(ho_match.group(1))

            # "16층", "3층" 패턴
            floor_match = re.search(r'(\d+층)', addr)
            if floor_match:
                unit_parts.append(floor_match.group(1))

            if unit_parts:
                name = f"{name} ({' '.join(unit_parts)})"

        results.append({
            "name": name,
            "address": addr,
            "label": get("label"),
            "memo": memo,
        })

    return results

# 진행상태 저장/불러오기 (resume 지원)
class Progress:
    def __init__(self, path: str):
        self.path = path
        self.done: set = set()
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                data = json.load(f)
                self.done = set(data.get("done", []))

    def mark(self, key: str, platform: str):
        self.done.add(f"{platform}:{key}")
        with open(self.path, "w", encoding="utf-8") as f:
            json.dump({"done": list(self.done)}, f, ensure_ascii=False)

    def is_done(self, key: str, platform: str) -> bool:
        return f"{platform}:{key}" in self.done

# 카카오맵 자동화
class KakaoMapRegistrar:
    def __init__(self, cfg: dict, logger: logging.Logger):
        self.cfg = cfg["kakao"]
        self.opts = cfg["options"]
        self.logger = logger
        self.page = None
        self.logged_in = False

    def login(self, page):
        self.logger.info("카카오맵 로그인 중...")
        # 로그인 페이지로 직접 이동 (map.kakao.com 경유하지 않음)
        page.goto("https://accounts.kakao.com/login?continue=https://map.kakao.com/")
        time.sleep(2)

        page.fill("input[name='loginId']", self.cfg["id"])
        time.sleep(1.5)
        page.fill("input[name='password']", self.cfg["password"])
        time.sleep(2)
        page.click("button.submit")
        time.sleep(3)

        # 2차 인증 대기: map.kakao.com으로 이동할 때까지 대기
        # 현재 페이지 URL, 새 탭, 브라우저 주소창 모두 확인
        self.logger.info("📱 카카오톡 알림을 확인하고 핸드폰에서 로그인을 승인해주세요... (최대 5분 대기)")
        context = page.context
        deadline = time.time() + 300
        target_page = None

        while time.time() < deadline:
            # 1) 현재 페이지 URL 확인
            url = page.url
            if url.startswith("https://map.kakao.com") or url.startswith("http://map.kakao.com"):
                target_page = page
                break

            # 2) 새 탭/팝업이 열렸는지 확인
            for p in context.pages:
                purl = p.url
                if purl.startswith("https://map.kakao.com") or purl.startswith("http://map.kakao.com"):
                    target_page = p
                    break
            if target_page:
                break

            # 3) JavaScript로 실제 location.href 확인 (URL 스푸핑 대응)
            try:
                real_url = page.evaluate("window.location.href")
                if "map.kakao.com" in real_url and "accounts.kakao.com" not in real_url:
                    target_page = page
                    break
            except Exception:
                pass

            time.sleep(3)
        else:
            raise Exception("카카오 로그인 타임아웃 — 5분 내에 인증을 완료하지 않았거나 로그인 정보가 잘못됨")

        # 새 탭이 열린 경우 해당 페이지를 사용
        if target_page and target_page != page:
            page = target_page

        time.sleep(2)
        self.logged_in = True
        self._login_page = page  # register에서 사용할 페이지 참조
        self.logger.info("✅ 카카오맵 로그인 성공")

    def _dismiss_overlays(self, page):
        """DimmedLayer, 코치 레이어, 약관 동의 등 오버레이 처리"""
        # DimmedLayer 숨기기
        page.evaluate('document.querySelector("#dimmedLayer")?.style?.setProperty("display","none")')
        # 코치 레이어 닫기
        for sel in [".coach_layer .btn_close", ".coach_layer button"]:
            btns = page.locator(sel)
            for i in range(btns.count()):
                try:
                    if btns.nth(i).is_visible():
                        btns.nth(i).click(force=True)
                        time.sleep(0.3)
                except Exception:
                    pass
        # 내 장소 약관 동의
        agree_btn = page.locator(".terms_myplace_layer button:has-text('동의')")
        try:
            if agree_btn.count() > 0 and agree_btn.first.is_visible():
                agree_btn.first.click(force=True)
                time.sleep(1)
        except Exception:
            pass

    @staticmethod
    def _clean_address(addr: str, level: int = 1) -> str:
        """검색에 불필요한 상세주소 제거. level이 높을수록 더 많이 제거"""
        import re
        # level 1: 기본 정리
        # 쉼표 뒤의 동·호 정보 제거: "내정로 186, 106동 1102호 (수내동,...)" → "내정로 186"
        addr = re.split(r',\s*\d+동', addr)[0]
        # 괄호 안 내용 제거
        addr = re.sub(r'\([^)]*\)', '', addr)

        if level >= 2:
            # level 2: 붙어있는 동호수 패턴도 제거 (아파트명101동502호 → 아파트명)
            addr = re.sub(r'\S*\d+동\d*호?\S*', '', addr)
            # 중복 단어 제거 (샛별마을라이프아파트 샛별마을라이프아파트 → 하나만)
            words = addr.split()
            seen = set()
            deduped = []
            for w in words:
                if w not in seen:
                    seen.add(w)
                    deduped.append(w)
            addr = ' '.join(deduped)

        addr = re.sub(r'\s+', ' ', addr).strip()
        return addr

    def register(self, page, item: dict) -> bool:
        """주소 검색 → 즐겨찾기 저장. 실패 시 더 강한 주소 정리로 자동 재시도 (재시도 횟수 무관)"""
        # level 1로 시도, 실패하면 level 2로 재시도
        for level in [1, 2]:
            result = self._try_register(page, item, level)
            if result:
                return True
            if level == 1:
                self.logger.info(f"    🔄 주소 재정리(level 2)로 재검색...")
        return False

    def _try_register(self, page, item: dict, clean_level: int) -> bool:
        """주소 검색 → 즐겨찾기 저장 (1회 시도)"""
        addr = item["address"]
        name = item["name"]
        delay = self.opts.get("delay_ms", 800) / 1000

        try:
            # 카카오맵 메인이 아니면 이동
            if "map.kakao.com" not in page.url:
                page.goto("https://map.kakao.com/")
                time.sleep(2)

            # 오버레이 제거
            self._dismiss_overlays(page)

            search_sel = '[id="search.keyword.query"]'
            # 검색창 로드 대기
            page.wait_for_selector(search_sel, timeout=10000)

            # 상세주소 정리
            clean_addr = self._clean_address(addr, level=clean_level)

            page.fill(search_sel, "")
            page.fill(search_sel, clean_addr)
            page.press(search_sel, "Enter")
            time.sleep(delay + 2)

            # 오버레이가 다시 나타날 수 있으므로 재처리
            self._dismiss_overlays(page)

            # 1) 장소 탭 결과 우선 시도
            try:
                page.wait_for_selector(".placelist .PlaceItem", timeout=6000)
                items = page.locator(".placelist .PlaceItem")
                if items.count() > 0:
                    first = items.nth(0)
                    btn = first.locator('[data-id="fav"]')
                    if btn.count() > 0:
                        btn.first.click(force=True)
                        time.sleep(delay + 1)
                        self._dismiss_overlays(page)
                        self._handle_save_popup(page, name)
                        self.logger.info(f"  ✅ 카카오맵 등록(장소): {name} | {addr}")
                        return True
            except PWTimeout:
                pass

            # 2) 주소 탭 결과 → 클릭해서 InfoWindow 열기 → 거기서 fav 클릭
            try:
                page.wait_for_selector(".addrlist li", timeout=4000)
                addr_items = page.locator(".addrlist li")
                if addr_items.count() > 0:
                    # 주소 리스트 첫 번째 항목 클릭 → 지도에 InfoWindow 팝업 열림
                    first_addr = addr_items.nth(0)
                    first_addr.click()
                    time.sleep(delay + 1)

                    # InfoWindow 안의 즐겨찾기 버튼 클릭
                    fav_btn = page.locator('.InfoWindow [data-id="fav"], .AddressInfoWindow [data-id="fav"]')
                    if fav_btn.count() > 0 and fav_btn.first.is_visible(timeout=3000):
                        fav_btn.first.click(force=True)
                        time.sleep(delay + 1)
                        self._dismiss_overlays(page)
                        self._handle_save_popup(page, name)
                        self.logger.info(f"  ✅ 카카오맵 등록(주소): {name} | {addr}")
                        return True
            except PWTimeout:
                pass

            # 3) 페이지 어디서든 보이는 fav 버튼 시도 (link_fav 포함)
            for sel in ['a[data-id="fav"].link_fav', 'a[data-id="fav"].fav',
                        '[data-id="fav"]']:
                btn = page.locator(sel)
                try:
                    if btn.count() > 0 and btn.first.is_visible(timeout=2000):
                        btn.first.click(force=True)
                        time.sleep(delay + 1)
                        self._dismiss_overlays(page)
                        self._handle_save_popup(page, name)
                        self.logger.info(f"  ✅ 카카오맵 등록(범용): {name} | {addr}")
                        return True
                except Exception:
                    pass

            if clean_level >= 2:
                self.logger.warning(f"  ❌ 검색결과/즐겨찾기 버튼 없음: {addr}")
            return False

        except Exception as e:
            if clean_level >= 2:
                self.logger.error(f"  ❌ 카카오맵 오류 ({addr}): {e}")
            return False

    def _handle_save_popup(self, page, name: str):
        """즐겨찾기 저장 팝업 처리: 폴더 선택/생성 → 이름 수정 → 완료
        동일 주소가 이미 폴더에 있으면(SAVED) 중복N 폴더로 자동 분산"""
        try:
            time.sleep(1)
            base_folder = self.cfg.get("folder", "")

            if base_folder:
                # 시도할 폴더 순서: base → base 중복1 → base 중복2 → ...
                folder_candidates = [base_folder]
                for n in range(1, 50):
                    folder_candidates.append(f"{base_folder} 중복{n}")

                selected = False
                for folder in folder_candidates:
                    # 정확한 폴더명 매칭 (JS로 텍스트 완전 일치 확인)
                    match_js = f'''
                        (() => {{
                            const els = document.querySelectorAll("strong.txt_folder");
                            for (const el of els) {{
                                if (el.textContent.trim() === "{folder}") return true;
                            }}
                            return false;
                        }})()
                    '''
                    folder_exists = page.evaluate(match_js)

                    if folder_exists:
                        # 정확히 일치하는 폴더의 li 요소에서 SAVED 확인
                        saved_js = f'''
                            (() => {{
                                const els = document.querySelectorAll("strong.txt_folder");
                                for (const el of els) {{
                                    if (el.textContent.trim() === "{folder}") {{
                                        const li = el.closest("li");
                                        return li ? li.classList.contains("SAVED") : false;
                                    }}
                                }}
                                return false;
                            }})()
                        '''
                        is_saved = page.evaluate(saved_js)

                        if is_saved:
                            self.logger.info(f"    📁 '{folder}' → 이미 등록된 주소, 다음 폴더 시도")
                            continue

                        # SAVED가 아니면 이 폴더 클릭
                        click_js = f'''
                            (() => {{
                                const els = document.querySelectorAll("strong.txt_folder");
                                for (const el of els) {{
                                    if (el.textContent.trim() === "{folder}") {{
                                        const link = el.closest("a.link_folder");
                                        if (link) {{ link.click(); return true; }}
                                        el.click(); return true;
                                    }}
                                }}
                                return false;
                            }})()
                        '''
                        page.evaluate(click_js)
                        self.logger.info(f"    📁 기존 폴더 선택: {folder}")
                        time.sleep(0.5)
                        selected = True
                        break
                    else:
                        # 폴더가 없으면 새로 생성
                        self.logger.info(f"    📁 폴더 '{folder}' 없음 → 새로 생성")
                        add_folder_btn = page.locator("span.ico_folder.add")
                        if add_folder_btn.count() > 0 and add_folder_btn.first.is_visible(timeout=2000):
                            add_folder_btn.first.click(force=True)
                            time.sleep(0.5)

                            folder_input = page.locator("#folderName")
                            folder_input.wait_for(state="visible", timeout=3000)
                            page.evaluate('document.querySelector("#folderName")?.removeAttribute("readonly")')
                            folder_input.fill(folder)
                            time.sleep(0.3)

                            page.locator('button[data-id="addFolderOK"]').click(force=True)
                            self.logger.info(f"    ✅ 폴더 생성 완료: {folder}")
                            time.sleep(0.5)
                        selected = True
                        break

                if not selected:
                    self.logger.warning(f"    ⚠ 사용 가능한 폴더를 찾지 못함")

            # 3) 즐겨찾기 이름 수정 (#display1)
            name_input = page.locator("#display1")
            if name_input.count() > 0:
                try:
                    name_input.wait_for(state="visible", timeout=3000)
                    # readonly 속성 제거 후 입력
                    page.evaluate('document.querySelector("#display1")?.removeAttribute("readonly")')
                    name_input.fill("")
                    name_input.fill(name)
                    time.sleep(0.3)
                except Exception:
                    pass

            # 4) 최종 완료 버튼 클릭
            ok_btn = page.locator('button[data-id="addOK"]')
            if ok_btn.count() > 0 and ok_btn.first.is_visible(timeout=2000):
                ok_btn.first.click(force=True)
                time.sleep(0.5)
            else:
                # fallback: 일반 완료/저장 버튼
                for sel in ['button.btn_submit:has-text("완료")', 'button:has-text("저장")', 'button:has-text("확인")']:
                    btn = page.locator(sel)
                    if btn.count() > 0 and btn.first.is_visible():
                        btn.first.click(force=True)
                        time.sleep(0.5)
                        break

        except Exception as e:
            self.logger.warning(f"    ⚠ 저장 팝업 처리 중 오류: {e}")


# 네이버지도 자동화
class NaverMapRegistrar:
    def __init__(self, cfg: dict, logger: logging.Logger):
        self.cfg = cfg["naver"]
        self.opts = cfg["options"]
        self.logger = logger

    def login(self, page):
        self.logger.info("네이버지도 로그인 중...")
        page.goto("https://nid.naver.com/nidlogin.login?url=https://map.naver.com/")
        time.sleep(1)

        page.fill("#id", self.cfg["id"])
        time.sleep(1.5)
        page.fill("#pw", self.cfg["password"])
        time.sleep(2)
        page.click("button.btn_login, .btn_global[type='submit']", timeout=5000)
        time.sleep(3)

        # 로그인 후 처리 (캡차, 2차 인증, 기기 등록 등)
        # 캡차/2차인증이 있으면 사용자가 직접 처리할 때까지 대기 (최대 5분)
        if "nid.naver.com" in page.url:
            # 캡차 감지
            captcha = page.locator("#captcha, .captcha_wrap")
            if captcha.count() > 0:
                self.logger.info("🔒 캡차가 감지되었습니다! 브라우저에서 직접 캡차를 풀고 로그인을 완료해주세요... (최대 5분 대기)")
            else:
                self.logger.info("📱 네이버 2차 인증을 확인해주세요... (최대 5분 대기)")

            deadline = time.time() + 300  # 5분 대기
            logged_in = False
            while time.time() < deadline:
                current_url = page.url

                # 이미 map.naver.com으로 이동했으면 성공
                if "map.naver.com" in current_url:
                    logged_in = True
                    break

                # 기기 등록 화면 → "등록안함" 클릭
                try:
                    dontsave = page.locator("#new\\.dontsave")
                    if dontsave.count() > 0 and dontsave.first.is_visible(timeout=500):
                        self.logger.info("  📱 기기 등록 화면 → '등록안함' 클릭")
                        dontsave.first.click(force=True)
                        time.sleep(2)
                        continue
                except Exception:
                    pass

                time.sleep(3)

            if not logged_in:
                if "map.naver.com" not in page.url:
                    raise Exception("네이버 로그인 타임아웃 — 캡차/2차 인증을 완료하지 않았거나 로그인 정보가 잘못됨")

        time.sleep(1)
        self.logger.info("✅ 네이버지도 로그인 성공")

    def register(self, page, item: dict) -> bool:
        addr = item["address"]
        name = item["name"]
        delay = self.opts.get("delay_ms", 800) / 1000

        try:
            # 매번 네이버지도 메인으로 이동 (검색 상태 초기화)
            page.goto("https://map.naver.com/")
            time.sleep(3)

            # 1) 검색
            search_input = page.locator("input.input_search")
            search_input.wait_for(state="visible", timeout=15000)
            search_input.fill("")
            search_input.fill(addr)
            search_input.press("Enter")
            time.sleep(delay + 2)

            # 2) 검색결과 iframe 진입 & 첫 번째 항목 클릭
            # 네이버지도는 검색결과가 iframe(#searchIframe) 안에 렌더링됨
            search_frame = None
            for frame in page.frames:
                if "search" in frame.url:
                    search_frame = frame
                    break

            if search_frame:
                # 검색 결과 첫 번째 항목 클릭 (상세 페이지 진입)
                try:
                    first_item = search_frame.locator("li.UEzoS a, li[data-laim-exp-id] a, .Ryr1F a").first
                    first_item.wait_for(state="visible", timeout=8000)
                    first_item.click()
                    time.sleep(delay + 1)
                except PWTimeout:
                    self.logger.warning(f"  ❌ 검색결과 없음: {addr}")
                    return False

            # 3) 상세 페이지 iframe에서 즐겨찾기 버튼 클릭
            entry_frame = None
            for frame in page.frames:
                if "place" in frame.url or "entry" in frame.url:
                    entry_frame = frame
                    break

            target = entry_frame if entry_frame else page

            fav_btn = target.locator("button.btn_favorite, button.btn.btn_favorite")
            try:
                fav_btn.first.wait_for(state="visible", timeout=8000)
                fav_btn.first.click()
                time.sleep(delay + 1)
            except PWTimeout:
                self.logger.warning(f"  ❌ 즐겨찾기 버튼 없음: {addr}")
                return False

            # 4) 폴더 선택/생성 팝업 처리
            self._handle_naver_save_popup(page, name)

            self.logger.info(f"  ✅ 네이버지도 등록: {name} | {addr}")
            return True

        except Exception as e:
            self.logger.error(f"  ❌ 네이버지도 오류 ({addr}): {e}")
            return False

    def _handle_naver_save_popup(self, page, name: str):
        """네이버지도 즐겨찾기 저장 팝업: 폴더 선택/생성 → 저장"""
        try:
            time.sleep(1)
            folder = self.cfg.get("folder", "")

            if folder:
                # 기존 폴더 찾기 (strong.swt-save-group-name 텍스트 매칭)
                existing = page.locator(f"strong.swt-save-group-name:has-text('{folder}')")
                if existing.count() > 0:
                    try:
                        if existing.first.is_visible(timeout=2000):
                            # 부모 button.swt-save-group-info 클릭
                            parent_btn = existing.first.locator("xpath=ancestor::button[contains(@class,'swt-save-group-info')]")
                            if parent_btn.count() > 0:
                                parent_btn.first.click(force=True)
                            else:
                                existing.first.click(force=True)
                            self.logger.info(f"    📁 기존 폴더 선택: {folder}")
                            time.sleep(0.5)
                    except Exception:
                        pass
                else:
                    # 폴더 새로 생성
                    self.logger.info(f"    📁 폴더 '{folder}' 없음 → 새로 생성")
                    try:
                        # "새 리스트 만들기" 버튼 클릭
                        add_btn = page.locator("button.swt-save-group-add-btn")
                        add_btn.first.wait_for(state="visible", timeout=3000)
                        add_btn.first.click(force=True)
                        time.sleep(0.5)

                        # 폴더명 입력
                        folder_input = page.locator("#swt-save-input-folderview-list")
                        folder_input.wait_for(state="visible", timeout=3000)
                        folder_input.fill(folder)
                        time.sleep(0.3)

                        # 완료 버튼으로 폴더 생성
                        page.locator("button.swt-complete-btn").click(force=True)
                        self.logger.info(f"    ✅ 폴더 생성 완료: {folder}")
                        time.sleep(0.5)
                    except Exception as e:
                        self.logger.warning(f"    ⚠ 폴더 생성 실패: {e}")

            # 저장 버튼 클릭
            save_btn = page.locator("button.swt-save-btn")
            try:
                if save_btn.count() > 0 and save_btn.first.is_visible(timeout=2000):
                    save_btn.first.click(force=True)
                    time.sleep(0.5)
            except Exception:
                pass

        except Exception as e:
            self.logger.warning(f"    ⚠ 네이버 저장 팝업 처리 중 오류: {e}")

# Chromium 실행파일 탐색 (다중 경로 fallback)
def _find_chromium_executable() -> Optional[str]:
    """여러 경로에서 Chromium 실행파일을 찾아 반환. 못 찾으면 None (Playwright 기본 사용)"""
    import glob as _glob
    home = os.path.expanduser("~")
    search_bases = [
        os.environ.get("PLAYWRIGHT_BROWSERS_PATH", ""),
        os.path.join(home, "AppData", "Local", "ms-playwright"),
        os.path.join(home, ".cache", "ms-playwright"),
        os.path.join(home, ".playwright-browsers"),
    ]
    for base in search_bases:
        if not base or not os.path.exists(base):
            continue
        for pattern in [
            os.path.join(base, "chromium-*", "chrome-win*", "chrome.exe"),
            os.path.join(base, "chromium-*", "chrome-linux", "chrome"),
        ]:
            found = _glob.glob(pattern)
            if found:
                return found[0]
    return None

# 등록 실행 (GUI/CLI 공용)
def run_registration(cfg, logger, progress, items, use_kakao, use_naver,
                     on_progress=None, stop_event=None):
    """
    Playwright 브라우저를 열고 카카오맵/네이버지도 즐겨찾기를 등록한다.
    on_progress(platform, status, item, stats): 각 항목 처리 후 콜백
    stop_event: threading.Event — set 되면 루프 중단
    """
    stats = {"kakao": {"ok": 0, "fail": 0, "skip": 0},
             "naver": {"ok": 0, "fail": 0, "skip": 0}}

    # 중복 주소 감지
    from collections import Counter
    addr_counts = Counter(item["address"] for item in items)
    duplicates = {addr: cnt for addr, cnt in addr_counts.items() if cnt > 1}
    if duplicates:
        logger.info(f"\n⚠ 동일 주소 {len(duplicates)}건 감지 (중복 폴더로 자동 분산됩니다)")
        for addr, cnt in list(duplicates.items())[:10]:
            names = [it["name"] for it in items if it["address"] == addr]
            logger.info(f"  [{cnt}회] {addr[:50]} → {', '.join(names[:5])}")
        if len(duplicates) > 10:
            logger.info(f"  ... 외 {len(duplicates) - 10}건")

    with sync_playwright() as p:
        # Chromium 실행파일 경로 탐색 (여러 경로 fallback)
        chrome_path = _find_chromium_executable()
        launch_opts = {
            "headless": cfg["options"].get("headless", False),
            "args": ["--start-maximized"],
        }
        if chrome_path:
            launch_opts["executable_path"] = chrome_path
        browser = p.chromium.launch(**launch_opts)
        context = browser.new_context(viewport={"width": 1400, "height": 900})

        # 카카오맵
        if use_kakao:
            page = context.new_page()
            reg = KakaoMapRegistrar(cfg, logger)
            try:
                reg.login(page)
                # 로그인 후 새 탭으로 이동한 경우 해당 페이지 사용
                if hasattr(reg, '_login_page') and reg._login_page != page:
                    page.close()
                    page = reg._login_page
                logger.info(f"\n🗺  카카오맵 즐겨찾기 등록 시작 ({len(items)}개)")
                for i, item in enumerate(items, 1):
                    if stop_event and stop_event.is_set():
                        logger.info("⏹ 사용자 중단 요청")
                        break
                    key = item["address"]
                    if cfg["options"].get("resume") and progress.is_done(key, "kakao"):
                        logger.info(f"  [{i}/{len(items)}] ⏭ 건너뜀(이미등록): {item['name']}")
                        stats["kakao"]["skip"] += 1
                        if on_progress:
                            on_progress("kakao", "skip", item, stats)
                        continue
                    logger.info(f"  [{i}/{len(items)}] {item['name']}")
                    ok = reg.register(page, item)
                    if ok:
                        stats["kakao"]["ok"] += 1
                        progress.mark(key, "kakao")
                    else:
                        stats["kakao"]["fail"] += 1
                    if on_progress:
                        on_progress("kakao", "ok" if ok else "fail", item, stats)
                    time.sleep(cfg["options"].get("delay_ms", 800) / 1000)
            except Exception as e:
                logger.error(f"카카오맵 오류: {e}")
            finally:
                page.close()

        # 네이버지도
        if use_naver:
            if stop_event and stop_event.is_set():
                browser.close()
                return stats
            page = context.new_page()
            reg = NaverMapRegistrar(cfg, logger)
            try:
                reg.login(page)
                logger.info(f"\n🗺  네이버지도 즐겨찾기 등록 시작 ({len(items)}개)")
                for i, item in enumerate(items, 1):
                    if stop_event and stop_event.is_set():
                        logger.info("⏹ 사용자 중단 요청")
                        break
                    key = item["address"]
                    if cfg["options"].get("resume") and progress.is_done(key, "naver"):
                        logger.info(f"  [{i}/{len(items)}] ⏭ 건너뜀(이미등록): {item['name']}")
                        stats["naver"]["skip"] += 1
                        if on_progress:
                            on_progress("naver", "skip", item, stats)
                        continue
                    logger.info(f"  [{i}/{len(items)}] {item['name']}")
                    ok = reg.register(page, item)
                    if ok:
                        stats["naver"]["ok"] += 1
                        progress.mark(key, "naver")
                    else:
                        stats["naver"]["fail"] += 1
                    if on_progress:
                        on_progress("naver", "ok" if ok else "fail", item, stats)
                    time.sleep(cfg["options"].get("delay_ms", 800) / 1000)
            except Exception as e:
                logger.error(f"네이버지도 오류: {e}")
            finally:
                page.close()

        browser.close()

    # 결과 요약
    logger.info("\n" + "=" * 50)
    logger.info("📊 최종 결과")
    if use_kakao:
        s = stats["kakao"]
        logger.info(f"  카카오맵: ✅{s['ok']} ❌{s['fail']} ⏭{s['skip']}")
    if use_naver:
        s = stats["naver"]
        logger.info(f"  네이버지도: ✅{s['ok']} ❌{s['fail']} ⏭{s['skip']}")
    if duplicates:
        logger.info(f"\n📋 동일 주소 리포트 ({len(duplicates)}건)")
        for addr, cnt in duplicates.items():
            names = [it["name"] for it in items if it["address"] == addr]
            logger.info(f"  [{cnt}회] {', '.join(names)} → {addr[:60]}")
    logger.info("=" * 50)
    return stats


# 메인 실행 (CLI)
def main():
    parser = argparse.ArgumentParser(
        description="엑셀/CSV 주소 목록을 카카오맵·네이버지도 즐겨찾기에 자동 등록"
    )
    parser.add_argument("-c", "--config", default="config/config.yaml", help="설정 파일 경로")
    parser.add_argument("--kakao-only", action="store_true", help="카카오맵만 실행")
    parser.add_argument("--naver-only", action="store_true", help="네이버지도만 실행")
    parser.add_argument("--dry-run", action="store_true", help="실제 등록 없이 데이터만 확인")
    parser.add_argument("--limit", type=int, default=0, help="처리 개수 제한 (테스트용)")
    args = parser.parse_args()

    cfg = load_config(args.config)
    logger = setup_logger(cfg["options"].get("log_file", "logs/result.log"))
    progress = Progress("logs/progress.json")

    # 데이터 로드
    logger.info("📂 데이터 로딩 중...")
    items = load_data(cfg)
    if args.limit:
        items = items[:args.limit]
    logger.info(f"  → {len(items)}개 항목 로드 완료")

    if args.dry_run:
        logger.info("\n[DRY RUN] 실제 등록하지 않고 미리보기만 표시합니다\n")
        for i, item in enumerate(items[:10], 1):
            print(f"  {i:>3}. {item['name']}")
            print(f"       주소: {item['address']}")
            if item['memo']:
                print(f"       메모: {item['memo']}")
        if len(items) > 10:
            print(f"  ... 외 {len(items)-10}개")
        return

    # 실행
    use_kakao = cfg["kakao"]["enabled"] and not args.naver_only
    use_naver = cfg["naver"]["enabled"] and not args.kakao_only
    run_registration(cfg, logger, progress, items, use_kakao, use_naver)

if __name__ == "__main__":
    main()

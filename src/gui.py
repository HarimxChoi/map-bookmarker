"""map-bookmarker GUI"""
import os, sys, threading, logging, json

# EXE 환경에서 Playwright 브라우저 경로 설정 (가장 먼저)
if getattr(sys, 'frozen', False):
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = os.path.join(
        os.path.expanduser("~"), ".playwright-browsers"
    )

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd

# PyInstaller EXE와 일반 실행 모두 지원
if getattr(sys, 'frozen', False):
    _base = sys._MEIPASS
    sys.path.insert(0, os.path.join(_base, "src"))
else:
    sys.path.insert(0, os.path.dirname(__file__))
from main import load_config, load_data, run_registration, Progress


# 로그 핸들러: ScrolledText 위젯으로 리디렉션
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record) + "\n"
        self.text_widget.after(0, self._append, msg)

    def _append(self, msg):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", msg)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")


# 색상/스타일
COLORS = {
    "bg": "#1e1e2e",
    "surface": "#2a2a3d",
    "card": "#313147",
    "accent": "#7c6ff7",
    "accent_hover": "#9b8afb",
    "text": "#e0e0e0",
    "text_dim": "#8888a0",
    "success": "#4ade80",
    "error": "#f87171",
    "warning": "#fbbf24",
    "border": "#3d3d56",
}


class App:
    # 앱 기본 경로 (EXE든 소스든 프로젝트 루트 기준)
    if getattr(sys, 'frozen', False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    DEFAULT_CONFIG = os.path.join(BASE_DIR, "config", "config.yaml")
    LOG_DIR = os.path.join(BASE_DIR, "logs")
    LOG_FILE = os.path.join(LOG_DIR, "result.log")
    PROGRESS_FILE = os.path.join(LOG_DIR, "progress.json")

    def __init__(self):
        self.root = tk.Tk()
        self.root.title("map-bookmarker")
        self.root.geometry("920x1050")
        self.root.minsize(850, 950)
        self.root.configure(bg=COLORS["bg"])

        self.stop_event = threading.Event()
        self.worker_thread = None
        self.detected_columns = []
        self.df_preview = None  # 미리보기용 DataFrame

        self._setup_style()
        self._build_ui()
        self._auto_load_config()

    def _setup_style(self):
        style = ttk.Style()
        style.theme_use("clam")

        style.configure(".", background=COLORS["bg"], foreground=COLORS["text"],
                        fieldbackground=COLORS["card"], borderwidth=0)
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"])
        style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["text"],
                        font=("Segoe UI", 10))
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"),
                        foreground=COLORS["text"])
        style.configure("Dim.TLabel", foreground=COLORS["text_dim"], font=("Segoe UI", 9))
        style.configure("TLabelframe", background=COLORS["surface"],
                        foreground=COLORS["text"])
        style.configure("TLabelframe.Label", background=COLORS["surface"],
                        foreground=COLORS["accent"], font=("Segoe UI", 10, "bold"))
        style.configure("TEntry", fieldbackground=COLORS["card"],
                        foreground=COLORS["text"], insertcolor=COLORS["text"])
        style.configure("TCombobox", fieldbackground=COLORS["card"],
                        foreground=COLORS["text"], selectbackground=COLORS["accent"],
                        arrowcolor=COLORS["text"])
        style.map("TCombobox",
                  fieldbackground=[("readonly", COLORS["card"])],
                  foreground=[("readonly", COLORS["text"])],
                  selectbackground=[("readonly", COLORS["accent"])],
                  selectforeground=[("readonly", "white")])
        # Combobox 드롭다운 리스트 색상 (tk 옵션)
        self.root.option_add("*TCombobox*Listbox.background", COLORS["card"])
        self.root.option_add("*TCombobox*Listbox.foreground", COLORS["text"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", COLORS["accent"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", "white")
        style.configure("TCheckbutton", background=COLORS["bg"],
                        foreground=COLORS["text"],
                        indicatorcolor=COLORS["card"],
                        indicatorrelief="flat")
        style.map("TCheckbutton",
                  indicatorcolor=[("selected", COLORS["accent"])])
        style.configure("TSpinbox", fieldbackground=COLORS["card"],
                        foreground=COLORS["text"])
        style.configure("Treeview", background=COLORS["card"],
                        foreground=COLORS["text"], fieldbackground=COLORS["card"],
                        rowheight=24, font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background=COLORS["surface"],
                        foreground=COLORS["accent"], font=("Segoe UI", 9, "bold"))
        style.map("Treeview", background=[("selected", COLORS["accent"])])

        # 버튼
        style.configure("Accent.TButton", background=COLORS["accent"],
                        foreground="white", font=("Segoe UI", 10, "bold"), padding=(16, 8))
        style.map("Accent.TButton",
                  background=[("active", COLORS["accent_hover"]), ("disabled", COLORS["border"])])
        style.configure("TButton", background=COLORS["surface"],
                        foreground=COLORS["text"], padding=(10, 6))
        style.map("TButton",
                  background=[("active", COLORS["card"])])

        # 노트북(탭)
        style.configure("TNotebook", background=COLORS["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=COLORS["surface"],
                        foreground=COLORS["text_dim"], padding=(16, 8),
                        font=("Segoe UI", 10))
        style.map("TNotebook.Tab",
                  background=[("selected", COLORS["accent"])],
                  foreground=[("selected", "white")])

        # 프로그레스바
        style.configure("Accent.Horizontal.TProgressbar",
                        troughcolor=COLORS["surface"], background=COLORS["accent"])

    # --- UI 빌드 ---
    def _build_ui(self):
        # 메인 컨테이너
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill="both", expand=True)

        # 상단: 즐겨찾기 폴더명
        self._build_folder_section(main)

        # 탭
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True, pady=(8, 0))

        self._build_tab_data()
        self._build_tab_platform()
        self._build_tab_options()

        # 하단: 버튼 + 프로그레스 + 로그
        self._build_bottom(main)

    # 즐겨찾기 폴더명 (상단)
    def _build_folder_section(self, parent):
        f = ttk.LabelFrame(parent, text="  즐겨찾기 폴더명  ", padding=10)
        f.pack(fill="x", pady=(0, 4))
        f.configure(style="TLabelframe")

        inner = ttk.Frame(f, style="TFrame")
        inner.pack(fill="x")
        inner.configure(style="TLabelframe")

        # 카카오
        r1 = ttk.Frame(inner)
        r1.pack(fill="x", pady=2)
        r1.configure(style="TLabelframe")
        ttk.Label(r1, text="카카오맵:", width=12, anchor="e",
                  background=COLORS["surface"]).pack(side="left")
        self.var_kakao_folder = tk.StringVar(value="AUTOMATED FAVORITES")
        ttk.Entry(r1, textvariable=self.var_kakao_folder, width=35).pack(side="left", padx=(4, 12))

        # 네이버
        ttk.Label(r1, text="네이버지도:", width=12, anchor="e",
                  background=COLORS["surface"]).pack(side="left")
        self.var_naver_folder = tk.StringVar(value="AUTOMATED FAVORITES")
        ttk.Entry(r1, textvariable=self.var_naver_folder, width=35).pack(side="left", padx=4)

    # 탭1: 데이터 설정
    def _build_tab_data(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="  데이터  ")

        # 파일 선택
        f_file = ttk.LabelFrame(tab, text="  입력 파일  ", padding=8)
        f_file.pack(fill="x", pady=(0, 8))

        row = ttk.Frame(f_file)
        row.pack(fill="x")
        row.configure(style="TLabelframe")
        self.var_file = tk.StringVar()
        e = ttk.Entry(row, textvariable=self.var_file, width=60)
        e.pack(side="left", fill="x", expand=True)
        tk.Button(row, text="파일 선택", command=self._browse_file,
                  bg=COLORS["accent"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 9, "bold"), padx=12, pady=2,
                  cursor="hand2", activebackground=COLORS["accent_hover"]).pack(side="left", padx=(6, 0))

        row2 = ttk.Frame(f_file)
        row2.pack(fill="x", pady=(6, 0))
        row2.configure(style="TLabelframe")
        ttk.Label(row2, text="Sheet:", background=COLORS["surface"]).pack(side="left")
        self.var_sheet = tk.StringVar(value="Sheet1")
        ttk.Entry(row2, textvariable=self.var_sheet, width=15).pack(side="left", padx=(4, 12))
        ttk.Label(row2, text="Header Row:", background=COLORS["surface"]).pack(side="left")
        self.var_header = tk.IntVar(value=1)
        ttk.Spinbox(row2, textvariable=self.var_header, from_=1, to=100, width=5).pack(side="left", padx=4)
        tk.Button(row2, text="새로고침", command=self._reload_file,
                  bg=COLORS["card"], fg=COLORS["text"], relief="flat",
                  font=("Segoe UI", 9), padx=10, pady=2,
                  cursor="hand2", activebackground=COLORS["surface"]).pack(side="right")

        # 컬럼 선택
        f_col = ttk.LabelFrame(tab, text="  컬럼 선택  ", padding=8)
        f_col.pack(fill="x", pady=(0, 8))

        # 주소 컬럼
        r1 = ttk.Frame(f_col)
        r1.pack(fill="x", pady=3)
        r1.configure(style="TLabelframe")
        ttk.Label(r1, text="주소 컬럼 (필수):", width=18, anchor="e",
                  background=COLORS["surface"]).pack(side="left")
        self.var_col_addr = tk.StringVar()
        self.combo_addr = ttk.Combobox(r1, textvariable=self.var_col_addr,
                                        width=30, state="readonly")
        self.combo_addr.pack(side="left", padx=4)
        self.combo_addr.bind("<<ComboboxSelected>>", lambda e: self._update_preview())
        ttk.Label(r1, text="* 필수", foreground=COLORS["error"],
                  background=COLORS["surface"], font=("Segoe UI", 9)).pack(side="left")

        # 즐겨찾기명 컬럼
        r2 = ttk.Frame(f_col)
        r2.pack(fill="x", pady=3)
        r2.configure(style="TLabelframe")
        ttk.Label(r2, text="즐겨찾기명 컬럼:", width=18, anchor="e",
                  background=COLORS["surface"]).pack(side="left")
        self.var_col_name = tk.StringVar()
        self.combo_name = ttk.Combobox(r2, textvariable=self.var_col_name,
                                        width=30, state="readonly")
        self.combo_name.pack(side="left", padx=4)
        self.combo_name.bind("<<ComboboxSelected>>", lambda e: self._update_preview())
        ttk.Label(r2, text="선택 안 하면 주소가 기본값",
                  foreground=COLORS["text_dim"], background=COLORS["surface"],
                  font=("Segoe UI", 9)).pack(side="left", padx=4)

        # 동/호수 자동 추가 체크박스
        r3 = ttk.Frame(f_col)
        r3.pack(fill="x", pady=3)
        r3.configure(style="TLabelframe")
        ttk.Label(r3, text="", width=18, background=COLORS["surface"]).pack(side="left")
        self.var_append_unit = tk.BooleanVar(value=True)
        ttk.Checkbutton(r3, text="동/호수 자동 추가  (예: 홍길동 → 홍길동 (106동 1102호))",
                        variable=self.var_append_unit).pack(side="left", padx=4)

        # 즐겨찾기명 초기화 버튼
        tk.Button(r2, text="초기화", command=self._clear_name_col,
                  bg=COLORS["card"], fg=COLORS["text"], relief="flat",
                  font=("Segoe UI", 9), padx=10, pady=2,
                  cursor="hand2", activebackground=COLORS["surface"]).pack(side="right")

        # 필터 조건
        f_filter = ttk.LabelFrame(tab, text="  필터 조건 (AND)  ", padding=6)
        f_filter.pack(fill="x", pady=(0, 8))

        # 필터 버튼 (상단)
        filter_btn_row = ttk.Frame(f_filter)
        filter_btn_row.pack(fill="x", pady=(0, 4))
        filter_btn_row.configure(style="TLabelframe")

        btn_add = tk.Button(filter_btn_row, text="+ 추가", command=self._add_filter,
                            bg=COLORS["accent"], fg="#ffffff", relief="flat",
                            font=("Segoe UI", 9, "bold"), padx=12, pady=2,
                            cursor="hand2", activebackground="#5a6abf")
        btn_add.pack(side="left")

        btn_del = tk.Button(filter_btn_row, text="- 삭제", command=self._remove_filter,
                            bg=COLORS["error"], fg="#ffffff", relief="flat",
                            font=("Segoe UI", 9, "bold"), padx=12, pady=2,
                            cursor="hand2", activebackground="#cc3333")
        btn_del.pack(side="left", padx=(6, 0))

        ttk.Label(filter_btn_row, text="조건을 모두 만족하는 행만 등록됩니다",
                  foreground=COLORS["text_dim"], background=COLORS["surface"],
                  font=("Segoe UI", 8)).pack(side="right")

        # 필터 테이블
        filter_cols = ("column", "type", "value")
        self.filter_tree = ttk.Treeview(f_filter, columns=filter_cols,
                                         show="headings", height=2)
        self.filter_tree.heading("column", text="컬럼")
        self.filter_tree.heading("type", text="조건")
        self.filter_tree.heading("value", text="값")
        self.filter_tree.column("column", width=150)
        self.filter_tree.column("type", width=130)
        self.filter_tree.column("value", width=250)
        self.filter_tree.pack(fill="x")

        # 데이터 미리보기
        f_preview = ttk.LabelFrame(tab, text="  데이터 미리보기  ", padding=4)
        f_preview.pack(fill="both", expand=True)

        # 미리보기 정보
        self.var_preview_info = tk.StringVar(value="파일을 선택하면 데이터가 표시됩니다")
        ttk.Label(f_preview, textvariable=self.var_preview_info,
                  foreground=COLORS["text_dim"],
                  background=COLORS["surface"]).pack(anchor="w", padx=4, pady=(2, 4))

        # Treeview
        tree_frame = ttk.Frame(f_preview)
        tree_frame.pack(fill="both", expand=True)
        tree_frame.configure(style="TLabelframe")

        self.preview_tree = ttk.Treeview(tree_frame, show="headings", height=5)
        self.preview_tree.pack(side="left", fill="both", expand=True)

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical",
                                     command=self.preview_tree.yview)
        scrollbar_y.pack(side="right", fill="y")
        self.preview_tree.configure(yscrollcommand=scrollbar_y.set)

        scrollbar_x = ttk.Scrollbar(f_preview, orient="horizontal",
                                     command=self.preview_tree.xview)
        scrollbar_x.pack(fill="x")
        self.preview_tree.configure(xscrollcommand=scrollbar_x.set)

    # 탭2: 플랫폼 설정
    def _build_tab_platform(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="  플랫폼  ")

        # 카카오맵
        f = ttk.LabelFrame(tab, text="  카카오맵  ", padding=10)
        f.pack(fill="x", pady=(0, 12))

        self.var_kakao_on = tk.BooleanVar(value=True)
        ttk.Checkbutton(f, text="카카오맵 사용", variable=self.var_kakao_on).pack(anchor="w")

        inner = ttk.Frame(f)
        inner.pack(fill="x", padx=(20, 0), pady=(6, 0))
        inner.configure(style="TLabelframe")

        self.var_kakao_id = tk.StringVar()
        self.var_kakao_pw = tk.StringVar()

        for lbl, var, show in [("ID:", self.var_kakao_id, ""),
                                ("Password:", self.var_kakao_pw, "*")]:
            r = ttk.Frame(inner)
            r.pack(fill="x", pady=2)
            r.configure(style="TLabelframe")
            ttk.Label(r, text=lbl, width=10, anchor="e",
                      background=COLORS["surface"]).pack(side="left")
            e = ttk.Entry(r, textvariable=var, width=35, show=show)
            e.pack(side="left", padx=4)
            if show == "*":
                self._kakao_pw_entry = e
                ttk.Button(r, text="Show", width=5,
                           command=lambda: self._toggle_pw(self._kakao_pw_entry)).pack(side="left")

        # 네이버지도
        f2 = ttk.LabelFrame(tab, text="  네이버지도  ", padding=10)
        f2.pack(fill="x")

        self.var_naver_on = tk.BooleanVar(value=True)
        ttk.Checkbutton(f2, text="네이버지도 사용", variable=self.var_naver_on).pack(anchor="w")

        inner2 = ttk.Frame(f2)
        inner2.pack(fill="x", padx=(20, 0), pady=(6, 0))
        inner2.configure(style="TLabelframe")

        self.var_naver_id = tk.StringVar()
        self.var_naver_pw = tk.StringVar()

        for lbl, var, show in [("ID:", self.var_naver_id, ""),
                                ("Password:", self.var_naver_pw, "*")]:
            r = ttk.Frame(inner2)
            r.pack(fill="x", pady=2)
            r.configure(style="TLabelframe")
            ttk.Label(r, text=lbl, width=10, anchor="e",
                      background=COLORS["surface"]).pack(side="left")
            e = ttk.Entry(r, textvariable=var, width=35, show=show)
            e.pack(side="left", padx=4)
            if show == "*":
                self._naver_pw_entry = e
                ttk.Button(r, text="Show", width=5,
                           command=lambda: self._toggle_pw(self._naver_pw_entry)).pack(side="left")

    # 탭3: 실행 옵션
    def _build_tab_options(self):
        tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(tab, text="  옵션  ")

        f = ttk.LabelFrame(tab, text="  실행 옵션  ", padding=10)
        f.pack(fill="x", pady=(0, 12))

        self.var_headless = tk.BooleanVar(value=False)
        self.var_delay = tk.IntVar(value=800)
        self.var_retry = tk.IntVar(value=3)
        self.var_resume = tk.BooleanVar(value=True)
        self.var_limit = tk.IntVar(value=0)

        opts = [
            ("Headless (브라우저 숨김):", self.var_headless, "check"),
            ("등록 간 대기 (ms):", self.var_delay, "spin"),
            ("최대 재시도:", self.var_retry, "spin"),
            ("이전 진행 이어하기:", self.var_resume, "check"),
            ("처리 제한 (0=전체):", self.var_limit, "spin"),
        ]

        for lbl, var, wtype in opts:
            r = ttk.Frame(f)
            r.pack(fill="x", pady=3)
            r.configure(style="TLabelframe")
            ttk.Label(r, text=lbl, width=22, anchor="e",
                      background=COLORS["surface"]).pack(side="left")
            if wtype == "check":
                ttk.Checkbutton(r, variable=var).pack(side="left", padx=4)
            else:
                ttk.Spinbox(r, textvariable=var, from_=0, to=99999, width=8).pack(side="left", padx=4)

        # 버튼들
        f2 = ttk.Frame(tab)
        f2.pack(fill="x", pady=8)
        ttk.Button(f2, text="설정 저장 (config.yaml)", command=self._save_config).pack(side="left")
        ttk.Button(f2, text="설정 불러오기", command=self._load_config_dialog).pack(side="left", padx=8)
        ttk.Button(f2, text="progress 초기화", command=self._reset_progress).pack(side="right")

    # 하단 영역: 버튼 + 프로그레스 + 로그
    def _build_bottom(self, parent):
        bottom = ttk.Frame(parent)
        bottom.pack(fill="both", expand=True, pady=(8, 0))

        # 버튼
        btn_row = ttk.Frame(bottom)
        btn_row.pack(fill="x")

        self.btn_start = ttk.Button(btn_row, text="  Start  ",
                                     command=self._start, style="Accent.TButton")
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(btn_row, text="  Stop  ",
                                    command=self._stop, state="disabled")
        self.btn_stop.pack(side="left", padx=4)
        ttk.Button(btn_row, text="  Dry Run  ", command=self._dry_run).pack(side="left", padx=4)

        # 통계
        self.var_stats = tk.StringVar(value="")
        ttk.Label(btn_row, textvariable=self.var_stats,
                  font=("Consolas", 9)).pack(side="right")

        # 프로그레스바
        self.progressbar = ttk.Progressbar(bottom, mode="determinate",
                                            style="Accent.Horizontal.TProgressbar")
        self.progressbar.pack(fill="x", pady=(6, 4))

        # 로그
        log_frame = ttk.LabelFrame(bottom, text="  Log  ", padding=4)
        log_frame.pack(fill="both", expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled",
            font=("Consolas", 9), wrap="word",
            bg=COLORS["card"], fg=COLORS["text"],
            insertbackground=COLORS["text"],
            selectbackground=COLORS["accent"]
        )
        self.log_text.pack(fill="both", expand=True)

        clear_row = ttk.Frame(log_frame)
        clear_row.pack(fill="x")
        clear_row.configure(style="TLabelframe")
        ttk.Button(clear_row, text="Clear", command=self._clear_log).pack(side="right")

    # --- 파일 선택 & 컬럼 감지 ---
    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="입력 파일 선택",
            filetypes=[
                ("지원 파일", "*.xlsx *.xls *.csv *.xml *.tsv"),
                ("Excel", "*.xlsx *.xls"),
                ("CSV", "*.csv *.tsv"),
                ("XML", "*.xml"),
                ("All", "*.*"),
            ]
        )
        if path:
            self.var_file.set(path)
            self._load_file(path)

    def _reload_file(self):
        path = self.var_file.get()
        if path and os.path.exists(path):
            # 컬럼 선택 리셋
            self.var_col_addr.set("")
            self.var_col_name.set("(선택 안 함)")
            self._load_file(path)
        else:
            self._log_msg("파일 경로가 비어있거나 존재하지 않습니다.")

    def _load_file(self, path):
        """파일 로드 → 컬럼 감지 → 자동 추천 → 미리보기"""
        try:
            ext = os.path.splitext(path)[1].lower()
            header_row = self.var_header.get() - 1

            if ext == ".csv" or ext == ".tsv":
                sep = "\t" if ext == ".tsv" else ","
                self.df_preview = pd.read_csv(path, encoding="utf-8-sig", sep=sep,
                                               header=header_row)
            elif ext == ".xml":
                self.df_preview = pd.read_xml(path)
            else:
                sheet = self.var_sheet.get() or 0
                self.df_preview = pd.read_excel(path, sheet_name=sheet,
                                                 header=header_row)

            cols = [str(c) for c in self.df_preview.columns.tolist()]
            self.detected_columns = cols

            # 드롭다운 업데이트
            empty_option = ["(선택 안 함)"]
            self.combo_addr["values"] = cols
            self.combo_name["values"] = empty_option + cols

            # 자동 추천: 주소 컬럼
            addr_keywords = ["주소", "address", "addr", "도로명", "지번", "위치", "location"]
            for col in cols:
                if any(kw in col.lower() for kw in addr_keywords):
                    self.var_col_addr.set(col)
                    break

            # 자동 추천: 이름 컬럼
            name_keywords = ["이름", "name", "상호", "업체", "매장", "고객", "장소"]
            for col in cols:
                if any(kw in col.lower() for kw in name_keywords):
                    self.var_col_name.set(col)
                    break

            # 미리보기 업데이트
            self._update_preview()
            self._log_msg(f"파일 로드 완료: {os.path.basename(path)} ({len(self.df_preview)}행, {len(cols)}열)")

        except Exception as e:
            self._log_msg(f"파일 로드 실패: {e}")
            self.df_preview = None

    def _update_preview(self):
        """미리보기 테이블 갱신"""
        tree = self.preview_tree

        # 기존 데이터 삭제
        tree.delete(*tree.get_children())
        for col in tree["columns"]:
            tree.heading(col, text="")
        tree["columns"] = ()

        if self.df_preview is None or self.df_preview.empty:
            self.var_preview_info.set("데이터 없음")
            return

        df = self.df_preview
        addr_col = self.var_col_addr.get()
        name_col = self.var_col_name.get()

        # 표시할 컬럼 결정 (선택된 컬럼 우선 + 나머지)
        show_cols = []
        if addr_col and addr_col in df.columns:
            show_cols.append(addr_col)
        if name_col and name_col != "(선택 안 함)" and name_col in df.columns:
            if name_col not in show_cols:
                show_cols.append(name_col)
        for c in df.columns:
            if c not in show_cols:
                show_cols.append(c)

        # 최대 10개 컬럼만 표시
        display_cols = show_cols[:10]
        tree["columns"] = display_cols

        for col in display_cols:
            w = 120
            if col == addr_col:
                w = 250
                tree.heading(col, text=f"📍 {col}")
            elif col == name_col and name_col != "(선택 안 함)":
                w = 150
                tree.heading(col, text=f"⭐ {col}")
            else:
                tree.heading(col, text=col)
            tree.column(col, width=w, minwidth=60)

        # 데이터 삽입 (최대 50행)
        for _, row in df.head(50).iterrows():
            vals = []
            for col in display_cols:
                v = row[col]
                vals.append(str(v) if pd.notna(v) else "")
            tree.insert("", "end", values=vals)

        total = len(df)
        shown = min(total, 50)
        info = f"총 {total}행"
        if shown < total:
            info += f" (상위 {shown}행 미리보기)"
        if addr_col:
            info += f"  |  주소: {addr_col}"
        if name_col and name_col != "(선택 안 함)":
            info += f"  |  즐겨찾기명: {name_col}"
        else:
            info += "  |  즐겨찾기명: 주소 기본값"
        self.var_preview_info.set(info)

    def _clear_name_col(self):
        self.var_col_name.set("(선택 안 함)")
        self._update_preview()

    # --- 필터 관리 ---
    def _add_filter(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("필터 추가")
        dlg.geometry("400x260")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg=COLORS["bg"])

        pad = {"padx": 16, "pady": (0, 0)}

        tk.Label(dlg, text="필터 추가", font=("Segoe UI", 12, "bold"),
                 bg=COLORS["bg"], fg=COLORS["accent"]).pack(pady=(12, 8))

        tk.Label(dlg, text="컬럼", bg=COLORS["bg"], fg=COLORS["text"],
                 font=("Segoe UI", 9)).pack(anchor="w", **pad)
        col_var = tk.StringVar()
        ttk.Combobox(dlg, textvariable=col_var,
                     values=self.detected_columns, width=32).pack(**pad)

        tk.Label(dlg, text="조건", bg=COLORS["bg"], fg=COLORS["text"],
                 font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(8, 0))
        type_var = tk.StringVar()
        type_cb = ttk.Combobox(dlg, textvariable=type_var,
                     values=["not_contains (제외)", "contains (포함)", "min (최소)", "max (최대)"],
                     width=32, state="readonly")
        type_cb.pack(**pad)

        tk.Label(dlg, text="값 (여러 개는 쉼표로 구분)", bg=COLORS["bg"],
                 fg=COLORS["text"], font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(8, 0))
        val_var = tk.StringVar()
        ttk.Entry(dlg, textvariable=val_var, width=34).pack(**pad)

        def ok():
            if col_var.get() and type_var.get() and val_var.get():
                # "not_contains (제외)" → "not_contains"
                ftype = type_var.get().split(" ")[0]
                self.filter_tree.insert("", "end",
                    values=(col_var.get(), ftype, val_var.get()))
            dlg.destroy()

        btn_frame = tk.Frame(dlg, bg=COLORS["bg"])
        btn_frame.pack(pady=12)
        tk.Button(btn_frame, text="추가", command=ok,
                  bg=COLORS["accent"], fg="#ffffff", relief="flat",
                  font=("Segoe UI", 10, "bold"), padx=20, pady=4,
                  cursor="hand2").pack(side="left", padx=4)
        tk.Button(btn_frame, text="취소", command=dlg.destroy,
                  bg=COLORS["surface"], fg=COLORS["text"], relief="flat",
                  font=("Segoe UI", 10), padx=20, pady=4,
                  cursor="hand2").pack(side="left", padx=4)

    def _remove_filter(self):
        sel = self.filter_tree.selection()
        for s in sel:
            self.filter_tree.delete(s)

    def _get_filters(self):
        """필터 Treeview → config용 리스트"""
        filters = []
        for item_id in self.filter_tree.get_children():
            vals = self.filter_tree.item(item_id, "values")
            col, ftype, fval = vals
            f = {"column": col}
            if ftype in ("not_contains", "contains"):
                f[ftype] = [v.strip() for v in fval.split(",")]
            elif ftype in ("min", "max"):
                try:
                    f[ftype] = float(fval)
                except ValueError:
                    f[ftype] = fval
            filters.append(f)
        return filters

    # --- 유틸 ---
    def _toggle_pw(self, entry):
        current = entry.cget("show")
        entry.configure(show="" if current == "*" else "*")

    def _log_msg(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _clear_log(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    # --- 설정 직렬화 ---
    def _build_config_dict(self):
        name_col = self.var_col_name.get()
        if name_col == "(선택 안 함)" or not name_col:
            name_col = None

        return {
            "input": {
                "file": self.var_file.get(),
                "sheet": self.var_sheet.get(),
                "header_row": self.var_header.get(),
            },
            "columns": {
                "name": name_col,
                "address": self.var_col_addr.get() or None,
                "label": None,
                "memo": None,
            },
            "bookmark_name": f"{{{name_col}}}" if name_col else "{" + (self.var_col_addr.get() or "주소") + "}",
            "bookmark_memo": "",
            "append_unit_info": self.var_append_unit.get(),
            "filters": self._get_filters(),
            "kakao": {
                "enabled": self.var_kakao_on.get(),
                "id": self.var_kakao_id.get(),
                "password": self.var_kakao_pw.get(),
                "folder": self.var_kakao_folder.get(),
            },
            "naver": {
                "enabled": self.var_naver_on.get(),
                "id": self.var_naver_id.get(),
                "password": self.var_naver_pw.get(),
                "folder": self.var_naver_folder.get(),
            },
            "options": {
                "headless": self.var_headless.get(),
                "delay_ms": self.var_delay.get(),
                "max_retry": self.var_retry.get(),
                "resume": self.var_resume.get(),
                "log_file": "logs/result.log",
            },
        }

    def _load_config_to_gui(self, cfg):
        inp = cfg.get("input", {})
        self.var_file.set(inp.get("file", ""))
        self.var_sheet.set(inp.get("sheet", "Sheet1"))
        self.var_header.set(inp.get("header_row", 1))

        if self.var_file.get() and os.path.exists(self.var_file.get()):
            self._load_file(self.var_file.get())

        cols = cfg.get("columns", {})
        if cols.get("address"):
            self.var_col_addr.set(cols["address"])
        if cols.get("name"):
            self.var_col_name.set(cols["name"])

        kakao = cfg.get("kakao", {})
        self.var_kakao_on.set(kakao.get("enabled", True))
        self.var_kakao_id.set(kakao.get("id", ""))
        self.var_kakao_pw.set(kakao.get("password", ""))
        self.var_kakao_folder.set(kakao.get("folder", "AUTOMATED FAVORITES"))

        naver = cfg.get("naver", {})
        self.var_naver_on.set(naver.get("enabled", True))
        self.var_naver_id.set(naver.get("id", ""))
        self.var_naver_pw.set(naver.get("password", ""))
        self.var_naver_folder.set(naver.get("folder", "AUTOMATED FAVORITES"))

        opts = cfg.get("options", {})
        self.var_headless.set(opts.get("headless", False))
        self.var_delay.set(opts.get("delay_ms", 800))
        self.var_retry.set(opts.get("max_retry", 3))
        self.var_resume.set(opts.get("resume", True))
        self.var_limit.set(0)

    def _save_config(self):
        path = filedialog.asksaveasfilename(
            title="설정 저장", initialfile="config.yaml", initialdir="config",
            filetypes=[("YAML", "*.yaml *.yml")]
        )
        if path:
            import yaml
            cfg = self._build_config_dict()
            with open(path, "w", encoding="utf-8") as f:
                yaml.dump(cfg, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            self._log_msg(f"설정 저장 완료: {path}")

    def _load_config_dialog(self):
        path = filedialog.askopenfilename(
            title="설정 불러오기", initialdir="config",
            filetypes=[("YAML", "*.yaml *.yml")]
        )
        if path:
            cfg = load_config(path)
            self._load_config_to_gui(cfg)
            self._log_msg(f"설정 불러오기 완료: {path}")

    def _auto_load_config(self):
        if os.path.exists(self.DEFAULT_CONFIG):
            try:
                cfg = load_config(self.DEFAULT_CONFIG)
                self._load_config_to_gui(cfg)
            except Exception:
                pass

    def _reset_progress(self):
        path = self.PROGRESS_FILE
        os.makedirs("logs", exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"done": []}, f)
        self._log_msg("progress.json 초기화 완료")

    # --- 실행 ---
    def _get_logger(self):
        logger = logging.getLogger("map-reg-gui")
        logger.setLevel(logging.DEBUG)
        logger.handlers.clear()
        th = TextHandler(self.log_text)
        th.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(th)
        os.makedirs(self.LOG_DIR, exist_ok=True)
        fh = logging.FileHandler(self.LOG_FILE, encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%H:%M:%S"))
        logger.addHandler(fh)
        return logger

    def _dry_run(self):
        cfg = self._build_config_dict()
        if not cfg["columns"].get("address"):
            messagebox.showerror("오류", "주소 컬럼을 선택하세요.")
            return
        try:
            items = load_data(cfg)
            limit = self.var_limit.get()
            if limit:
                items = items[:limit]
            self._log_msg(f"\n[DRY RUN] {len(items)}개 항목 미리보기\n")
            for i, item in enumerate(items[:20], 1):
                self._log_msg(f"  {i:>3}. {item['name']}")
                self._log_msg(f"       주소: {item['address']}")
            if len(items) > 20:
                self._log_msg(f"  ... 외 {len(items) - 20}개")
        except Exception as e:
            self._log_msg(f"오류: {e}")

    def _start(self):
        cfg = self._build_config_dict()
        if not cfg["columns"].get("address"):
            messagebox.showerror("오류", "주소 컬럼을 선택하세요.")
            return
        if not cfg["input"].get("file"):
            messagebox.showerror("오류", "입력 파일을 선택하세요.")
            return

        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.stop_event.clear()

        self.worker_thread = threading.Thread(target=self._run_worker, args=(cfg,), daemon=True)
        self.worker_thread.start()

    def _stop(self):
        self.stop_event.set()
        self.btn_stop.configure(state="disabled")
        self._log_msg("중지 요청됨... 현재 작업 완료 후 중단합니다.")

    def _run_worker(self, cfg):
        logger = self._get_logger()
        progress = Progress(self.PROGRESS_FILE)

        try:
            logger.info("데이터 로딩 중...")
            items = load_data(cfg)
            limit = self.var_limit.get()
            if limit:
                items = items[:limit]
            logger.info(f"  -> {len(items)}개 항목 로드 완료")

            total = len(items)
            self.root.after(0, lambda: self.progressbar.configure(maximum=total, value=0))
            self._processed = 0

            use_kakao = cfg["kakao"]["enabled"]
            use_naver = cfg["naver"]["enabled"]

            def on_progress(platform, status, item, stats):
                self._processed += 1
                self.root.after(0, self._update_ui, stats, self._processed, total)

            run_registration(cfg, logger, progress, items, use_kakao, use_naver,
                             on_progress=on_progress, stop_event=self.stop_event)
        except Exception as e:
            logger.error(f"실행 오류: {e}")
        finally:
            self.root.after(0, self._on_worker_done)

    def _update_ui(self, stats, processed, total):
        self.progressbar.configure(value=processed)
        k = stats["kakao"]
        n = stats["naver"]
        self.var_stats.set(
            f"Kakao: O{k['ok']} X{k['fail']} Skip{k['skip']}  |  "
            f"Naver: O{n['ok']} X{n['fail']} Skip{n['skip']}  "
            f"({processed}/{total})"
        )

    def _on_worker_done(self):
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")
        self._log_msg("작업 완료.")

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()

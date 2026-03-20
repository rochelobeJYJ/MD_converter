"""
PDF → Markdown 변환기
=====================
opendataloader-pdf 라이브러리를 사용하여 PDF 파일을 Markdown으로 변환하는 GUI 앱.
tkinter + tkinterdnd2 기반 경량 데스크탑 애플리케이션.
"""

import os
import sys
import threading
import time
import subprocess
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from collections import defaultdict

# ── 윈도우 서브프로세스 패치 (깜빡이는 콘솔창 숨김) ───────────────
if os.name == 'nt':
    _real_Popen = subprocess.Popen
    def _Popen(*args, **kwargs):
        if 'creationflags' not in kwargs:
            kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
        return _real_Popen(*args, **kwargs)
    subprocess.Popen = _Popen

# ── Java 환경 설정 ──────────────────────────────────────────────────
JAVA_HOME = r"C:\Program Files\Eclipse Adoptium\jdk-17.0.18.8-hotspot"
JAVA_BIN = os.path.join(JAVA_HOME, "bin")
if JAVA_BIN not in os.environ.get("PATH", ""):
    os.environ["PATH"] = JAVA_BIN + os.pathsep + os.environ.get("PATH", "")
os.environ["JAVA_HOME"] = JAVA_HOME

# ── tkinterdnd2 임포트 ──────────────────────────────────────────────
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

import opendataloader_pdf


# ═══════════════════════════════════════════════════════════════════
#  색상 팔레트 & 스타일 상수 (그레이톤)
# ═══════════════════════════════════════════════════════════════════
class Theme:
    # 배경
    BG_DARK = "#f3f4f6"      # 전체 배경 (연한 회색)
    BG_CARD = "#ffffff"      # 카드 패널 배경 (흰색)
    BG_INPUT = "#e5e7eb"     # 입력/리스트 박스 배경
    BG_HOVER = "#d1d5db"     # 버튼 호버 배경

    # 액센트
    ACCENT = "#4b5563"       # 주요 버튼 색상 (진한 회색)
    ACCENT_HOVER = "#374151" # 버튼 호버시 (더 진한 회색)
    ACCENT_GLOW = "#1f2937"
    SUCCESS = "#059669"      # 성공 (구분용 녹색 유지)
    ERROR = "#dc2626"        # 에러 (구분용 적색 유지)
    WARNING = "#d97706"      # 경고 (구분용 주황 유지)

    # 텍스트
    TEXT_PRIMARY = "#111827" # 기본 텍스트
    TEXT_SECONDARY = "#374151"
    TEXT_MUTED = "#6b7280"   # 흐린 텍스트
    TEXT_ON_ACCENT = "#ffffff" # 버튼 위 텍스트

    # 보더
    BORDER = "#d1d5db"
    BORDER_ACTIVE = "#9ca3af"

    # 폰트
    FONT_FAMILY = "맑은 고딕"
    FONT_TITLE = ("맑은 고딕", 16, "bold")
    FONT_SUBTITLE = ("맑은 고딕", 11)
    FONT_BODY = ("맑은 고딕", 10)
    FONT_SMALL = ("맑은 고딕", 9)
    FONT_BUTTON = ("맑은 고딕", 10, "bold")
    FONT_LOG = ("Consolas", 9)


# ═══════════════════════════════════════════════════════════════════
#  메인 애플리케이션 클래스
# ═══════════════════════════════════════════════════════════════════
class PDFtoMDApp:
    def __init__(self):
        # ── 루트 윈도우 설정 ────────────────────────────────
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()

        self.root.title("PDF → Markdown 변환기")
        self.root.geometry("680x620")
        self.root.minsize(650, 580)
        self.root.configure(bg=Theme.BG_DARK)

        # 아이콘 설정
        try:
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icon.ico")
            if os.path.exists(icon_path):
                self.root.iconbitmap(default=icon_path)
        except Exception as e:
            print(f"아이콘 로드 실패: {e}")

        # ── 상태 변수 ──────────────────────────────────────
        self.pdf_files: list[str] = []
        self.custom_output_dir: str | None = None
        self.is_converting = False
        self.conversion_just_finished = False

        # ── ttk 스타일 ─────────────────────────────────────
        self._setup_styles()

        # ── UI 구성 ────────────────────────────────────────
        self._build_ui()

        # 윈도우 중앙 배치
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f"+{x}+{y}")

    # ─────────────────────────────────────────────────────────
    #  ttk 스타일 설정
    # ─────────────────────────────────────────────────────────
    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use("clam")

        # 프로그레스바
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=Theme.BG_INPUT,
            background=Theme.ACCENT,
            bordercolor=Theme.BORDER,
            lightcolor=Theme.ACCENT,
            darkcolor=Theme.ACCENT_GLOW,
            thickness=22,
        )

    # ─────────────────────────────────────────────────────────
    #  UI 빌드
    # ─────────────────────────────────────────────────────────
    def _build_ui(self):
        # ── 메인 컨테이너 ──
        main = tk.Frame(self.root, bg=Theme.BG_DARK)
        main.pack(fill="both", expand=True, padx=16, pady=12)

        # ── 타이틀 ──
        title_frame = tk.Frame(main, bg=Theme.BG_DARK)
        title_frame.pack(fill="x", pady=(0, 8))

        tk.Label(
            title_frame,
            text="📄  PDF → Markdown 변환기",
            font=Theme.FONT_TITLE,
            fg=Theme.TEXT_PRIMARY,
            bg=Theme.BG_DARK,
        ).pack(side="left")

        # ═══════════════════════════════════════════════════
        #  섹션 1: 파일 입력
        # ═══════════════════════════════════════════════════
        file_card = self._make_card(main)
        file_card.pack(fill="both", expand=True, pady=(0, 6))

        # 카드 헤더
        header = tk.Frame(file_card, bg=Theme.BG_CARD)
        header.pack(fill="x", padx=12, pady=(10, 4))

        tk.Label(
            header,
            text="📁  PDF 파일 목록",
            font=("맑은 고딕", 11, "bold"),
            fg=Theme.TEXT_PRIMARY,
            bg=Theme.BG_CARD,
        ).pack(side="left")

        self.file_count_label = tk.Label(
            header,
            text="0개 파일",
            font=Theme.FONT_SMALL,
            fg=Theme.TEXT_MUTED,
            bg=Theme.BG_CARD,
        )
        self.file_count_label.pack(side="right")

        # ── 파일 리스트 박스 (드래그 앤 드롭 영역) ──
        list_frame = tk.Frame(file_card, bg=Theme.BG_INPUT, bd=0, highlightthickness=1,
                              highlightbackground=Theme.BORDER, highlightcolor=Theme.BORDER_ACTIVE)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(2, 6))

        self.file_listbox = tk.Listbox(
            list_frame,
            height=6,
            bg=Theme.BG_INPUT,
            fg=Theme.TEXT_PRIMARY,
            font=Theme.FONT_BODY,
            selectbackground=Theme.ACCENT,
            selectforeground=Theme.TEXT_ON_ACCENT,
            activestyle="none",
            bd=0,
            highlightthickness=0,
            relief="flat",
        )
        self.file_listbox.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        scrollbar = tk.Scrollbar(list_frame, command=self.file_listbox.yview,
                                  bg=Theme.BG_INPUT, troughcolor=Theme.BG_INPUT,
                                  activebackground=Theme.ACCENT)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=scrollbar.set)

        # 드래그 앤 드롭 안내 (리스트가 비어있을 때)
        self.drop_hint_id = None
        self._show_drop_hint()

        # DnD 바인딩
        if DND_AVAILABLE:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind("<<Drop>>", self._on_drop)

        # ── 버튼 행 ──
        btn_row = tk.Frame(file_card, bg=Theme.BG_CARD)
        btn_row.pack(fill="x", padx=12, pady=(0, 10))

        self._make_button(btn_row, "📂 파일 찾기", self._browse_files).pack(side="left", padx=(0, 6))
        self._make_button(btn_row, "🗑️ 선택 삭제", self._remove_selected, style="secondary").pack(side="left", padx=(0, 6))
        self._make_button(btn_row, "🧹 전체 삭제", self._clear_files, style="secondary").pack(side="left")

        # ═══════════════════════════════════════════════════
        #  섹션 2: 저장 경로 설정
        # ═══════════════════════════════════════════════════
        save_card = self._make_card(main)
        save_card.pack(fill="x", pady=(0, 6))

        save_inner = tk.Frame(save_card, bg=Theme.BG_CARD)
        save_inner.pack(fill="x", padx=12, pady=10)

        tk.Label(
            save_inner,
            text="💾  저장 경로",
            font=("맑은 고딕", 11, "bold"),
            fg=Theme.TEXT_PRIMARY,
            bg=Theme.BG_CARD,
        ).pack(anchor="w")

        path_row = tk.Frame(save_inner, bg=Theme.BG_CARD)
        path_row.pack(fill="x", pady=(8, 0))

        self.save_path_var = tk.StringVar(value="원본 PDF와 같은 폴더에 저장 (기본)")
        self.save_path_entry = tk.Entry(
            path_row,
            textvariable=self.save_path_var,
            font=Theme.FONT_BODY,
            bg=Theme.BG_INPUT,
            fg=Theme.TEXT_SECONDARY,
            insertbackground=Theme.TEXT_PRIMARY,
            bd=0,
            highlightthickness=1,
            highlightbackground=Theme.BORDER,
            highlightcolor=Theme.BORDER_ACTIVE,
            state="readonly",
            readonlybackground=Theme.BG_INPUT,
        )
        self.save_path_entry.pack(side="left", fill="x", expand=True, ipady=5, padx=(0, 6))

        self._make_button(path_row, "📁 폴더 변경", self._browse_output_dir).pack(side="left", padx=(0, 4))
        self._make_button(path_row, "↩️ 기본값", self._reset_output_dir, style="secondary").pack(side="left")

        # 폴더 자동 열기 옵션
        self.auto_open_var = tk.BooleanVar(value=True)
        opt_check = tk.Checkbutton(
            save_inner,
            text="변환 완료 후 저장 폴더 열기",
            variable=self.auto_open_var,
            bg=Theme.BG_CARD,
            fg=Theme.TEXT_SECONDARY,
            activebackground=Theme.BG_CARD,
            activeforeground=Theme.TEXT_PRIMARY,
            selectcolor=Theme.BG_INPUT,
            font=Theme.FONT_SMALL,
            cursor="hand2",
            bd=0,
            relief="flat",
        )
        opt_check.pack(anchor="w", pady=(8, 0))

        # ═══════════════════════════════════════════════════
        #  섹션 3: 로그 & 진행 상황
        # ═══════════════════════════════════════════════════
        log_card = self._make_card(main)
        log_card.pack(fill="x", pady=(0, 6))

        log_inner = tk.Frame(log_card, bg=Theme.BG_CARD)
        log_inner.pack(fill="x", padx=12, pady=10)

        tk.Label(
            log_inner,
            text="📋  변환 로그",
            font=("맑은 고딕", 11, "bold"),
            fg=Theme.TEXT_PRIMARY,
            bg=Theme.BG_CARD,
        ).pack(anchor="w")

        log_frame = tk.Frame(log_inner, bg=Theme.BG_INPUT, bd=0, highlightthickness=1,
                             highlightbackground=Theme.BORDER, highlightcolor=Theme.BORDER_ACTIVE)
        log_frame.pack(fill="x", pady=(4, 8))

        self.log_text = tk.Text(
            log_frame,
            height=4,  # 세로 크기 축소
            bg=Theme.BG_INPUT,
            fg=Theme.TEXT_SECONDARY,
            font=Theme.FONT_LOG,
            bd=0,
            highlightthickness=0,
            relief="flat",
            state="disabled",
            wrap="word",
        )
        self.log_text.pack(side="left", fill="both", expand=True, padx=2, pady=2)

        log_scroll = tk.Scrollbar(log_frame, command=self.log_text.yview,
                                   bg=Theme.BG_INPUT, troughcolor=Theme.BG_INPUT)
        log_scroll.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=log_scroll.set)

        # 로그 태그 설정
        self.log_text.tag_configure("info", foreground=Theme.TEXT_SECONDARY)
        self.log_text.tag_configure("success", foreground=Theme.SUCCESS)
        self.log_text.tag_configure("error", foreground=Theme.ERROR)
        self.log_text.tag_configure("warning", foreground=Theme.WARNING)
        self.log_text.tag_configure("accent", foreground=Theme.ACCENT)

        # ── 프로그레스 바 ──
        progress_frame = tk.Frame(log_inner, bg=Theme.BG_CARD)
        progress_frame.pack(fill="x")

        self.progress_label = tk.Label(
            progress_frame,
            text="대기 중",
            font=Theme.FONT_SMALL,
            fg=Theme.TEXT_MUTED,
            bg=Theme.BG_CARD,
        )
        self.progress_label.pack(anchor="w")

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            style="Custom.Horizontal.TProgressbar",
        )
        self.progress_bar.pack(fill="x", pady=(4, 0))

        # ═══════════════════════════════════════════════════
        #  섹션 4: 하단 버튼
        # ═══════════════════════════════════════════════════
        bottom = tk.Frame(main, bg=Theme.BG_DARK)
        bottom.pack(fill="x")

        # 변환 시작 버튼
        self.convert_btn = tk.Button(
            bottom,
            text="🚀  변환 시작",
            font=("맑은 고딕", 11, "bold"),
            fg=Theme.TEXT_ON_ACCENT,
            bg=Theme.ACCENT,
            activeforeground=Theme.TEXT_ON_ACCENT,
            activebackground=Theme.ACCENT_HOVER,
            bd=0,
            padx=28,
            pady=8,
            cursor="hand2",
            command=self._start_conversion,
        )
        self.convert_btn.pack(side="right")
        self.convert_btn.bind("<Enter>", lambda e: self.convert_btn.config(bg=Theme.ACCENT_HOVER))
        self.convert_btn.bind("<Leave>", lambda e: self.convert_btn.config(bg=Theme.ACCENT))

    # ─────────────────────────────────────────────────────────
    #  UI 헬퍼 메서드
    # ─────────────────────────────────────────────────────────
    def _make_card(self, parent) -> tk.Frame:
        """둥근 느낌의 카드 패널을 생성합니다."""
        card = tk.Frame(
            parent,
            bg=Theme.BG_CARD,
            bd=0,
            highlightthickness=1,
            highlightbackground=Theme.BORDER,
            highlightcolor=Theme.BORDER,
        )
        return card

    def _make_button(self, parent, text, command, style="primary") -> tk.Button:
        """스타일이 적용된 버튼을 생성합니다."""
        if style == "primary":
            bg, fg, hover_bg = Theme.ACCENT, Theme.TEXT_ON_ACCENT, Theme.ACCENT_HOVER
        else:
            bg, fg, hover_bg = Theme.BG_INPUT, Theme.TEXT_SECONDARY, Theme.BG_HOVER

        btn = tk.Button(
            parent,
            text=text,
            font=Theme.FONT_BUTTON,
            fg=fg,
            bg=bg,
            activeforeground=fg,
            activebackground=hover_bg,
            bd=0,
            padx=14,
            pady=6,
            cursor="hand2",
            command=command,
        )
        btn.bind("<Enter>", lambda e, b=btn, h=hover_bg: b.config(bg=h))
        btn.bind("<Leave>", lambda e, b=btn, o=bg: b.config(bg=o))
        return btn

    def _show_drop_hint(self):
        """리스트가 비어있을 때 드래그 앤 드롭 안내를 표시합니다."""
        if len(self.pdf_files) == 0:
            hint = "여기에 PDF 파일을 끌어다 놓거나\n'파일 찾기' 버튼을 눌러 추가하세요" if DND_AVAILABLE \
                else "'파일 찾기' 버튼을 눌러 PDF 파일을 추가하세요"
            self.file_listbox.config(fg=Theme.TEXT_MUTED, justify="center")
            self.file_listbox.insert("end", "")
            self.file_listbox.insert("end", hint)
            self.file_listbox.config(state="disabled")
        else:
            self.file_listbox.config(state="normal")

    def _refresh_file_list(self):
        """파일 리스트를 갱신합니다."""
        self.file_listbox.config(state="normal", fg=Theme.TEXT_PRIMARY, justify="left")
        self.file_listbox.delete(0, "end")
        if not self.pdf_files:
            self._show_drop_hint()
        else:
            for f in self.pdf_files:
                display = f"  📄  {os.path.basename(f)}    [{os.path.dirname(f)}]"
                self.file_listbox.insert("end", display)
        self.file_count_label.config(text=f"{len(self.pdf_files)}개 파일")

    # ─────────────────────────────────────────────────────────
    #  이벤트 핸들러
    # ─────────────────────────────────────────────────────────
    def _on_drop(self, event):
        """드래그 앤 드롭으로 파일 추가."""
        if self.conversion_just_finished:
            self.pdf_files.clear()
            self.conversion_just_finished = False

        raw = event.data
        # tkinterdnd2 는 {} 나 공백으로 구분된 경로를 반환
        files = self._parse_dnd_data(raw)
        added = 0
        for f in files:
            f = f.strip().strip('"').strip("'")
            if f.lower().endswith(".pdf") and os.path.isfile(f) and f not in self.pdf_files:
                self.pdf_files.append(f)
                added += 1
        if added > 0:
            self._refresh_file_list()
            self._log(f"✅ {added}개 파일이 추가되었습니다.", "success")

    @staticmethod
    def _parse_dnd_data(data: str) -> list[str]:
        """tkinterdnd2 의 드롭 데이터를 파싱합니다."""
        files = []
        i = 0
        while i < len(data):
            if data[i] == '{':
                j = data.index('}', i)
                files.append(data[i+1:j])
                i = j + 1
            elif data[i] == ' ':
                i += 1
            else:
                j = data.find(' ', i)
                if j == -1:
                    files.append(data[i:])
                    break
                files.append(data[i:j])
                i = j + 1
        return files

    def _browse_files(self):
        """파일 탐색기를 열어 PDF 파일을 선택합니다."""
        if self.conversion_just_finished:
            self.pdf_files.clear()
            self.conversion_just_finished = False

        filepaths = filedialog.askopenfilenames(
            title="PDF 파일 선택",
            filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")],
        )
        added = 0
        for f in filepaths:
            if f not in self.pdf_files:
                self.pdf_files.append(f)
                added += 1
        if added > 0:
            self._refresh_file_list()
            self._log(f"✅ {added}개 파일이 추가되었습니다.", "success")

    def _remove_selected(self):
        """선택된 파일을 목록에서 제거합니다."""
        self.conversion_just_finished = False
        selection = self.file_listbox.curselection()
        if not selection:
            return
        for idx in reversed(selection):
            if 0 <= idx < len(self.pdf_files):
                self.pdf_files.pop(idx)
        self._refresh_file_list()
        self._log("🗑️ 선택된 파일이 삭제되었습니다.", "info")

    def _clear_files(self):
        """모든 파일을 목록에서 제거합니다."""
        self.conversion_just_finished = False
        self.pdf_files.clear()
        self._refresh_file_list()
        self._log("🧹 파일 목록이 초기화되었습니다.", "info")

    def _browse_output_dir(self):
        """저장 폴더를 사용자 지정으로 변경합니다."""
        dirpath = filedialog.askdirectory(title="변환 파일 저장 폴더 선택")
        if dirpath:
            self.custom_output_dir = dirpath
            self.save_path_var.set(dirpath)
            self._log(f"💾 저장 경로 변경: {dirpath}", "info")

    def _reset_output_dir(self):
        """저장 경로를 기본값(원본 위치)으로 복원합니다."""
        self.custom_output_dir = None
        self.save_path_var.set("원본 PDF와 같은 폴더에 저장 (기본)")
        self._log("↩️ 저장 경로가 기본값으로 복원되었습니다.", "info")

    # ─────────────────────────────────────────────────────────
    #  로그
    # ─────────────────────────────────────────────────────────
    def _log(self, message: str, tag: str = "info"):
        """로그 창에 메시지를 추가합니다."""
        self.log_text.config(state="normal")
        timestamp = time.strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{timestamp}] {message}\n", tag)
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    # ─────────────────────────────────────────────────────────
    #  변환 처리
    # ─────────────────────────────────────────────────────────
    def _start_conversion(self):
        """변환 작업을 시작합니다."""
        if self.is_converting:
            messagebox.showwarning("⚠️ 알림", "이미 변환 작업이 진행 중입니다.")
            return

        if not self.pdf_files:
            messagebox.showwarning("⚠️ 알림", "변환할 PDF 파일을 먼저 추가해주세요.")
            return

        self.is_converting = True
        self.convert_btn.config(text="⏳  변환 중...", bg=Theme.TEXT_MUTED, state="disabled")
        self.progress_var.set(0)
        self._log("🚀 변환 작업을 시작합니다...", "accent")

        # 백그라운드 스레드에서 실행
        thread = threading.Thread(target=self._conversion_worker, daemon=True)
        thread.start()

    def _conversion_worker(self):
        """백그라운드 스레드에서 PDF→MD 변환을 수행합니다."""
        total = len(self.pdf_files)
        success_count = 0
        fail_count = 0

        if self.custom_output_dir:
            # ── 사용자 지정 폴더: 일괄 배치 처리 (JVM 1회) ──
            self.root.after(0, self._log, f"📦 일괄 변환 모드 ({total}개 파일 → {self.custom_output_dir})", "info")
            try:
                opendataloader_pdf.convert(
                    input_path=list(self.pdf_files),
                    output_dir=self.custom_output_dir,
                    format="markdown",
                    image_output="external",
                    image_format="png",
                    quiet=True,
                )
                success_count = total
                for i in range(total):
                    name = os.path.basename(self.pdf_files[i])
                    self.root.after(0, self._log, f"  ✅ {name} — 변환 완료", "success")
                    pct = ((i + 1) / total) * 100
                    self.root.after(0, self._update_progress, pct, f"{i + 1}/{total} 완료")
            except Exception as e:
                fail_count = total
                self.root.after(0, self._log, f"  ❌ 일괄 변환 실패: {e}", "error")
        else:
            # ── 원본 위치 저장: 폴더별 그룹화 후 배치 처리 ──
            groups: dict[str, list[str]] = defaultdict(list)
            for f in self.pdf_files:
                groups[os.path.dirname(f)].append(f)

            processed = 0
            for out_dir, files in groups.items():
                group_label = out_dir if len(out_dir) < 50 else f"...{out_dir[-47:]}"
                self.root.after(0, self._log, f"📂 폴더: {group_label} ({len(files)}개)", "info")
                try:
                    opendataloader_pdf.convert(
                        input_path=files,
                        output_dir=out_dir,
                        format="markdown",
                        image_output="external",
                        image_format="png",
                        quiet=True,
                    )
                    for f in files:
                        name = os.path.basename(f)
                        self.root.after(0, self._log, f"  ✅ {name} — 변환 완료", "success")
                        processed += 1
                        pct = (processed / total) * 100
                        self.root.after(0, self._update_progress, pct, f"{processed}/{total} 완료")
                    success_count += len(files)
                except Exception as e:
                    for f in files:
                        name = os.path.basename(f)
                        self.root.after(0, self._log, f"  ❌ {name} — 실패: {e}", "error")
                        processed += 1
                        pct = (processed / total) * 100
                        self.root.after(0, self._update_progress, pct, f"{processed}/{total} 완료")
                    fail_count += len(files)

        # ── 완료 ──
        summary = f"🏁 변환 완료! 성공: {success_count}개"
        if fail_count > 0:
            summary += f", 실패: {fail_count}개"
        self.root.after(0, self._log, summary, "accent")

        # 폴더 열기 설정 체크 및 마지막 폴더 오픈
        if self.auto_open_var.get() and success_count > 0:
            if self.custom_output_dir and os.path.exists(self.custom_output_dir):
                self.root.after(0, os.startfile, self.custom_output_dir)
            elif len(self.pdf_files) > 0:
                last_out_dir = os.path.dirname(self.pdf_files[-1])
                if os.path.exists(last_out_dir):
                    self.root.after(0, os.startfile, last_out_dir)

        self.root.after(0, self._conversion_done)

    def _update_progress(self, pct: float, text: str):
        """프로그레스 바를 갱신합니다."""
        self.progress_var.set(pct)
        self.progress_label.config(text=text, fg=Theme.TEXT_PRIMARY)

    def _conversion_done(self):
        """변환 완료 후 UI 상태를 복원합니다."""
        self.is_converting = False
        self.conversion_just_finished = True
        self.convert_btn.config(text="🚀  변환 시작", bg=Theme.ACCENT, state="normal")
        self.progress_label.config(text="완료 ✨", fg=Theme.SUCCESS)

    # ─────────────────────────────────────────────────────────
    #  실행
    # ─────────────────────────────────────────────────────────
    def run(self):
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════════
#  엔트리 포인트
# ═══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = PDFtoMDApp()
    app.run()

"""MSS SKU 매칭 및 발주 파일 변환기 — GUI 앱 (UX 개선 버전)

개선사항:
- [P1] 작업 이력 대시보드
- [P2] 출력 미리보기
- [P3] 프리셋 저장/로드
- [P4] 매칭 실패 리포트 내보내기
- [P5] 파일 덮어쓰기 보호
- [D1] 시각적 진행 상태 (프로그레스 바)
- [D2] 파일 상태 인디케이터
- [D3] 브랜드 선택 영역 확대
- [D4] 에러 메시지 사용자 친화적 변환
- [D5] 드래그앤드롭 파일 입력
- [E1] 세션 상태 자동 저장/복원
- [E2] 파일 스키마 검증
- [E3] MouseWheel 스코프 수정
- [E4] 반응형 창 크기
- [E5] 취소 기능
- 리셋 기능, 툴팁 추가
"""
import sys
import os
import threading
from datetime import datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

import pandas as pd

from core.loader import load_order_data, load_matching_data, load_option_data, get_brand_list, filter_by_brand
from core.matcher import match_barcode_to_uid, detect_option_products, match_option_info
from core.generator import generate_system_upload, generate_brand_order
from core.merger import merge_order_files
from core.ui_helpers import ToolTip, friendly_error, validate_file_schema, FileStatusLabel
from core.session import (
    save_session, load_session, add_history_entry, get_history,
    save_preset, get_presets, load_preset, delete_preset,
)

# 드래그앤드롭 (선택적 의존성)
# tkinterdnd2는 __file__ 기준으로 tkdnd DLL을 찾으므로
# PyInstaller hook이 올바르게 번들링하면 별도 경로 설정 불필요
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except (ImportError, RuntimeError, tk.TclError):
    HAS_DND = False


def _create_root():
    """루트 윈도우 생성 (DnD 지원 여부에 따라 분기)"""
    if HAS_DND:
        return TkinterDnD.Tk()
    return tk.Tk()


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MSS SKU 매칭 및 발주 파일 변환기")
        self.root.geometry("780x900")
        # [E4] 반응형 창 크기
        self.root.resizable(True, True)
        self.root.minsize(720, 800)

        # [E5] 취소 플래그
        self._cancel_event = threading.Event()

        self._build_ui()
        # [E1] 세션 복원
        self._restore_session()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main_frame, text="MSS SKU 매칭 및 발주 파일 변환기", font=("", 15, "bold"))
        title.pack(pady=(0, 10))

        # 탭
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._build_tab_a()
        self._build_tab_b()
        self._build_tab_history()

    # ================================================================
    #  탭 A: SKU 매칭 변환
    # ================================================================
    def _build_tab_a(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  A. SKU 매칭 변환  ")

        # 상태
        self.a_order_df = None
        self.a_matching_df = None
        self.a_option_df = None
        self.a_file_paths = {1: None, 2: None, 3: None}
        self.a_all_brands = []
        self.a_save_path = str(Path.home() / "Desktop")
        self.a_last_warnings = []  # [P4] 매칭 실패 리포트용

        # ── 프리셋 영역 [P3] ──
        preset_frame = ttk.Frame(tab)
        preset_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(preset_frame, text="프리셋:", font=("", 9)).pack(side=tk.LEFT)
        self.a_preset_var = tk.StringVar()
        self.a_preset_combo = ttk.Combobox(
            preset_frame, textvariable=self.a_preset_var,
            state='readonly', width=20,
        )
        self.a_preset_combo.pack(side=tk.LEFT, padx=(5, 5))
        ttk.Button(preset_frame, text="불러오기", command=self._a_load_preset, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="저장", command=self._a_save_preset, width=6).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="삭제", command=self._a_delete_preset, width=6).pack(side=tk.LEFT, padx=2)
        ToolTip(self.a_preset_combo, "자주 쓰는 파일+브랜드 조합을 프리셋으로 저장하세요")
        self._a_refresh_presets()

        # ── 파일 입력 ──
        file_frame = ttk.LabelFrame(tab, text="입력 파일", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        file_labels = [
            ("1. 총 발주 수량 데이터셋", 1, "88코드, 브랜드명, 수량(오프) 등이 포함된 발주 데이터"),
            ("2. 상품코드-바코드 매칭 데이터셋", 2, "바코드 → 상품코드(UID), 스타일번호 매핑 파일"),
            ("3. 상품별 옵션 정보", 3, "사이즈유형, Size01~Size30 옵션 정보 파일"),
        ]
        self.a_file_entries = {}
        self.a_file_status = {}  # [D2] 파일 상태 인디케이터
        for label_text, fnum, tooltip_text in file_labels:
            row = ttk.Frame(file_frame)
            row.pack(fill=tk.X, pady=2)
            lbl = ttk.Label(row, text=label_text, width=30, anchor=tk.W)
            lbl.pack(side=tk.LEFT)
            ToolTip(lbl, tooltip_text)

            entry = ttk.Entry(row, state='readonly')
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
            self.a_file_entries[fnum] = entry

            # [D5] 드래그앤드롭
            if HAS_DND:
                entry.drop_target_register(DND_FILES)
                entry.dnd_bind('<<Drop>>', lambda e, n=fnum: self._a_on_drop(e, n))

            btn = ttk.Button(row, text="파일 선택",
                             command=lambda n=fnum: self._a_select_file(n))
            btn.pack(side=tk.RIGHT)

            # [D2] 상태 표시
            status = FileStatusLabel(row)
            status.pack(side=tk.RIGHT, padx=(0, 5))
            self.a_file_status[fnum] = status

        # [D5] 드래그앤드롭 안내 (DnD 가능 시)
        if HAS_DND:
            dnd_label = ttk.Label(file_frame, text="파일을 끌어다 놓을 수도 있습니다",
                                  font=("", 8), foreground="#999")
            dnd_label.pack(anchor=tk.W, pady=(2, 0))

        # ── 브랜드 선택 ──
        brand_frame = ttk.LabelFrame(tab, text="브랜드 선택", padding=8)
        brand_frame.pack(fill=tk.X, pady=(0, 8))

        # 모드 선택
        radio_row = ttk.Frame(brand_frame)
        radio_row.pack(fill=tk.X, pady=(0, 3))
        self.a_brand_mode = tk.StringVar(value='all')
        ttk.Radiobutton(radio_row, text="모든 브랜드", variable=self.a_brand_mode,
                        value='all', command=self._a_on_brand_mode).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Radiobutton(radio_row, text="브랜드 선택", variable=self.a_brand_mode,
                        value='select', command=self._a_on_brand_mode).pack(side=tk.LEFT)

        # 선택된 브랜드 표시 영역
        self.a_selected_frame = ttk.Frame(brand_frame)
        self.a_selected_frame.pack(fill=tk.X, pady=(3, 0))
        self.a_selected_label = ttk.Label(self.a_selected_frame, text="", font=("", 9),
                                          foreground="#2563eb", wraplength=650)
        self.a_selected_label.pack(fill=tk.X)

        # 검색
        search_row = ttk.Frame(brand_frame)
        search_row.pack(fill=tk.X, pady=(3, 3))
        ttk.Label(search_row, text="검색:", width=5, anchor=tk.W).pack(side=tk.LEFT)
        self.a_search_var = tk.StringVar()
        self.a_search_var.trace_add('write', lambda *_: self._a_render_checkboxes())
        self.a_search_entry = ttk.Entry(search_row, textvariable=self.a_search_var, state='disabled')
        self.a_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 전체선택/해제 버튼
        btn_row = ttk.Frame(brand_frame)
        btn_row.pack(fill=tk.X, pady=(0, 3))
        self.a_select_all_btn = ttk.Button(btn_row, text="전체 선택",
                                           command=lambda: self._a_toggle_all(True), state='disabled')
        self.a_select_all_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.a_deselect_all_btn = ttk.Button(btn_row, text="전체 해제",
                                             command=lambda: self._a_toggle_all(False), state='disabled')
        self.a_deselect_all_btn.pack(side=tk.LEFT)

        # [D3] 체크박스 리스트 — 높이 확대 (100 → 180)
        self.a_cb_canvas = tk.Canvas(brand_frame, height=180, highlightthickness=0)
        self.a_cb_scrollbar = ttk.Scrollbar(brand_frame, orient=tk.VERTICAL, command=self.a_cb_canvas.yview)
        self.a_cb_inner = ttk.Frame(self.a_cb_canvas)
        self.a_cb_inner.bind('<Configure>',
                             lambda e: self.a_cb_canvas.configure(scrollregion=self.a_cb_canvas.bbox('all')))
        self.a_cb_canvas.create_window((0, 0), window=self.a_cb_inner, anchor='nw')
        self.a_cb_canvas.configure(yscrollcommand=self.a_cb_scrollbar.set)
        self.a_cb_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.a_cb_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # [E3] 마우스 휠 — 캔버스 스코프 바인딩 (bind_all 대신)
        self.a_cb_canvas.bind('<Enter>', lambda e: self._bind_mousewheel(self.a_cb_canvas))
        self.a_cb_canvas.bind('<Leave>', lambda e: self._unbind_mousewheel())

        self.a_brand_vars = {}
        self.a_cb_widgets = []

        # ── 저장 위치 ──
        save_frame = ttk.LabelFrame(tab, text="저장 위치", padding=8)
        save_frame.pack(fill=tk.X, pady=(0, 8))
        save_row = ttk.Frame(save_frame)
        save_row.pack(fill=tk.X)
        self.a_save_entry = ttk.Entry(save_row, state='readonly')
        self.a_save_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.a_save_entry.configure(state='normal')
        self.a_save_entry.insert(0, self.a_save_path)
        self.a_save_entry.configure(state='readonly')
        ttk.Button(save_row, text="폴더 선택",
                   command=lambda: self._select_folder('a')).pack(side=tk.RIGHT)

        # ── 실행 버튼 영역 ──
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=8)
        self.a_run_btn = ttk.Button(btn_frame, text="변환 실행", command=self._a_run)
        self.a_run_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.a_run_btn, "선택한 브랜드에 대해 SKU 매칭 변환을 실행합니다")

        # [E5] 취소 버튼
        self.a_cancel_btn = ttk.Button(btn_frame, text="취소", command=self._a_cancel, state='disabled')
        self.a_cancel_btn.pack(side=tk.LEFT, padx=5)

        # 리셋 버튼
        self.a_reset_btn = ttk.Button(btn_frame, text="초기화", command=self._a_reset)
        self.a_reset_btn.pack(side=tk.LEFT, padx=5)
        ToolTip(self.a_reset_btn, "모든 입력을 초기 상태로 되돌립니다")

        # [D1] 프로그레스 바
        progress_frame = ttk.Frame(tab)
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        self.a_progress = ttk.Progressbar(progress_frame, mode='determinate')
        self.a_progress.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 8))
        self.a_progress_label = ttk.Label(progress_frame, text="", font=("", 9), width=30, anchor=tk.W)
        self.a_progress_label.pack(side=tk.LEFT)

        # ── 로그 ──
        log_frame = ttk.LabelFrame(tab, text="처리 결과", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.a_log = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
        self.a_log.pack(fill=tk.BOTH, expand=True)

        # 하단 버튼
        bottom_row = ttk.Frame(tab)
        bottom_row.pack(fill=tk.X, pady=(5, 0))
        self.a_open_btn = ttk.Button(bottom_row, text="출력 폴더 열기",
                                     command=lambda: self._open_folder(self.a_save_path), state='disabled')
        self.a_open_btn.pack(side=tk.LEFT, padx=(0, 5))

        # [P4] 매칭 실패 리포트 내보내기
        self.a_report_btn = ttk.Button(bottom_row, text="매칭 실패 리포트",
                                       command=self._a_export_warning_report, state='disabled')
        self.a_report_btn.pack(side=tk.LEFT)
        ToolTip(self.a_report_btn, "옵션 매칭 경고/미매칭 바코드를 Excel로 내보냅니다")

    # ================================================================
    #  탭 B: 발주 파일 통합
    # ================================================================
    def _build_tab_b(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  B. 발주 파일 통합  ")

        self.b_file_list = []
        self.b_matching_path = None
        self.b_save_path = str(Path.home() / "Desktop")

        # 발주 파일 목록
        file_frame = ttk.LabelFrame(tab, text="발주 파일 (복수 선택)", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        self.b_file_listbox = tk.Listbox(file_frame, height=5, font=("", 9))
        self.b_file_listbox.pack(fill=tk.X, expand=True)

        # [D5] 드래그앤드롭
        if HAS_DND:
            self.b_file_listbox.drop_target_register(DND_FILES)
            self.b_file_listbox.dnd_bind('<<Drop>>', self._b_on_drop)

        # 리스트 툴팁 (파일 전체 경로 표시)
        self._b_listbox_tooltip = ToolTip(self.b_file_listbox, "")
        self.b_file_listbox.bind('<Motion>', self._b_update_listbox_tooltip)

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_row, text="+ 파일 추가", command=self._b_add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_row, text="- 선택 제거", command=self._b_remove_file).pack(side=tk.LEFT)

        if HAS_DND:
            ttk.Label(btn_row, text="  (파일을 끌어다 놓을 수도 있습니다)",
                      font=("", 8), foreground="#999").pack(side=tk.LEFT, padx=10)

        # 매칭 파일 (선택)
        match_frame = ttk.LabelFrame(tab, text="상품코드-바코드 매칭 파일 (선택)", padding=8)
        match_frame.pack(fill=tk.X, pady=(0, 8))
        match_row = ttk.Frame(match_frame)
        match_row.pack(fill=tk.X)
        self.b_match_entry = ttk.Entry(match_row, state='readonly')
        self.b_match_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(match_row, text="파일 선택", command=self._b_select_matching).pack(side=tk.RIGHT)
        ToolTip(self.b_match_entry, "선택하면 88코드 → 상품번호(UID)를 자동으로 채워줍니다")

        # 저장 위치
        save_frame = ttk.LabelFrame(tab, text="저장 위치", padding=8)
        save_frame.pack(fill=tk.X, pady=(0, 8))
        save_row = ttk.Frame(save_frame)
        save_row.pack(fill=tk.X)
        self.b_save_entry = ttk.Entry(save_row, state='readonly')
        self.b_save_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.b_save_entry.configure(state='normal')
        self.b_save_entry.insert(0, self.b_save_path)
        self.b_save_entry.configure(state='readonly')
        ttk.Button(save_row, text="폴더 선택",
                   command=lambda: self._select_folder('b')).pack(side=tk.RIGHT)

        # 실행 버튼 영역
        btn_frame = ttk.Frame(tab)
        btn_frame.pack(pady=8)
        self.b_run_btn = ttk.Button(btn_frame, text="통합 실행", command=self._b_run)
        self.b_run_btn.pack(side=tk.LEFT, padx=5)

        # [E5] 취소 버튼
        self.b_cancel_btn = ttk.Button(btn_frame, text="취소", command=self._b_cancel, state='disabled')
        self.b_cancel_btn.pack(side=tk.LEFT, padx=5)

        # 리셋 버튼
        ttk.Button(btn_frame, text="초기화", command=self._b_reset).pack(side=tk.LEFT, padx=5)

        # [D1] 프로그레스 바
        progress_frame = ttk.Frame(tab)
        progress_frame.pack(fill=tk.X, pady=(0, 5))
        self.b_progress = ttk.Progressbar(progress_frame, mode='indeterminate')
        self.b_progress.pack(fill=tk.X, side=tk.LEFT, expand=True, padx=(0, 8))
        self.b_progress_label = ttk.Label(progress_frame, text="", font=("", 9), width=30, anchor=tk.W)
        self.b_progress_label.pack(side=tk.LEFT)

        # 로그
        log_frame = ttk.LabelFrame(tab, text="처리 결과", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.b_log = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
        self.b_log.pack(fill=tk.BOTH, expand=True)

        self.b_open_btn = ttk.Button(tab, text="출력 폴더 열기",
                                     command=lambda: self._open_folder(self.b_save_path), state='disabled')
        self.b_open_btn.pack(pady=(5, 0))

    # ================================================================
    #  [P1] 탭 C: 작업 이력
    # ================================================================
    def _build_tab_history(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="  작업 이력  ")

        top_row = ttk.Frame(tab)
        top_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(top_row, text="최근 작업 이력", font=("", 12, "bold")).pack(side=tk.LEFT)
        ttk.Button(top_row, text="새로고침", command=self._refresh_history).pack(side=tk.RIGHT, padx=5)
        ttk.Button(top_row, text="이력 삭제", command=self._clear_history).pack(side=tk.RIGHT)

        # 이력 테이블
        columns = ('time', 'type', 'brands', 'result', 'warnings')
        self.history_tree = ttk.Treeview(tab, columns=columns, show='headings', height=15)
        self.history_tree.heading('time', text='시간')
        self.history_tree.heading('type', text='유형')
        self.history_tree.heading('brands', text='브랜드')
        self.history_tree.heading('result', text='결과')
        self.history_tree.heading('warnings', text='경고')
        self.history_tree.column('time', width=140)
        self.history_tree.column('type', width=60)
        self.history_tree.column('brands', width=300)
        self.history_tree.column('result', width=100)
        self.history_tree.column('warnings', width=70)

        scrollbar = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self._refresh_history()

    def _refresh_history(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for entry in get_history():
            brands = ', '.join(entry.get('brands', []))
            if len(brands) > 40:
                brands = brands[:37] + '...'
            self.history_tree.insert('', tk.END, values=(
                entry.get('timestamp', ''),
                '매칭 변환' if entry.get('type') == 'A' else '파일 통합',
                brands,
                f"성공 {entry.get('success', 0)} / 실패 {entry.get('fail', 0)}",
                entry.get('warnings', 0),
            ))

    def _clear_history(self):
        if messagebox.askyesno("확인", "모든 작업 이력을 삭제하시겠습니까?"):
            from core.session import clear_history
            clear_history()
            self._refresh_history()

    # ================================================================
    #  공통 헬퍼
    # ================================================================
    def _log(self, widget, msg: str):
        widget.configure(state='normal')
        widget.insert(tk.END, msg + "\n")
        widget.see(tk.END)
        widget.configure(state='disabled')

    def _clear_log(self, widget):
        widget.configure(state='normal')
        widget.delete('1.0', tk.END)
        widget.configure(state='disabled')

    def _select_folder(self, tab_id: str):
        folder = filedialog.askdirectory(title="저장 폴더 선택")
        if not folder:
            return
        if tab_id == 'a':
            self.a_save_path = folder
            entry = self.a_save_entry
        else:
            self.b_save_path = folder
            entry = self.b_save_entry
        entry.configure(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, folder)
        entry.configure(state='readonly')
        self._save_session()

    def _open_folder(self, path):
        if not path:
            return
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')

    def _set_entry(self, entry: ttk.Entry, value: str):
        entry.configure(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, value)
        entry.configure(state='readonly')

    # [E3] 마우스 휠 스코프 바인딩
    def _bind_mousewheel(self, canvas: tk.Canvas):
        if sys.platform == 'darwin':
            canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1 * e.delta), "units"))
        else:
            canvas.bind_all('<MouseWheel>', lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _unbind_mousewheel(self):
        self.root.unbind_all('<MouseWheel>')

    # [E1] 세션 저장
    def _save_session(self):
        state = {
            'tab_a': {
                'file1': self.a_file_paths.get(1),
                'file2': self.a_file_paths.get(2),
                'file3': self.a_file_paths.get(3),
                'save_path': self.a_save_path,
                'brand_mode': self.a_brand_mode.get(),
                'selected_brands': [b for b, v in self.a_brand_vars.items() if v.get()],
            },
            'tab_b': {
                'files': list(self.b_file_list),
                'matching_file': self.b_matching_path,
                'save_path': self.b_save_path,
            },
            'last_tab': self.notebook.index(self.notebook.select()) if self.notebook.select() else 0,
        }
        try:
            save_session(state)
        except Exception:
            pass  # 세션 저장 실패는 무시

    # [E1] 세션 복원
    def _restore_session(self):
        state = load_session()
        if not state:
            return

        tab_a = state.get('tab_a', {})
        tab_b = state.get('tab_b', {})

        # 탭 A 파일 복원
        for fnum, key in [(1, 'file1'), (2, 'file2'), (3, 'file3')]:
            fp = tab_a.get(key)
            if fp and Path(fp).exists():
                self.a_file_paths[fnum] = fp
                self._set_entry(self.a_file_entries[fnum], fp)
                if fnum == 1:
                    self._a_load_brands(fp)
                else:
                    # 비동기 검증
                    threading.Thread(
                        target=self._a_validate_file, args=(fnum, fp), daemon=True
                    ).start()

        # 저장 경로 복원
        if tab_a.get('save_path') and Path(tab_a['save_path']).exists():
            self.a_save_path = tab_a['save_path']
            self._set_entry(self.a_save_entry, self.a_save_path)

        # 브랜드 선택 복원
        if tab_a.get('brand_mode'):
            self.a_brand_mode.set(tab_a['brand_mode'])
            self._a_on_brand_mode()
        saved_brands = tab_a.get('selected_brands', [])
        if saved_brands:
            for b in saved_brands:
                if b in self.a_brand_vars:
                    self.a_brand_vars[b].set(True)
            self._a_update_selected_label()

        # 탭 B 복원
        for fp in tab_b.get('files', []):
            if Path(fp).exists() and fp not in self.b_file_list:
                self.b_file_list.append(fp)
                self.b_file_listbox.insert(tk.END, Path(fp).name)

        match_fp = tab_b.get('matching_file')
        if match_fp and Path(match_fp).exists():
            self.b_matching_path = match_fp
            self._set_entry(self.b_match_entry, match_fp)

        if tab_b.get('save_path') and Path(tab_b['save_path']).exists():
            self.b_save_path = tab_b['save_path']
            self._set_entry(self.b_save_entry, self.b_save_path)

        # 마지막 탭 복원
        last_tab = state.get('last_tab', 0)
        try:
            self.notebook.select(last_tab)
        except Exception:
            pass

    # ================================================================
    #  탭 A 로직
    # ================================================================

    # [D5] 드래그앤드롭 처리
    def _a_on_drop(self, event, fnum):
        fp = event.data.strip('{}')  # Windows에서 중괄호로 감싸는 경우 처리
        if fp:
            self._a_set_file(fnum, fp)

    def _a_select_file(self, fnum):
        fp = filedialog.askopenfilename(
            title=f"파일 {fnum} 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        if not fp:
            return
        self._a_set_file(fnum, fp)

    def _a_set_file(self, fnum, fp):
        self.a_file_paths[fnum] = fp
        self._set_entry(self.a_file_entries[fnum], fp)
        # [D2/E2] 파일 검증
        self.a_file_status[fnum].set_status('loading', '검증 중...')
        if fnum == 1:
            self._a_load_brands(fp)
        else:
            threading.Thread(
                target=self._a_validate_file, args=(fnum, fp), daemon=True
            ).start()
        self._save_session()

    def _a_validate_file(self, fnum, fp):
        """[E2] 파일 스키마 검증 (백그라운드)"""
        result = validate_file_schema(fp, fnum)
        self.root.after(0, lambda: self._a_update_file_status(fnum, result))

    def _a_update_file_status(self, fnum, result):
        status = 'valid' if result['valid'] else 'invalid'
        self.a_file_status[fnum].set_status(status, result['message'])

    def _a_load_brands(self, filepath):
        try:
            self.a_order_df = load_order_data(filepath)
            self.a_all_brands = get_brand_list(self.a_order_df)

            self.a_brand_vars = {}
            for b in self.a_all_brands:
                self.a_brand_vars[b] = tk.BooleanVar(value=False)

            self._a_on_brand_mode()
            self._a_render_checkboxes()

            # 파일 상태 업데이트
            row_count = len(self.a_order_df)
            self.a_file_status[1].set_status('valid', f"{row_count}행, {len(self.a_all_brands)}개 브랜드")
            self._log(self.a_log, f"브랜드 로드 완료 ({len(self.a_all_brands)}개): {', '.join(self.a_all_brands)}")
        except Exception as e:
            self.a_file_status[1].set_status('invalid', friendly_error(e)[:50])
            messagebox.showerror("오류", f"파일 1 로드 실패:\n{friendly_error(e)}")

    def _a_on_brand_mode(self):
        is_select = self.a_brand_mode.get() == 'select'
        state = 'normal' if is_select else 'disabled'
        self.a_search_entry.configure(state=state)
        self.a_select_all_btn.configure(state=state)
        self.a_deselect_all_btn.configure(state=state)

        if not is_select:
            self.a_search_var.set('')
            self.a_selected_label.configure(text="")
        else:
            self._a_update_selected_label()

        self._a_render_checkboxes()

    def _a_render_checkboxes(self):
        for w in self.a_cb_widgets:
            w.destroy()
        self.a_cb_widgets.clear()

        q = self.a_search_var.get().strip().lower()
        is_select = self.a_brand_mode.get() == 'select'

        for b in self.a_all_brands:
            if q and q not in b.lower():
                continue
            var = self.a_brand_vars.get(b)
            if var is None:
                continue
            cb = ttk.Checkbutton(
                self.a_cb_inner, text=b, variable=var,
                command=self._a_update_selected_label,
            )
            if not is_select:
                cb.configure(state='disabled')
            cb.pack(anchor=tk.W, padx=5, pady=1)
            self.a_cb_widgets.append(cb)

        self.a_cb_inner.update_idletasks()
        self.a_cb_canvas.configure(scrollregion=self.a_cb_canvas.bbox('all'))

    def _a_update_selected_label(self):
        selected = [b for b, v in self.a_brand_vars.items() if v.get()]
        if selected:
            self.a_selected_label.configure(
                text=f"선택됨 ({len(selected)}개): {', '.join(selected)}")
        else:
            self.a_selected_label.configure(text="")

    def _a_toggle_all(self, value: bool):
        q = self.a_search_var.get().strip().lower()
        for b, var in self.a_brand_vars.items():
            if not q or q in b.lower():
                var.set(value)
        self._a_update_selected_label()

    def _a_get_brands(self):
        if self.a_brand_mode.get() == 'all':
            return list(self.a_all_brands)
        return [b for b, v in self.a_brand_vars.items() if v.get()]

    # [P2] 출력 미리보기
    def _a_show_preview(self, brands) -> bool:
        """실행 전 미리보기 다이얼로그. True 반환 시 실행 진행."""
        lines = [f"처리할 브랜드: {len(brands)}개\n"]
        for i, b in enumerate(brands, 1):
            if i <= 10:
                lines.append(f"  {i}. {b}")
            elif i == 11:
                lines.append(f"  ... 외 {len(brands) - 10}개")
                break

        lines.append(f"\n출력 파일 (브랜드당 2개):")
        sample = brands[0] if brands else '브랜드'
        lines.append(f"  - {{브랜드명}}_시스템업로드_최종파일.xlsx")
        lines.append(f"  - {{브랜드명}}_발주리스트_최종파일.xlsx")
        lines.append(f"\n총 {len(brands) * 2}개 파일 생성 예정")
        lines.append(f"저장 위치: {self.a_save_path}")

        # [P5] 덮어쓰기 경고
        existing = []
        for b in brands:
            sp = Path(self.a_save_path) / f'{b}_시스템업로드_최종파일.xlsx'
            bp = Path(self.a_save_path) / f'{b}_발주리스트_최종파일.xlsx'
            if sp.exists():
                existing.append(sp.name)
            if bp.exists():
                existing.append(bp.name)
        if existing:
            lines.append(f"\n!! 덮어쓰기 경고: {len(existing)}개 파일이 이미 존재합니다")
            for f in existing[:5]:
                lines.append(f"  - {f}")
            if len(existing) > 5:
                lines.append(f"  ... 외 {len(existing) - 5}개")

        return messagebox.askokcancel("실행 확인", "\n".join(lines))

    def _a_run(self):
        for n in [1, 2, 3]:
            if not self.a_file_paths[n]:
                messagebox.showwarning("입력 필요", f"파일 {n}을 선택해주세요.")
                return
        brands = self._a_get_brands()
        if not brands:
            messagebox.showwarning("입력 필요", "브랜드를 1개 이상 선택해주세요.")
            return

        # [P2] 미리보기
        if not self._a_show_preview(brands):
            return

        self._cancel_event.clear()
        self.a_run_btn.configure(state='disabled')
        self.a_cancel_btn.configure(state='normal')
        self.a_open_btn.configure(state='disabled')
        self.a_report_btn.configure(state='disabled')
        self._clear_log(self.a_log)
        self.a_last_warnings = []

        # [D1] 프로그레스 바 초기화
        self.a_progress['maximum'] = len(brands)
        self.a_progress['value'] = 0
        self.a_progress_label.configure(text=f"준비 중... (0/{len(brands)})")

        self._log(self.a_log, f"변환 시작... ({len(brands)}개 브랜드)\n")
        threading.Thread(target=self._a_do_run, args=(brands,), daemon=True).start()

    # [E5] 취소
    def _a_cancel(self):
        self._cancel_event.set()
        self.a_cancel_btn.configure(state='disabled')
        self._log(self.a_log, "\n취소 요청됨... 현재 브랜드 처리 완료 후 중단됩니다.")

    def _a_do_run(self, brands):
        log = self.a_log
        all_warnings = []
        try:
            self._log(log, "파일 로딩 중...")
            if self.a_order_df is None:
                self.a_order_df = load_order_data(self.a_file_paths[1])
            self.a_matching_df = load_matching_data(self.a_file_paths[2])
            self.a_option_df = load_option_data(self.a_file_paths[3])

            ok, fail = 0, 0
            for i, brand in enumerate(brands, 1):
                # [E5] 취소 확인
                if self._cancel_event.is_set():
                    self._log(log, f"\n{'='*45}\n취소됨: {ok}건 완료, {len(brands) - i + 1}건 건너뜀")
                    break

                # [D1] 프로그레스 업데이트
                self.root.after(0, lambda v=i, b=brand: self._a_update_progress(v, b, len(brands)))

                self._log(log, f"{'='*45}\n[{i}/{len(brands)}] {brand}\n{'='*45}")
                try:
                    filtered = filter_by_brand(self.a_order_df, brand)
                    if len(filtered) == 0:
                        self._log(log, "  -- 상품 없음, 건너뜁니다.")
                        continue
                    merged, unmatched = match_barcode_to_uid(filtered, self.a_matching_df)
                    self._log(log, f"  매칭: {len(merged)}건" + (f" (미매칭: {len(unmatched)})" if unmatched else ""))

                    # 미매칭 바코드 경고 수집
                    if unmatched:
                        for bc in unmatched:
                            all_warnings.append({
                                '브랜드': brand, '88코드': bc,
                                '원인': '바코드 매칭 실패 (파일2에 해당 바코드 없음)',
                            })

                    has_option = detect_option_products(merged, self.a_matching_df)
                    merged, warnings = match_option_info(merged, self.a_option_df, has_option, self.a_matching_df)
                    if warnings:
                        self._log(log, f"  -- 옵션 경고: {len(warnings)}건")
                        for w in warnings:
                            all_warnings.append({
                                '브랜드': brand,
                                '88코드': w.get('88코드', ''),
                                '상품명': w.get('상품명', ''),
                                '원인': w.get('원인', ''),
                            })

                    sp = Path(self.a_save_path) / f'{brand}_시스템업로드_최종파일.xlsx'
                    bp = Path(self.a_save_path) / f'{brand}_발주리스트_최종파일.xlsx'
                    generate_system_upload(merged, self.a_matching_df, self.a_option_df, has_option, brand, sp)
                    generate_brand_order(merged, has_option, brand, bp)
                    self._log(log, f"  >> {sp.name}\n  >> {bp.name}")
                    ok += 1
                except Exception as e:
                    fail += 1
                    # [D4] 사용자 친화적 에러 메시지
                    self._log(log, f"  !! 오류: {friendly_error(e)}")
                self._log(log, "")

            cancelled = self._cancel_event.is_set()
            status_text = "취소됨" if cancelled else "완료"
            self._log(log, f"{'='*45}\n{status_text}: 성공 {ok} / 실패 {fail} / 총 {len(brands)}건")
            if all_warnings:
                self._log(log, f"경고: {len(all_warnings)}건 (매칭 실패 리포트 버튼으로 상세 확인)")
            self._log(log, f"저장: {self.a_save_path}")

            # 경고 데이터 저장
            self.a_last_warnings = all_warnings

            # [P1] 이력 저장
            add_history_entry({
                'type': 'A',
                'brands': brands[:20],
                'success': ok,
                'fail': fail,
                'total': len(brands),
                'files': {
                    'file1': Path(self.a_file_paths[1]).name if self.a_file_paths[1] else '',
                    'file2': Path(self.a_file_paths[2]).name if self.a_file_paths[2] else '',
                    'file3': Path(self.a_file_paths[3]).name if self.a_file_paths[3] else '',
                },
                'output_path': self.a_save_path,
                'warnings': len(all_warnings),
            })

            self.root.after(0, lambda: self.a_open_btn.configure(state='normal'))
            if all_warnings:
                self.root.after(0, lambda: self.a_report_btn.configure(state='normal'))

        except Exception as e:
            # [D4] 사용자 친화적 에러 메시지
            self._log(log, f"\n!! 오류:\n{friendly_error(e)}")
        finally:
            self.root.after(0, self._a_run_finished)
            self._save_session()

    def _a_run_finished(self):
        self.a_run_btn.configure(state='normal')
        self.a_cancel_btn.configure(state='disabled')
        self.a_progress_label.configure(text="완료")
        self._refresh_history()

    # [D1] 프로그레스 업데이트
    def _a_update_progress(self, value, brand, total):
        self.a_progress['value'] = value
        self.a_progress_label.configure(text=f"{brand} ({value}/{total})")

    # [P4] 매칭 실패 리포트 내보내기
    def _a_export_warning_report(self):
        if not self.a_last_warnings:
            messagebox.showinfo("알림", "내보낼 경고가 없습니다.")
            return

        fp = filedialog.asksaveasfilename(
            title="매칭 실패 리포트 저장",
            defaultextension=".xlsx",
            initialfile=f"매칭실패_리포트_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            initialdir=self.a_save_path,
            filetypes=[("Excel", "*.xlsx")],
        )
        if not fp:
            return

        try:
            df = pd.DataFrame(self.a_last_warnings)
            df.to_excel(fp, index=False)
            messagebox.showinfo("완료", f"리포트 저장 완료:\n{fp}")
        except Exception as e:
            messagebox.showerror("오류", f"리포트 저장 실패:\n{friendly_error(e)}")

    # 리셋
    def _a_reset(self):
        if messagebox.askyesno("확인", "모든 입력을 초기화하시겠습니까?"):
            self.a_order_df = None
            self.a_matching_df = None
            self.a_option_df = None
            self.a_file_paths = {1: None, 2: None, 3: None}
            self.a_all_brands = []
            self.a_last_warnings = []

            for fnum in [1, 2, 3]:
                self._set_entry(self.a_file_entries[fnum], '')
                self.a_file_status[fnum].set_status('none', '')

            self.a_brand_mode.set('all')
            self.a_brand_vars = {}
            self.a_cb_widgets.clear()
            for w in self.a_cb_inner.winfo_children():
                w.destroy()
            self.a_selected_label.configure(text="")
            self.a_search_var.set('')
            self._a_on_brand_mode()

            self.a_progress['value'] = 0
            self.a_progress_label.configure(text="")
            self._clear_log(self.a_log)
            self.a_open_btn.configure(state='disabled')
            self.a_report_btn.configure(state='disabled')
            self._save_session()

    # [P3] 프리셋 관리
    def _a_refresh_presets(self):
        presets = get_presets()
        self.a_preset_combo['values'] = list(presets.keys())

    def _a_save_preset(self):
        name = self.a_preset_var.get().strip()
        if not name:
            # 새 이름 입력 다이얼로그
            dialog = tk.Toplevel(self.root)
            dialog.title("프리셋 저장")
            dialog.geometry("300x120")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()

            ttk.Label(dialog, text="프리셋 이름:").pack(padx=20, pady=(15, 5))
            name_var = tk.StringVar()
            name_entry = ttk.Entry(dialog, textvariable=name_var, width=30)
            name_entry.pack(padx=20)
            name_entry.focus()

            def do_save():
                n = name_var.get().strip()
                if n:
                    self._a_do_save_preset(n)
                    dialog.destroy()

            ttk.Button(dialog, text="저장", command=do_save).pack(pady=10)
            dialog.bind('<Return>', lambda e: do_save())
            return

        self._a_do_save_preset(name)

    def _a_do_save_preset(self, name):
        config = {
            'file1': self.a_file_paths.get(1),
            'file2': self.a_file_paths.get(2),
            'file3': self.a_file_paths.get(3),
            'save_path': self.a_save_path,
            'brand_mode': self.a_brand_mode.get(),
            'selected_brands': [b for b, v in self.a_brand_vars.items() if v.get()],
        }
        save_preset(name, config)
        self._a_refresh_presets()
        self.a_preset_var.set(name)
        messagebox.showinfo("완료", f"프리셋 '{name}' 저장 완료")

    def _a_load_preset(self):
        name = self.a_preset_var.get().strip()
        if not name:
            messagebox.showwarning("선택 필요", "불러올 프리셋을 선택해주세요.")
            return
        config = load_preset(name)
        if not config:
            messagebox.showwarning("오류", f"프리셋 '{name}'을 찾을 수 없습니다.")
            return

        # 파일 복원
        for fnum, key in [(1, 'file1'), (2, 'file2'), (3, 'file3')]:
            fp = config.get(key)
            if fp and Path(fp).exists():
                self._a_set_file(fnum, fp)
            elif fp:
                self._set_entry(self.a_file_entries[fnum], f"(파일 없음) {fp}")
                self.a_file_status[fnum].set_status('invalid', '파일이 존재하지 않습니다')

        # 저장 경로 복원
        sp = config.get('save_path', '')
        if sp and Path(sp).exists():
            self.a_save_path = sp
            self._set_entry(self.a_save_entry, sp)

        # 브랜드 모드 복원
        mode = config.get('brand_mode', 'all')
        self.a_brand_mode.set(mode)
        self._a_on_brand_mode()

        # 브랜드 선택 복원
        for b in config.get('selected_brands', []):
            if b in self.a_brand_vars:
                self.a_brand_vars[b].set(True)
        self._a_update_selected_label()

    def _a_delete_preset(self):
        name = self.a_preset_var.get().strip()
        if not name:
            return
        if messagebox.askyesno("확인", f"프리셋 '{name}'을 삭제하시겠습니까?"):
            delete_preset(name)
            self._a_refresh_presets()
            self.a_preset_var.set('')

    # ================================================================
    #  탭 B 로직
    # ================================================================

    # [D5] 드래그앤드롭 처리
    def _b_on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for fp in files:
            fp = fp.strip('{}')
            if fp and fp not in self.b_file_list:
                self.b_file_list.append(fp)
                self.b_file_listbox.insert(tk.END, Path(fp).name)
        self._save_session()

    def _b_update_listbox_tooltip(self, event):
        """리스트박스 항목 위에서 전체 경로를 툴팁으로 표시"""
        idx = self.b_file_listbox.nearest(event.y)
        if 0 <= idx < len(self.b_file_list):
            self._b_listbox_tooltip.update_text(self.b_file_list[idx])

    def _b_add_files(self):
        fps = filedialog.askopenfilenames(
            title="발주 파일 선택 (복수 가능)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        for fp in fps:
            if fp not in self.b_file_list:
                self.b_file_list.append(fp)
                self.b_file_listbox.insert(tk.END, Path(fp).name)
        self._save_session()

    def _b_remove_file(self):
        for idx in reversed(self.b_file_listbox.curselection()):
            self.b_file_listbox.delete(idx)
            del self.b_file_list[idx]
        self._save_session()

    def _b_select_matching(self):
        fp = filedialog.askopenfilename(
            title="매칭 파일 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        if not fp:
            return
        self.b_matching_path = fp
        self._set_entry(self.b_match_entry, fp)
        self._save_session()

    def _b_run(self):
        if not self.b_file_list:
            messagebox.showwarning("입력 필요", "발주 파일을 1개 이상 추가해주세요.")
            return

        # [P5] 덮어쓰기 확인
        output_path = Path(self.b_save_path) / '통합_발주수량.xlsx'
        if output_path.exists():
            if not messagebox.askyesno("덮어쓰기 확인",
                                       f"'{output_path.name}' 파일이 이미 존재합니다.\n덮어쓰시겠습니까?"):
                return

        self._cancel_event.clear()
        self.b_run_btn.configure(state='disabled')
        self.b_cancel_btn.configure(state='normal')
        self.b_open_btn.configure(state='disabled')
        self._clear_log(self.b_log)

        # [D1] 프로그레스 시작
        self.b_progress.start(10)
        self.b_progress_label.configure(text="처리 중...")

        self._log(self.b_log, f"통합 시작... ({len(self.b_file_list)}개 파일)\n")
        threading.Thread(target=self._b_do_run, daemon=True).start()

    # [E5] 취소
    def _b_cancel(self):
        self._cancel_event.set()
        self.b_cancel_btn.configure(state='disabled')
        self._log(self.b_log, "\n취소 요청됨...")

    def _b_do_run(self):
        log = self.b_log
        try:
            for fp in self.b_file_list:
                self._log(log, f"  파일: {Path(fp).name}")

            if self.b_matching_path:
                self._log(log, f"  매칭: {Path(self.b_matching_path).name}")

            output_path = Path(self.b_save_path) / '통합_발주수량.xlsx'

            self._log(log, "\n통합 처리 중...")
            df = merge_order_files(
                self.b_file_list,
                self.b_matching_path,
                output_path,
            )

            brands = df['브랜드명'].unique().tolist()
            matched = df['상품번호'].apply(lambda x: x != '' and x is not None).sum()

            self._log(log, f"\n{'='*45}")
            self._log(log, f">> 통합 완료!")
            self._log(log, f"  총 {len(df)}행 ({len(brands)}개 브랜드: {', '.join(brands)})")
            self._log(log, f"  상품번호 매칭: {matched}/{len(df)}건")
            self._log(log, f"  출력: {output_path.name}")
            self._log(log, f"  저장: {self.b_save_path}")

            # [P1] 이력 저장
            add_history_entry({
                'type': 'B',
                'brands': brands[:20],
                'success': 1,
                'fail': 0,
                'total': len(self.b_file_list),
                'output_path': self.b_save_path,
                'warnings': 0,
            })

            self.root.after(0, lambda: self.b_open_btn.configure(state='normal'))
        except Exception as e:
            # [D4] 사용자 친화적 에러 메시지
            self._log(log, f"\n!! 오류:\n{friendly_error(e)}")
        finally:
            self.root.after(0, self._b_run_finished)
            self._save_session()

    def _b_run_finished(self):
        self.b_run_btn.configure(state='normal')
        self.b_cancel_btn.configure(state='disabled')
        self.b_progress.stop()
        self.b_progress_label.configure(text="완료")
        self._refresh_history()

    # 리셋
    def _b_reset(self):
        if messagebox.askyesno("확인", "모든 입력을 초기화하시겠습니까?"):
            self.b_file_list.clear()
            self.b_file_listbox.delete(0, tk.END)
            self.b_matching_path = None
            self._set_entry(self.b_match_entry, '')
            self.b_progress.stop()
            self.b_progress_label.configure(text="")
            self._clear_log(self.b_log)
            self.b_open_btn.configure(state='disabled')
            self._save_session()


def main():
    root = _create_root()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()

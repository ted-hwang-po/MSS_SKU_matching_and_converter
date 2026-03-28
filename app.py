"""MSS SKU 매칭 및 발주 파일 변환기 — GUI 앱"""
import sys
import os
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core.loader import load_order_data, load_matching_data, load_option_data, get_brand_list, filter_by_brand
from core.matcher import match_barcode_to_uid, detect_option_products, match_option_info
from core.generator import generate_system_upload, generate_brand_order
from core.merger import merge_order_files


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MSS SKU 매칭 및 발주 파일 변환기")
        self.root.geometry("720x800")
        self.root.resizable(False, False)

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main_frame, text="MSS SKU 매칭 및 발주 파일 변환기", font=("", 15, "bold"))
        title.pack(pady=(0, 10))

        # ── 탭 ──
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._build_tab_a()
        self._build_tab_b()

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

        # 파일 입력
        file_frame = ttk.LabelFrame(tab, text="입력 파일", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        file_labels = [
            ("1. 총 발주 수량 데이터셋", 1),
            ("2. 상품코드-바코드 매칭 데이터셋", 2),
            ("3. 상품별 옵션 정보", 3),
        ]
        self.a_file_entries = {}
        for label_text, fnum in file_labels:
            row = ttk.Frame(file_frame)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=label_text, width=30, anchor=tk.W).pack(side=tk.LEFT)
            entry = ttk.Entry(row, state='readonly')
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
            self.a_file_entries[fnum] = entry
            ttk.Button(row, text="파일 선택",
                       command=lambda n=fnum: self._a_select_file(n)).pack(side=tk.RIGHT)

        # 브랜드 선택
        brand_frame = ttk.LabelFrame(tab, text="브랜드 선택", padding=8)
        brand_frame.pack(fill=tk.X, pady=(0, 8))

        radio_row = ttk.Frame(brand_frame)
        radio_row.pack(fill=tk.X, pady=(0, 3))
        self.a_brand_mode = tk.StringVar(value='all')
        ttk.Radiobutton(radio_row, text="모든 브랜드", variable=self.a_brand_mode,
                        value='all', command=self._a_on_brand_mode).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Radiobutton(radio_row, text="브랜드 선택", variable=self.a_brand_mode,
                        value='select', command=self._a_on_brand_mode).pack(side=tk.LEFT)

        search_row = ttk.Frame(brand_frame)
        search_row.pack(fill=tk.X, pady=(3, 3))
        ttk.Label(search_row, text="검색:", width=5, anchor=tk.W).pack(side=tk.LEFT)
        self.a_search_var = tk.StringVar()
        self.a_search_var.trace_add('write', lambda *_: self._a_filter_brands())
        self.a_search_entry = ttk.Entry(search_row, textvariable=self.a_search_var, state='disabled')
        self.a_search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        list_frame = ttk.Frame(brand_frame)
        list_frame.pack(fill=tk.X)
        self.a_brand_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, height=4,
                                          exportselection=False, font=("", 10), state='disabled')
        self.a_brand_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Scrollbar(list_frame, orient=tk.VERTICAL,
                      command=self.a_brand_listbox.yview).pack(side=tk.RIGHT, fill=tk.Y)

        # 저장 위치
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

        # 실행
        self.a_run_btn = ttk.Button(tab, text="▶  변환 실행", command=self._a_run)
        self.a_run_btn.pack(pady=10)

        # 로그
        log_frame = ttk.LabelFrame(tab, text="처리 결과", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.a_log = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
        self.a_log.pack(fill=tk.BOTH, expand=True)

        self.a_open_btn = ttk.Button(tab, text="출력 폴더 열기",
                                     command=lambda: self._open_folder(self.a_save_path), state='disabled')
        self.a_open_btn.pack(pady=(5, 0))

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

        btn_row = ttk.Frame(file_frame)
        btn_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_row, text="+ 파일 추가", command=self._b_add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_row, text="- 선택 제거", command=self._b_remove_file).pack(side=tk.LEFT)

        # 매칭 파일 (선택)
        match_frame = ttk.LabelFrame(tab, text="상품코드-바코드 매칭 파일 (선택)", padding=8)
        match_frame.pack(fill=tk.X, pady=(0, 8))
        match_row = ttk.Frame(match_frame)
        match_row.pack(fill=tk.X)
        self.b_match_entry = ttk.Entry(match_row, state='readonly')
        self.b_match_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(match_row, text="파일 선택", command=self._b_select_matching).pack(side=tk.RIGHT)

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

        # 실행
        self.b_run_btn = ttk.Button(tab, text="▶  통합 실행", command=self._b_run)
        self.b_run_btn.pack(pady=10)

        # 로그
        log_frame = ttk.LabelFrame(tab, text="처리 결과", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.b_log = scrolledtext.ScrolledText(log_frame, height=8, state='disabled', font=("Consolas", 9))
        self.b_log.pack(fill=tk.BOTH, expand=True)

        self.b_open_btn = ttk.Button(tab, text="출력 폴더 열기",
                                     command=lambda: self._open_folder(self.b_save_path), state='disabled')
        self.b_open_btn.pack(pady=(5, 0))

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

    def _open_folder(self, path):
        if not path:
            return
        if sys.platform == 'win32':
            os.startfile(path)
        elif sys.platform == 'darwin':
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}"')

    # ================================================================
    #  탭 A 로직
    # ================================================================
    def _a_select_file(self, fnum):
        fp = filedialog.askopenfilename(
            title=f"파일 {fnum} 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        if not fp:
            return
        self.a_file_paths[fnum] = fp
        entry = self.a_file_entries[fnum]
        entry.configure(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, fp)
        entry.configure(state='readonly')
        if fnum == 1:
            self._a_load_brands(fp)

    def _a_load_brands(self, filepath):
        try:
            self.a_order_df = load_order_data(filepath)
            self.a_all_brands = get_brand_list(self.a_order_df)
            self.a_brand_listbox.configure(state='normal')
            self.a_brand_listbox.delete(0, tk.END)
            for b in self.a_all_brands:
                self.a_brand_listbox.insert(tk.END, b)
            self._a_on_brand_mode()
            self._log(self.a_log, f"브랜드 로드 완료 ({len(self.a_all_brands)}개): {', '.join(self.a_all_brands)}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 1 로드 실패:\n{e}")

    def _a_on_brand_mode(self):
        if self.a_brand_mode.get() == 'all':
            self.a_brand_listbox.selection_clear(0, tk.END)
            self.a_brand_listbox.configure(state='disabled')
            self.a_search_var.set('')
            self.a_search_entry.configure(state='disabled')
        else:
            self.a_brand_listbox.configure(state='normal')
            self.a_search_entry.configure(state='normal')

    def _a_filter_brands(self):
        q = self.a_search_var.get().strip().lower()
        self.a_brand_listbox.configure(state='normal')
        self.a_brand_listbox.delete(0, tk.END)
        for b in self.a_all_brands:
            if not q or q in b.lower():
                self.a_brand_listbox.insert(tk.END, b)

    def _a_get_brands(self):
        if self.a_brand_mode.get() == 'all':
            return list(self.a_all_brands)
        return [self.a_brand_listbox.get(i) for i in self.a_brand_listbox.curselection()]

    def _a_run(self):
        for n in [1, 2, 3]:
            if not self.a_file_paths[n]:
                messagebox.showwarning("입력 필요", f"파일 {n}을 선택해주세요.")
                return
        brands = self._a_get_brands()
        if not brands:
            messagebox.showwarning("입력 필요", "브랜드를 1개 이상 선택해주세요.")
            return
        self.a_run_btn.configure(state='disabled')
        self.a_open_btn.configure(state='disabled')
        self._clear_log(self.a_log)
        self._log(self.a_log, f"변환 시작... ({len(brands)}개 브랜드)\n")
        threading.Thread(target=self._a_do_run, args=(brands,), daemon=True).start()

    def _a_do_run(self, brands):
        log = self.a_log
        try:
            self._log(log, "파일 로딩 중...")
            if self.a_order_df is None:
                self.a_order_df = load_order_data(self.a_file_paths[1])
            self.a_matching_df = load_matching_data(self.a_file_paths[2])
            self.a_option_df = load_option_data(self.a_file_paths[3])

            ok, fail = 0, 0
            for i, brand in enumerate(brands, 1):
                self._log(log, f"{'='*45}\n[{i}/{len(brands)}] {brand}\n{'='*45}")
                try:
                    filtered = filter_by_brand(self.a_order_df, brand)
                    if len(filtered) == 0:
                        self._log(log, "  ⚠️ 상품 없음, 건너뜁니다.")
                        continue
                    merged, unmatched = match_barcode_to_uid(filtered, self.a_matching_df)
                    self._log(log, f"  매칭: {len(merged)}건" + (f" (실패: {len(unmatched)})" if unmatched else ""))
                    has_option = detect_option_products(merged, self.a_matching_df)
                    merged, warnings = match_option_info(merged, self.a_option_df, has_option, self.a_matching_df)
                    if warnings:
                        self._log(log, f"  ⚠️ 옵션 경고: {len(warnings)}건")

                    sp = Path(self.a_save_path) / f'{brand}_시스템업로드_최종파일.xlsx'
                    bp = Path(self.a_save_path) / f'{brand}_발주리스트_최종파일.xlsx'
                    generate_system_upload(merged, self.a_matching_df, self.a_option_df, has_option, brand, sp)
                    generate_brand_order(merged, has_option, brand, bp)
                    self._log(log, f"  ✅ {sp.name}\n  ✅ {bp.name}")
                    ok += 1
                except Exception as e:
                    fail += 1
                    self._log(log, f"  ❌ {e}\n{traceback.format_exc()}")
                self._log(log, "")

            self._log(log, f"{'='*45}\n완료: 성공 {ok} / 실패 {fail} / 총 {len(brands)}건\n저장: {self.a_save_path}")
            self.root.after(0, lambda: self.a_open_btn.configure(state='normal'))
        except Exception as e:
            self._log(log, f"\n❌ {traceback.format_exc()}")
        finally:
            self.root.after(0, lambda: self.a_run_btn.configure(state='normal'))

    # ================================================================
    #  탭 B 로직
    # ================================================================
    def _b_add_files(self):
        fps = filedialog.askopenfilenames(
            title="발주 파일 선택 (복수 가능)",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        for fp in fps:
            if fp not in self.b_file_list:
                self.b_file_list.append(fp)
                self.b_file_listbox.insert(tk.END, Path(fp).name)

    def _b_remove_file(self):
        for idx in reversed(self.b_file_listbox.curselection()):
            self.b_file_listbox.delete(idx)
            del self.b_file_list[idx]

    def _b_select_matching(self):
        fp = filedialog.askopenfilename(
            title="매칭 파일 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All", "*.*")])
        if not fp:
            return
        self.b_matching_path = fp
        self.b_match_entry.configure(state='normal')
        self.b_match_entry.delete(0, tk.END)
        self.b_match_entry.insert(0, fp)
        self.b_match_entry.configure(state='readonly')

    def _b_run(self):
        if not self.b_file_list:
            messagebox.showwarning("입력 필요", "발주 파일을 1개 이상 추가해주세요.")
            return
        self.b_run_btn.configure(state='disabled')
        self.b_open_btn.configure(state='disabled')
        self._clear_log(self.b_log)
        self._log(self.b_log, f"통합 시작... ({len(self.b_file_list)}개 파일)\n")
        threading.Thread(target=self._b_do_run, daemon=True).start()

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
            self._log(log, f"✅ 통합 완료!")
            self._log(log, f"  총 {len(df)}행 ({len(brands)}개 브랜드: {', '.join(brands)})")
            self._log(log, f"  상품번호 매칭: {matched}/{len(df)}건")
            self._log(log, f"  출력: {output_path.name}")
            self._log(log, f"  저장: {self.b_save_path}")

            self.root.after(0, lambda: self.b_open_btn.configure(state='normal'))
        except Exception as e:
            self._log(log, f"\n❌ 오류:\n{traceback.format_exc()}")
        finally:
            self.root.after(0, lambda: self.b_run_btn.configure(state='normal'))


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == '__main__':
    main()

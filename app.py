"""MSS SKU 매칭 및 발주 파일 변환기 — GUI 앱"""
import sys
import os
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext


# 핵심 로직 임포트
from core.loader import load_order_data, load_matching_data, load_option_data, get_brand_list, filter_by_brand
from core.matcher import match_barcode_to_uid, detect_option_products, match_option_info
from core.generator import generate_system_upload, generate_brand_order


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("MSS SKU 매칭 및 발주 파일 변환기")
        self.root.geometry("700x780")
        self.root.resizable(False, False)

        # 데이터 상태
        self.order_df = None
        self.matching_df = None
        self.option_df = None
        self.file_paths = {1: None, 2: None, 3: None}
        self.all_brands = []

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 제목
        title = ttk.Label(main_frame, text="MSS SKU 매칭 및 발주 파일 변환기", font=("", 16, "bold"))
        title.pack(pady=(0, 15))

        # ── 파일 입력 영역 ──
        file_frame = ttk.LabelFrame(main_frame, text="입력 파일", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))

        file_labels = [
            ("1. 총 발주 수량 데이터셋", 1),
            ("2. 상품코드-바코드 매칭 데이터셋", 2),
            ("3. 상품별 옵션 정보", 3),
        ]

        self.file_entries = {}
        for label_text, file_num in file_labels:
            row = ttk.Frame(file_frame)
            row.pack(fill=tk.X, pady=3)

            lbl = ttk.Label(row, text=label_text, width=30, anchor=tk.W)
            lbl.pack(side=tk.LEFT)

            entry = ttk.Entry(row, state='readonly')
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
            self.file_entries[file_num] = entry

            btn = ttk.Button(row, text="파일 선택",
                             command=lambda n=file_num: self._select_file(n))
            btn.pack(side=tk.RIGHT)

        # ── 브랜드 선택 ──
        brand_frame = ttk.LabelFrame(main_frame, text="브랜드 선택", padding=10)
        brand_frame.pack(fill=tk.X, pady=(0, 10))

        # 라디오 버튼: 전체 / 선택
        radio_row = ttk.Frame(brand_frame)
        radio_row.pack(fill=tk.X, pady=(0, 5))

        self.brand_mode = tk.StringVar(value='all')
        ttk.Radiobutton(radio_row, text="모든 브랜드", variable=self.brand_mode,
                        value='all', command=self._on_brand_mode_change).pack(side=tk.LEFT, padx=(0, 15))
        ttk.Radiobutton(radio_row, text="브랜드 선택", variable=self.brand_mode,
                        value='select', command=self._on_brand_mode_change).pack(side=tk.LEFT)

        # 검색창
        search_row = ttk.Frame(brand_frame)
        search_row.pack(fill=tk.X, pady=(5, 3))

        ttk.Label(search_row, text="검색:", width=5, anchor=tk.W).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', lambda *_: self._filter_brand_list())
        self.search_entry = ttk.Entry(search_row, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 브랜드 리스트박스 (복수 선택)
        list_frame = ttk.Frame(brand_frame)
        list_frame.pack(fill=tk.X, pady=(3, 0))

        self.brand_listbox = tk.Listbox(
            list_frame, selectmode=tk.EXTENDED, height=5,
            exportselection=False, font=("", 10),
        )
        self.brand_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.brand_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.brand_listbox.configure(yscrollcommand=scrollbar.set)

        # 초기 상태: "모든 브랜드" → 리스트/검색 비활성
        self.brand_listbox.configure(state='disabled')
        self.search_entry.configure(state='disabled')

        # ── 저장 위치 ──
        save_frame = ttk.LabelFrame(main_frame, text="저장 위치", padding=10)
        save_frame.pack(fill=tk.X, pady=(0, 10))

        save_row = ttk.Frame(save_frame)
        save_row.pack(fill=tk.X)

        self.save_entry = ttk.Entry(save_row, state='readonly')
        self.save_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        default_save = str(Path.home() / "Desktop")
        self.save_path = default_save
        self.save_entry.configure(state='normal')
        self.save_entry.insert(0, default_save)
        self.save_entry.configure(state='readonly')

        ttk.Button(save_row, text="폴더 선택",
                   command=self._select_save_folder).pack(side=tk.RIGHT)

        # ── 실행 버튼 ──
        self.run_btn = ttk.Button(main_frame, text="▶  변환 실행",
                                  command=self._run_conversion)
        self.run_btn.pack(pady=15)

        # ── 결과 로그 ──
        log_frame = ttk.LabelFrame(main_frame, text="처리 결과", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, state='disabled',
                                                   font=("Consolas", 10))
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # ── 출력 폴더 열기 ──
        self.open_folder_btn = ttk.Button(main_frame, text="출력 폴더 열기",
                                          command=self._open_output_folder, state='disabled')
        self.open_folder_btn.pack(pady=(10, 0))

    def _on_brand_mode_change(self):
        if self.brand_mode.get() == 'all':
            self.brand_listbox.selection_clear(0, tk.END)
            self.brand_listbox.configure(state='disabled')
            self.search_var.set('')
            self.search_entry.configure(state='disabled')
        else:
            self.brand_listbox.configure(state='normal')
            self.search_entry.configure(state='normal')

    def _filter_brand_list(self):
        query = self.search_var.get().strip().lower()
        self.brand_listbox.configure(state='normal')
        self.brand_listbox.delete(0, tk.END)
        for b in self.all_brands:
            if not query or query in b.lower():
                self.brand_listbox.insert(tk.END, b)

    def _select_file(self, file_num: int):
        filepath = filedialog.askopenfilename(
            title=f"파일 {file_num} 선택",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if not filepath:
            return

        self.file_paths[file_num] = filepath

        entry = self.file_entries[file_num]
        entry.configure(state='normal')
        entry.delete(0, tk.END)
        entry.insert(0, filepath)
        entry.configure(state='readonly')

        if file_num == 1:
            self._load_brands(filepath)

    def _load_brands(self, filepath: str):
        try:
            self.order_df = load_order_data(filepath)
            self.all_brands = get_brand_list(self.order_df)

            self.brand_listbox.configure(state='normal')
            self.brand_listbox.delete(0, tk.END)
            for b in self.all_brands:
                self.brand_listbox.insert(tk.END, b)

            # 모드에 따라 상태 설정
            self._on_brand_mode_change()

            self._log(f"브랜드 목록 로드 완료 ({len(self.all_brands)}개): {', '.join(self.all_brands)}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 1 로드 실패:\n{e}")

    def _get_selected_brands(self) -> list[str]:
        if self.brand_mode.get() == 'all':
            return list(self.all_brands)
        else:
            indices = self.brand_listbox.curselection()
            return [self.brand_listbox.get(i) for i in indices]

    def _select_save_folder(self):
        folder = filedialog.askdirectory(title="저장 폴더 선택")
        if folder:
            self.save_path = folder
            self.save_entry.configure(state='normal')
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, folder)
            self.save_entry.configure(state='readonly')

    def _log(self, msg: str):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')

    def _clear_log(self):
        self.log_text.configure(state='normal')
        self.log_text.delete('1.0', tk.END)
        self.log_text.configure(state='disabled')

    def _run_conversion(self):
        for n in [1, 2, 3]:
            if not self.file_paths[n]:
                messagebox.showwarning("입력 필요", f"파일 {n}을 선택해주세요.")
                return

        brands = self._get_selected_brands()
        if not brands:
            messagebox.showwarning("입력 필요", "브랜드를 1개 이상 선택해주세요.")
            return

        if not self.save_path:
            messagebox.showwarning("입력 필요", "저장 위치를 선택해주세요.")
            return

        self.run_btn.configure(state='disabled')
        self.open_folder_btn.configure(state='disabled')
        self._clear_log()

        mode_text = "모든 브랜드" if self.brand_mode.get() == 'all' else f"선택 브랜드 {len(brands)}개"
        self._log(f"변환을 시작합니다... ({mode_text})\n")

        thread = threading.Thread(target=self._do_conversion_multi, args=(brands,), daemon=True)
        thread.start()

    def _do_conversion_multi(self, brands: list[str]):
        try:
            # 파일 로딩 (1회)
            self._log("파일 로딩 중...")
            if self.order_df is None:
                self.order_df = load_order_data(self.file_paths[1])
            self.matching_df = load_matching_data(self.file_paths[2])
            self.option_df = load_option_data(self.file_paths[3])
            self._log(f"  파일1: {len(self.order_df)}행, 파일2: {len(self.matching_df)}행, 파일3: {len(self.option_df)}행\n")

            total_brands = len(brands)
            success_count = 0
            fail_count = 0

            for i, brand in enumerate(brands, 1):
                self._log(f"{'='*50}")
                self._log(f"[{i}/{total_brands}] {brand}")
                self._log(f"{'='*50}")

                try:
                    self._do_conversion_single(brand)
                    success_count += 1
                except Exception as e:
                    fail_count += 1
                    self._log(f"  ❌ 오류: {e}")
                    self._log(traceback.format_exc())

                self._log("")

            # 최종 요약
            self._log(f"{'='*50}")
            self._log(f"전체 완료: 성공 {success_count}건 / 실패 {fail_count}건 / 총 {total_brands}건")
            self._log(f"저장 위치: {self.save_path}")

            self.root.after(0, lambda: self.open_folder_btn.configure(state='normal'))

        except Exception as e:
            self._log(f"\n❌ 오류 발생:\n{traceback.format_exc()}")
            self.root.after(0, lambda: messagebox.showerror("오류", str(e)))

        finally:
            self.root.after(0, lambda: self.run_btn.configure(state='normal'))

    def _do_conversion_single(self, brand: str):
        # 브랜드 필터링
        filtered = filter_by_brand(self.order_df, brand)
        self._log(f"  필터링: {len(filtered)}건")

        if len(filtered) == 0:
            self._log(f"  ⚠️ 해당 브랜드의 상품이 없습니다. 건너뜁니다.")
            return

        # 바코드-UID 매칭
        merged, unmatched = match_barcode_to_uid(filtered, self.matching_df)
        self._log(f"  바코드-UID 매칭: {len(merged)}건")
        if unmatched:
            self._log(f"  ⚠️ 매칭 실패 바코드: {unmatched}")

        # 옵션 판별
        has_option = detect_option_products(merged, self.matching_df)
        opt_count = sum(1 for v in has_option.values() if v)
        self._log(f"  옵션 판별: 옵션 {opt_count}건, 단일 {len(has_option) - opt_count}건")

        # 옵션 매핑
        merged, warnings = match_option_info(merged, self.option_df, has_option, self.matching_df)
        if warnings:
            self._log(f"  ⚠️ 옵션 매핑 경고: {len(warnings)}건")
            for w in warnings:
                self._log(f"    - {w['상품명']}: {w['원인']}")

        # 파일 생성
        sys_path = Path(self.save_path) / f'{brand}_시스템업로드_최종파일.xlsx'
        brand_path = Path(self.save_path) / f'{brand}_발주리스트_최종파일.xlsx'

        generate_system_upload(merged, self.matching_df, self.option_df, has_option, brand, sys_path)
        generate_brand_order(merged, has_option, brand, brand_path)

        self._log(f"  ✅ {sys_path.name}")
        self._log(f"  ✅ {brand_path.name}")

    def _open_output_folder(self):
        if not self.save_path:
            return
        if sys.platform == 'win32':
            os.startfile(self.save_path)
        elif sys.platform == 'darwin':
            os.system(f'open "{self.save_path}"')
        else:
            os.system(f'xdg-open "{self.save_path}"')


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == '__main__':
    main()

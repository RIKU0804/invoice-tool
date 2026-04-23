"""
支払通知書自動抽出ツール - GUIエントリーポイント

実行:
    python gui.py
"""

import sys
import os
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("支払通知書 自動抽出ツール")
        self.resizable(False, False)
        self._build_ui()
        self._center_window(520, 460)

    def _center_window(self, w: int, h: int):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        pad = {"padx": 16, "pady": 6}

        # --- PDF選択 ---
        frm_pdf = ttk.LabelFrame(self, text="① 入力PDF")
        frm_pdf.pack(fill="x", **pad)

        self.pdf_var = tk.StringVar()
        ttk.Entry(frm_pdf, textvariable=self.pdf_var, width=48).pack(
            side="left", padx=(8, 4), pady=8
        )
        ttk.Button(frm_pdf, text="参照…", command=self._browse_pdf).pack(
            side="left", padx=(0, 8), pady=8
        )

        # --- テンプレートExcel選択 ---
        frm_tpl = ttk.LabelFrame(self, text="② テンプレートExcel（集計用.xlsx）")
        frm_tpl.pack(fill="x", **pad)

        self.tpl_var = tk.StringVar()
        ttk.Entry(frm_tpl, textvariable=self.tpl_var, width=48).pack(
            side="left", padx=(8, 4), pady=8
        )
        ttk.Button(frm_tpl, text="参照…", command=self._browse_tpl).pack(
            side="left", padx=(0, 8), pady=8
        )

        # --- シート名 / 振込金額 ---
        frm_info = ttk.LabelFrame(self, text="③ 処理設定")
        frm_info.pack(fill="x", **pad)

        inner = ttk.Frame(frm_info)
        inner.pack(padx=8, pady=8, fill="x")

        ttk.Label(inner, text="新規シート名:").grid(row=0, column=0, sticky="w")
        self.sheet_var = tk.StringVar(value="2025年1月")
        ttk.Entry(inner, textvariable=self.sheet_var, width=16).grid(
            row=0, column=1, padx=(4, 24), sticky="w"
        )

        ttk.Label(inner, text="振込金額（税込・円）:").grid(row=0, column=2, sticky="w")
        self.furikomi_var = tk.StringVar()
        ttk.Entry(inner, textvariable=self.furikomi_var, width=16).grid(
            row=0, column=3, padx=(4, 0), sticky="w"
        )
        ttk.Label(inner, text="※空欄でもOK（後でExcel入力）", foreground="gray").grid(
            row=1, column=2, columnspan=2, sticky="w", pady=(2, 0)
        )

        # --- 処理開始ボタン ---
        self.run_btn = ttk.Button(
            self, text="▶  処理開始", command=self._start, style="Accent.TButton"
        )
        self.run_btn.pack(pady=(4, 2))

        # --- 進捗バー ---
        self.progress = ttk.Progressbar(self, mode="indeterminate", length=460)
        self.progress.pack(padx=16, pady=(0, 4))

        # --- ログ ---
        frm_log = ttk.LabelFrame(self, text="ログ")
        frm_log.pack(fill="both", expand=True, padx=16, pady=(0, 12))

        self.log_text = tk.Text(
            frm_log, height=9, state="disabled", wrap="word",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
            relief="flat", bd=0
        )
        sb = ttk.Scrollbar(frm_log, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True, padx=4, pady=4)

    # --------------------------------------------------------------- ファイル選択
    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="支払通知書PDFを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)

    def _browse_tpl(self):
        path = filedialog.askopenfilename(
            title="テンプレートExcelを選択",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.tpl_var.set(path)

    # --------------------------------------------------------------- 処理実行
    def _start(self):
        pdf_path = self.pdf_var.get().strip()
        tpl_path = self.tpl_var.get().strip()
        sheet = self.sheet_var.get().strip()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("エラー", "PDFファイルを選択してください")
            return
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("エラー", "テンプレートExcelを選択してください")
            return
        if not sheet:
            messagebox.showerror("エラー", "シート名を入力してください")
            return

        furikomi_raw = self.furikomi_var.get().strip().replace(",", "")
        furikomi = int(furikomi_raw) if furikomi_raw.isdigit() else None

        self.run_btn.config(state="disabled")
        self.progress.start(12)
        self._log_clear()
        self._log("処理を開始します...\n")

        threading.Thread(
            target=self._run_extraction,
            args=(pdf_path, tpl_path, sheet, furikomi),
            daemon=True
        ).start()

    def _run_extraction(self, pdf_path, tpl_path, sheet, furikomi):
        try:
            from config import CONFIG
            from pdf_converter import pdf_to_jpegs
            from orchestrator import run_parallel_extraction, select_final_result
            from excel_writer import classify_and_aggregate, write_to_template

            # 出力パス（テンプレートと同フォルダに保存）
            out_dir = Path(tpl_path).parent
            out_path = out_dir / f"集計用_{sheet}_自動反映.xlsx"

            self._log(f"[Step 1] PDF → JPEG 変換 (DPI={CONFIG['image_dpi']})")
            image_paths = pdf_to_jpegs(pdf_path, CONFIG["image_temp_dir"], CONFIG["image_dpi"], CONFIG["jpeg_quality"])

            self._log(f"[Step 2] pdfplumber & AI 並列抽出")
            parallel_results = run_parallel_extraction(
                pdf_path=pdf_path,
                image_paths=image_paths,
                models=CONFIG["models"],
                api_key=CONFIG["openrouter_api_key"],
                base_url=CONFIG["openrouter_base_url"],
                run_plumber_parallel=CONFIG["run_pdfplumber_in_parallel"],
                max_workers=CONFIG["max_workers"],
            )
            self._log(f"  成功ソース数: {parallel_results['n_successful_sources']}")

            self._log(f"[Step 3] 結果採用")
            final = select_final_result(parallel_results, tolerance=CONFIG["discrepancy_tolerance"], strategy=CONFIG["voting_strategy"])
            self._log(f"  採用ソース: {final['adopted_source']}")
            self._log(f"  明細行数: {len(final['rows'])}")
            if final["discrepancies"]:
                self._log(f"  ⚠ 不一致: {len(final['discrepancies'])}件")

            if not final["rows"]:
                raise RuntimeError("データ抽出に失敗しました")

            self._log(f"[Step 4] 邸別集計 & 分類ルール適用")
            aggregated = classify_and_aggregate(final["rows"])
            self._log(f"  集計後の邸数: {len(aggregated)}")

            self._log(f"[Step 5] Excel に書き込み")
            write_to_template(
                template_path=tpl_path,
                output_path=str(out_path),
                sheet_name=sheet,
                aggregated=aggregated,
                furikomi_kingaku=furikomi,
                pdf_koujidai_zeikomi=None,
                pdf_sousai_zeikomi=None,
            )

            self._log(f"\n✅ 完了: {out_path}")
            self.after(0, lambda: self._on_success(str(out_path)))

        except Exception as e:
            self._log(f"\n❌ エラー: {e}")
            self.after(0, self._on_error)

    def _on_success(self, out_path: str):
        self.progress.stop()
        self.run_btn.config(state="normal")
        if messagebox.askyesno("完了", f"処理が完了しました。\n\nExcelを開きますか？\n{out_path}"):
            os.startfile(out_path)

    def _on_error(self):
        self.progress.stop()
        self.run_btn.config(state="normal")
        messagebox.showerror("エラー", "処理に失敗しました。ログを確認してください。")

    # --------------------------------------------------------------- ログ
    def _log(self, msg: str):
        def _write():
            self.log_text.config(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        self.after(0, _write)

    def _log_clear(self):
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.config(state="disabled")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

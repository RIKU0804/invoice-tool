"""
支払通知書自動抽出ツール - GUIエントリーポイント（customtkinter版）

実行:
    python gui.py
"""

import sys
import os
import threading
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("支払通知書 自動抽出ツール")
        self.geometry("560x540")
        self.resizable(False, False)
        self._build_ui()
        self._center_window()

    def _center_window(self):
        self.update_idletasks()
        w, h = 560, 540
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # タイトル
        ctk.CTkLabel(
            self, text="支払通知書 自動抽出ツール",
            font=ctk.CTkFont(size=18, weight="bold")
        ).grid(row=0, column=0, pady=(24, 4))

        ctk.CTkLabel(
            self, text="PDFを選択して処理開始ボタンを押してください",
            font=ctk.CTkFont(size=12), text_color="gray"
        ).grid(row=1, column=0, pady=(0, 16))

        # --- PDF選択 ---
        frm_pdf = ctk.CTkFrame(self)
        frm_pdf.grid(row=2, column=0, padx=24, pady=6, sticky="ew")
        frm_pdf.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frm_pdf, text="入力PDF", width=110, anchor="w").grid(
            row=0, column=0, padx=(16, 8), pady=12
        )
        self.pdf_var = ctk.StringVar()
        ctk.CTkEntry(frm_pdf, textvariable=self.pdf_var, placeholder_text="PDFファイルを選択…").grid(
            row=0, column=1, padx=(0, 8), pady=12, sticky="ew"
        )
        ctk.CTkButton(frm_pdf, text="参照", width=60, command=self._browse_pdf).grid(
            row=0, column=2, padx=(0, 12), pady=12
        )

        # --- テンプレートExcel選択 ---
        frm_tpl = ctk.CTkFrame(self)
        frm_tpl.grid(row=3, column=0, padx=24, pady=6, sticky="ew")
        frm_tpl.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(frm_tpl, text="テンプレートExcel", width=110, anchor="w").grid(
            row=0, column=0, padx=(16, 8), pady=12
        )
        self.tpl_var = ctk.StringVar()
        ctk.CTkEntry(frm_tpl, textvariable=self.tpl_var, placeholder_text="集計用.xlsx を選択…").grid(
            row=0, column=1, padx=(0, 8), pady=12, sticky="ew"
        )
        ctk.CTkButton(frm_tpl, text="参照", width=60, command=self._browse_tpl).grid(
            row=0, column=2, padx=(0, 12), pady=12
        )

        # --- シート名 / 振込金額 ---
        frm_info = ctk.CTkFrame(self)
        frm_info.grid(row=4, column=0, padx=24, pady=6, sticky="ew")
        frm_info.grid_columnconfigure(1, weight=1)
        frm_info.grid_columnconfigure(3, weight=1)

        ctk.CTkLabel(frm_info, text="新規シート名", width=110, anchor="w").grid(
            row=0, column=0, padx=(16, 8), pady=12
        )
        self.sheet_var = ctk.StringVar(value="2025年1月")
        ctk.CTkEntry(frm_info, textvariable=self.sheet_var, width=120).grid(
            row=0, column=1, padx=(0, 16), pady=12, sticky="w"
        )

        ctk.CTkLabel(frm_info, text="振込金額（税込）", width=110, anchor="w").grid(
            row=0, column=2, padx=(0, 8), pady=12
        )
        self.furikomi_var = ctk.StringVar()
        ctk.CTkEntry(
            frm_info, textvariable=self.furikomi_var,
            placeholder_text="空欄でもOK", width=140
        ).grid(row=0, column=3, padx=(0, 16), pady=12, sticky="w")

        # --- 処理開始ボタン ---
        self.run_btn = ctk.CTkButton(
            self, text="▶  処理開始",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=10,
            command=self._start
        )
        self.run_btn.grid(row=5, column=0, padx=24, pady=(12, 6), sticky="ew")

        # --- 進捗バー ---
        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.grid(row=6, column=0, padx=24, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        # --- ログ ---
        self.log_box = ctk.CTkTextbox(
            self, height=180, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", wrap="word"
        )
        self.log_box.grid(row=7, column=0, padx=24, pady=(0, 20), sticky="ew")

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

        self.run_btn.configure(state="disabled")
        self.progress.start()
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

            out_dir = Path(tpl_path).parent
            out_path = out_dir / f"集計用_{sheet}_自動反映.xlsx"

            self._log(f"[Step 1] PDF → JPEG 変換 (DPI={CONFIG['image_dpi']})")
            image_paths = pdf_to_jpegs(
                pdf_path, CONFIG["image_temp_dir"], CONFIG["image_dpi"], CONFIG["jpeg_quality"]
            )

            self._log("[Step 2] pdfplumber & AI 並列抽出")
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

            self._log("[Step 3] 結果採用")
            final = select_final_result(
                parallel_results,
                tolerance=CONFIG["discrepancy_tolerance"],
                strategy=CONFIG["voting_strategy"]
            )
            self._log(f"  採用ソース: {final['adopted_source']}")
            self._log(f"  明細行数: {len(final['rows'])}")
            if final["discrepancies"]:
                self._log(f"  ⚠ 不一致: {len(final['discrepancies'])}件")

            if not final["rows"]:
                raise RuntimeError("データ抽出に失敗しました")

            self._log("[Step 4] 邸別集計 & 分類ルール適用")
            aggregated = classify_and_aggregate(final["rows"])
            self._log(f"  集計後の邸数: {len(aggregated)}")

            self._log("[Step 5] Excel に書き込み")
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
        self.progress.set(1)
        self.run_btn.configure(state="normal")
        if messagebox.askyesno("完了", f"処理が完了しました！\n\nExcelを開きますか？\n{out_path}"):
            os.startfile(out_path)

    def _on_error(self):
        self.progress.stop()
        self.progress.set(0)
        self.run_btn.configure(state="normal")
        messagebox.showerror("エラー", "処理に失敗しました。ログを確認してください。")

    # --------------------------------------------------------------- ログ
    def _log(self, msg: str):
        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _write)

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

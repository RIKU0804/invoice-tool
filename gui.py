"""
支払通知書自動抽出ツール - GUIエントリーポイント（customtkinter版）

実行:
    python gui.py
"""

import sys
import os
import json
import datetime
import tempfile
import threading
import webbrowser
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

from version import VERSION


def _settings_path() -> Path:
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home())) / "invoice-tool"
    else:
        base = Path.home() / ".invoice-tool"
    base.mkdir(parents=True, exist_ok=True)
    return base / "settings.json"


def _load_settings() -> dict:
    try:
        p = _settings_path()
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_settings(data: dict) -> None:
    try:
        _settings_path().write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    except Exception:
        pass


def _bundled_template() -> str:
    """PyInstaller同梱テンプレートのパスを返す（開発時はローカルパス）"""
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "template", "集計用.xlsx")
    return os.path.join(os.path.dirname(__file__), "template", "集計用.xlsx")

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.settings = _load_settings()
        self.title(f"PDF 明細抽出  v{VERSION}")
        self.geometry("560x720")
        self.resizable(False, False)
        self._build_ui()
        self._center_window()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(1000, self._check_update)

    def _on_close(self):
        self._persist_settings()
        self.destroy()

    def _persist_settings(self):
        _save_settings({
            "out_dir": self.out_dir_var.get().strip(),
        })

    def _check_update(self):
        threading.Thread(target=self._run_update_check, daemon=True).start()

    def _run_update_check(self, manual: bool = False):
        from updater import set_update_callback, run_update_check

        def _on_update(kind: str, payload):
            def _show():
                if kind == "new":
                    self._show_update_dialog(payload["version"], payload["url"])
                elif kind == "current" and manual:
                    messagebox.showinfo("最新です", str(payload))
                elif kind == "error" and manual:
                    messagebox.showerror("更新確認エラー", str(payload))
            self.after(0, _show)

        try:
            set_update_callback(_on_update)
            run_update_check(silent_if_current=not manual, force=manual)
        except Exception as e:
            if manual:
                self.after(0, lambda: messagebox.showerror("更新確認エラー", str(e)))

    def _show_update_dialog(self, version: str, html_url: str):
        """3択のアップデートダイアログを表示"""
        dlg = ctk.CTkToplevel(self)
        dlg.title("アップデートのお知らせ")
        dlg.geometry("440x220")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        ctk.CTkLabel(
            dlg, text=f"新しいバージョン {version} が利用可能です",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(pady=(20, 6))
        ctk.CTkLabel(
            dlg, text=f"現在のバージョン: v{VERSION}", text_color="gray"
        ).pack(pady=(0, 16))

        progress = ctk.CTkProgressBar(dlg, mode="determinate", width=360)
        progress.set(0)
        progress.pack(pady=(0, 6))
        status = ctk.CTkLabel(dlg, text="", font=ctk.CTkFont(size=11))
        status.pack(pady=(0, 8))

        btn_frm = ctk.CTkFrame(dlg, fg_color="transparent")
        btn_frm.pack(pady=(4, 12))

        def _do_in_app_update():
            btn_now.configure(state="disabled")
            btn_browser.configure(state="disabled")
            btn_later.configure(state="disabled")
            status.configure(text="ダウンロード中...")
            threading.Thread(
                target=self._run_in_app_update,
                args=(dlg, progress, status),
                daemon=True,
            ).start()

        def _open_browser():
            webbrowser.open(html_url)
            dlg.destroy()

        btn_now = ctk.CTkButton(btn_frm, text="今すぐ更新", width=110, command=_do_in_app_update)
        btn_now.grid(row=0, column=0, padx=6)
        btn_browser = ctk.CTkButton(btn_frm, text="ブラウザで開く", width=110, command=_open_browser)
        btn_browser.grid(row=0, column=1, padx=6)
        btn_later = ctk.CTkButton(
            btn_frm, text="後で", width=80, fg_color="gray", command=dlg.destroy
        )
        btn_later.grid(row=0, column=2, padx=6)

        # frozenでない環境では「今すぐ更新」無効化
        if not getattr(sys, "frozen", False):
            btn_now.configure(state="disabled")
            status.configure(text="(開発環境では自動更新不可)")

    def _run_in_app_update(self, dlg, progress, status):
        """バックグラウンドでexeをDL→差し替えbat起動→終了"""
        from updater import get_latest_exe_asset, download_exe, perform_self_update_swap

        def _update_progress(dl: int, total: int):
            if total <= 0:
                return
            ratio = dl / total
            text = f"ダウンロード中... {dl // 1024}KB / {total // 1024}KB"
            def _apply():
                progress.set(ratio)
                status.configure(text=text)
            self.after(0, _apply)

        try:
            asset = get_latest_exe_asset()
            if not asset or not asset.get("url"):
                raise RuntimeError("最新exeのURLが取得できませんでした")

            tmp_dir = Path(tempfile.gettempdir())
            new_exe = tmp_dir / f"invoice-tool-new-{asset['version']}.exe"

            ok = download_exe(asset["url"], new_exe, progress_cb=_update_progress)
            if not ok:
                raise RuntimeError("ダウンロードに失敗しました")

            self.after(0, lambda: status.configure(text="再起動して更新を適用します..."))
            self.after(300, lambda: perform_self_update_swap(new_exe))
        except Exception as e:
            err_msg = str(e)
            self.after(0, lambda: status.configure(text=f"失敗: {err_msg}"))
            self.after(0, lambda: messagebox.showerror(
                "更新に失敗しました",
                f"{err_msg}\n\n手動でGitHubからダウンロードしてください。",
                parent=dlg,
            ))

    def _center_window(self):
        self.update_idletasks()
        w, h = 560, 720
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)

        # タイトル
        ctk.CTkLabel(
            self, text="PDF 明細抽出",
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

        ctk.CTkLabel(frm_pdf, text="出力先フォルダ", width=110, anchor="w").grid(
            row=1, column=0, padx=(16, 8), pady=(0, 12)
        )
        saved_out = self.settings.get("out_dir", "").strip()
        if saved_out and Path(saved_out).is_dir():
            default_out = saved_out
        else:
            docs = Path.home() / "Documents"
            default_out = str(docs if docs.is_dir() else Path.home())
        self.out_dir_var = ctk.StringVar(value=default_out)
        ctk.CTkEntry(frm_pdf, textvariable=self.out_dir_var).grid(
            row=1, column=1, padx=(0, 8), pady=(0, 12), sticky="ew"
        )
        ctk.CTkButton(frm_pdf, text="参照", width=60, command=self._browse_out_dir).grid(
            row=1, column=2, padx=(0, 12), pady=(0, 12)
        )

        # --- シート名 / 振込金額 ---
        frm_info = ctk.CTkFrame(self)
        frm_info.grid(row=3, column=0, padx=24, pady=6, sticky="ew")
        frm_info.grid_columnconfigure(5, weight=1)

        _now = datetime.date.today()

        ctk.CTkLabel(frm_info, text="シート名（年月）", width=110, anchor="w").grid(
            row=0, column=0, padx=(16, 8), pady=12
        )
        self.year_var = ctk.StringVar(value=str(_now.year))
        self.month_var = ctk.StringVar(value=str(_now.month))
        years = [str(y) for y in range(_now.year - 1, _now.year + 3)]
        months = [str(m) for m in range(1, 13)]
        ctk.CTkOptionMenu(frm_info, variable=self.year_var, values=years, width=80).grid(
            row=0, column=1, padx=(0, 4), pady=12, sticky="w"
        )
        ctk.CTkLabel(frm_info, text="年", anchor="w").grid(row=0, column=2, padx=(0, 8), pady=12)
        ctk.CTkOptionMenu(frm_info, variable=self.month_var, values=months, width=64).grid(
            row=0, column=3, padx=(0, 4), pady=12, sticky="w"
        )
        ctk.CTkLabel(frm_info, text="月", anchor="w").grid(row=0, column=4, padx=(0, 8), pady=12)

        ctk.CTkLabel(frm_info, text="振込金額（税込）", width=110, anchor="w").grid(
            row=1, column=0, padx=(16, 8), pady=(0, 12)
        )
        self.furikomi_var = ctk.StringVar()
        ctk.CTkEntry(
            frm_info, textvariable=self.furikomi_var,
            placeholder_text="空欄でもOK", width=200
        ).grid(row=1, column=1, columnspan=4, padx=(0, 16), pady=(0, 12), sticky="w")

        ctk.CTkLabel(frm_info, text="税込相殺", width=110, anchor="w").grid(
            row=2, column=0, padx=(16, 8), pady=(0, 12)
        )
        self.sousai_var = ctk.StringVar()
        ctk.CTkEntry(
            frm_info, textvariable=self.sousai_var,
            placeholder_text="例: -15000（空欄で0扱い）", width=200
        ).grid(row=2, column=1, columnspan=4, padx=(0, 16), pady=(0, 12), sticky="w")

        # --- 処理開始ボタン ---
        self.run_btn = ctk.CTkButton(
            self, text="▶  処理開始",
            font=ctk.CTkFont(size=15, weight="bold"),
            height=44, corner_radius=10,
            command=self._start
        )
        self.run_btn.grid(row=4, column=0, padx=24, pady=(12, 6), sticky="ew")

        # --- 進捗バー ---
        self.progress = ctk.CTkProgressBar(self, mode="indeterminate")
        self.progress.grid(row=5, column=0, padx=24, pady=(0, 8), sticky="ew")
        self.progress.set(0)

        # --- ログ ---
        self.log_box = ctk.CTkTextbox(
            self, height=180, font=ctk.CTkFont(family="Consolas", size=11),
            state="disabled", wrap="word"
        )
        self.log_box.grid(row=6, column=0, padx=24, pady=(0, 6), sticky="ew")

        # --- 更新確認ボタン ---
        frm_bottom = ctk.CTkFrame(self, fg_color="transparent")
        frm_bottom.grid(row=7, column=0, padx=24, pady=(0, 12), sticky="ew")
        frm_bottom.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            frm_bottom, text=f"v{VERSION}", text_color="gray", font=ctk.CTkFont(size=11)
        ).grid(row=0, column=0, sticky="w")
        ctk.CTkButton(
            frm_bottom, text="更新を確認", width=100, height=28,
            command=lambda: threading.Thread(
                target=self._run_update_check, args=(True,), daemon=True
            ).start()
        ).grid(row=0, column=1, sticky="e")

    # --------------------------------------------------------------- ファイル選択
    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="支払通知書PDFを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if path:
            self.pdf_var.set(path)

    def _browse_out_dir(self):
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.out_dir_var.set(path)
            self._persist_settings()

    # --------------------------------------------------------------- 処理実行
    def _start(self):
        pdf_path = self.pdf_var.get().strip()
        tpl_path = _bundled_template()
        sheet = f"{self.year_var.get()}年{self.month_var.get()}月"
        out_dir = self.out_dir_var.get().strip()

        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("エラー", "PDFファイルを選択してください")
            return
        if not os.path.exists(tpl_path):
            messagebox.showerror("エラー", f"テンプレートが見つかりません:\n{tpl_path}")
            return
        if not out_dir or not os.path.isdir(out_dir):
            messagebox.showerror("エラー", "出力先フォルダを選択してください")
            return

        furikomi_raw = self.furikomi_var.get().strip().replace(",", "")
        if furikomi_raw and not furikomi_raw.lstrip("-").isdigit():
            messagebox.showerror("入力エラー", f"振込金額は半角数字で入力してください: {furikomi_raw}")
            return
        furikomi = int(furikomi_raw) if furikomi_raw else None

        sousai_raw = self.sousai_var.get().strip().replace(",", "")
        if sousai_raw and not sousai_raw.lstrip("-").isdigit():
            messagebox.showerror("入力エラー", f"税込相殺は半角数字で入力してください: {sousai_raw}")
            return
        sousai = int(sousai_raw) if sousai_raw else 0

        self.run_btn.configure(state="disabled")
        self.progress.start()
        self._log_clear()
        self._log("処理を開始します...\n")

        threading.Thread(
            target=self._run_extraction,
            args=(pdf_path, tpl_path, sheet, furikomi, sousai, out_dir),
            daemon=True
        ).start()

    def _run_extraction(self, pdf_path: str, tpl_path: str, sheet: str, furikomi, sousai: int, out_dir: str):
        try:
            from plumber_extractor import extract_with_pdfplumber, extract_payment_date
            from excel_writer import classify_and_aggregate, write_to_template

            out_path = Path(out_dir) / f"集計用_{sheet}.xlsx"

            self._log("[Step 1] PDFから明細を抽出 (pdfplumber)")
            plumber_result = extract_with_pdfplumber(pdf_path)
            if not plumber_result or not plumber_result.get("rows"):
                raise RuntimeError(
                    "PDFから明細を抽出できませんでした。\n"
                    "テキストPDF（コピペできるPDF）である必要があります。\n"
                    "画像PDFには対応していません。"
                )
            rows = plumber_result["rows"]
            self._log(f"  抽出行数: {len(rows)}")

            self._log("[Step 2] 支払日を抽出")
            payment_date = extract_payment_date(pdf_path)
            if payment_date:
                self._log(f"  支払日: {payment_date}")

            self._log("[Step 3] 邸別集計 & 分類ルール適用")
            aggregated = classify_and_aggregate(rows)
            self._log(f"  集計後の邸数: {len(aggregated)}")

            self._log("[Step 4] Excel に書き込み")
            write_to_template(
                template_path=tpl_path,
                output_path=str(out_path),
                sheet_name=sheet,
                aggregated=aggregated,
                furikomi_kingaku=furikomi,
                pdf_koujidai_zeikomi=None,
                pdf_sousai_zeikomi=sousai if sousai != 0 else None,
                payment_date=payment_date,
            )

            self._log(f"\n✅ 完了: {out_path}")
            self.after(0, lambda: self._on_success(str(out_path)))

        except PermissionError as e:
            err_path = getattr(e, "filename", "") or ""
            friendly = (
                "出力ファイルに書き込めません。\n\n"
                f"ファイル: {err_path}\n\n"
                "同じファイル名のExcelを開いているとロックされます。\n"
                "以下を確認してください:\n"
                "  ① 前回の出力Excelを閉じる（Excelのタスクバー確認）\n"
                "  ② OneDrive同期中なら少し待ってから再実行\n"
                "  ③ 出力先フォルダに書き込み権限があるか"
            )
            self._log(f"\n❌ ファイルが開かれています: {err_path}")
            self.after(0, lambda: self._on_permission_error(friendly))
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

    def _on_permission_error(self, msg: str):
        self.progress.stop()
        self.progress.set(0)
        self.run_btn.configure(state="normal")
        messagebox.showerror("ファイルが開かれています", msg)

    # --------------------------------------------------------------- ログ
    def _log(self, msg: str):
        def _write():
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")
        self.after(0, _write)

    def _log_clear(self):
        def _clear():
            self.log_box.configure(state="normal")
            self.log_box.delete("1.0", "end")
            self.log_box.configure(state="disabled")
        self.after(0, _clear)


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()

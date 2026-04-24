"""
支払通知書自動抽出ツール 設定ファイル

AI機能は廃止されました。pdfplumber でテキストPDFのみ処理します。
"""

import os


CONFIG = {
    # ===== 入出力パス =====
    "input_pdf": "input/支払通知書.pdf",
    "template_xlsx": "template/集計用.xlsx",
    "output_xlsx": "output/集計用_自動反映.xlsx",
    "image_temp_dir": os.path.join(
        os.getenv("APPDATA", os.path.expanduser("~")), "invoice-tool", "temp"
    ),
    "sheet_name": "2025年1月",

    # ===== 以下は廃止されたが、呼び出し側の互換性のため残す =====
    "openrouter_api_key": "",
    "openrouter_base_url": "",
    "models": [],
    "run_pdfplumber_in_parallel": True,
    "max_workers": 4,
    "voting_strategy": "all_agree",
    "discrepancy_tolerance": 10,
    "image_dpi": 250,
    "image_format": "JPEG",
    "jpeg_quality": 90,

    # ===== 分類ルール（ドキュメント用） =====
    "classification_rules": {
        "①税抜(D)": "プラス金額をすべて加算（社保プラス・柱脚含む）",
        "②社保(E)": "マイナス×工種が「防水(社保)」×備考に「生産課中口分」を含む → 絶対値",
        "③生産課(F)": "マイナス×社保以外×備考に「生産課中口分」を含む → 絶対値",
        "④材料費(G)": "防水シート(相殺) + 上記に当てはまらないマイナス → 絶対値",
    },
}

"""
支払通知書自動抽出ツール 設定ファイル

使い方:
- models リストに使いたいAIモデルを並べる
- 1個なら単体実行、2個以上ならクロスチェックモードで並列実行
- pdfplumber は常に並列で試行（成功すれば最優先で採用）

APIキーの取得優先順位:
1. _secret.py (PyInstallerビルド時にGitHub Actionsが生成)
2. 環境変数 OPENROUTER_API_KEY
3. .env ファイル（ローカル開発用）
"""

import os

def _load_api_key() -> str:
    # 1. ビルド時埋め込み (_secret.py はGitignore済み、GitHub Actionsが生成)
    try:
        from _secret import OPENROUTER_API_KEY
        return OPENROUTER_API_KEY
    except ImportError:
        pass

    # 2. 環境変数
    key = os.environ.get("OPENROUTER_API_KEY", "")
    if key:
        return key

    # 3. .env ファイル（ローカル開発用）
    env_path = os.path.join(os.path.dirname(__file__), ".env")
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("OPENROUTER_API_KEY="):
                    return line.split("=", 1)[1].strip()

    return ""


CONFIG = {
    # ===== OpenRouter 認証 =====
    "openrouter_api_key": _load_api_key(),
    "openrouter_base_url": "https://openrouter.ai/api/v1",

    # ===== 使用するAIモデル（リスト形式で複数指定可）=====
    # 1個 → 単体モード
    # 2個以上 → クロスチェックモード（並列実行、結果を突き合わせ）
    "models": [
        "anthropic/claude-opus-4.6",
        # "google/gemini-3.1-pro-preview",     # コメント外せば2個になる
        # "openai/gpt-5.4",                     # コメント外せば3個になる
    ],

    # ===== 並列実行の設定 =====
    "run_pdfplumber_in_parallel": True,  # pdfplumberをAIと並列に試行する
    "max_workers": 4,                     # 同時実行スレッド数

    # ===== クロスチェック（モデル2個以上のとき有効）=====
    "voting_strategy": "all_agree",  # "all_agree"(全員一致) or "majority"(多数決)
    "discrepancy_tolerance": 10,     # 金額の許容差(円)。このくらいまでは一致とみなす

    # ===== 画像変換設定 =====
    "image_dpi": 250,           # PDF→JPEG の解像度（高いほど精度◎ トークン◎）
    "image_format": "JPEG",
    "jpeg_quality": 90,

    # ===== 入出力パス =====
    "input_pdf": "input/支払通知書.pdf",
    "template_xlsx": "template/集計用.xlsx",
    "output_xlsx": "output/集計用_自動反映.xlsx",
    "image_temp_dir": "temp_images",
    "sheet_name": "2025年1月",  # 新規作成するシート名

    # ===== 分類ルール（プロンプトに埋め込まれる）=====
    "classification_rules": {
        "①税抜(D)": "プラス金額をすべて加算（社保プラス・柱脚含む）",
        "②社保(E)": "マイナス×工種が「防水(社保)」×備考に「生産課中口分」を含む → 絶対値",
        "③生産課(F)": "マイナス×社保以外×備考に「生産課中口分」を含む → 絶対値",
        "④材料費(G)": "防水シート(相殺) + 上記に当てはまらないマイナス → 絶対値",
    },
}

# 支払通知書 自動抽出ツール

内装工事会社の支払通知書（PDF）から明細を自動抽出し、
集計用Excelテンプレートに反映するツール。

## 特徴

- **並列処理**: pdfplumber と AIモデルを同時並行で実行
- **テキストPDF優先**: pdfplumberで読めれば無料・高速でAI不使用
- **AIフォールバック**: 画像PDFの場合はAIが自動で引き継ぎ
- **拡張可能**: AIモデルを何個でも追加可能（クロスチェック対応）
- **分類ルール自動適用**: 社保/生産課/材料費への振り分けまで自動
- **自動アップデート**: GitHub Releases経由で新バージョンを自動取得

## アーキテクチャ

```
[input.pdf]
   ↓
[Step 0] アプリ起動時 → 自動更新チェック（新版あれば案内）
   ↓
[Step 1] PDF → JPEG (各ページを画像化)
   ↓
[Step 2] 並列実行
   ├─ pdfplumber (テキストPDFなら成功)
   ├─ Claude Opus 4.6 (画像PDFならこっち)
   └─ （設定でAIモデル追加可能）
   ↓
[Step 3] 結果採用
   ↓
[Step 4] 分類ルール適用
   ↓
[Step 5] Excel新シートに書き込み
   ↓
[output.xlsx]
```

## エンドユーザー向け：使い方

1. `shiharai-tool.exe` をダブルクリック
2. 新しいバージョンがある場合は案内が出る（「はい」で自動更新）
3. 処理完了を待つ
4. 出力Excelが生成される

## 開発者向け：セットアップ

### 1. Python依存関係のインストール

```bash
pip install -r requirements.txt
```

### 2. OpenRouter APIキー取得

1. https://openrouter.ai にサインアップ
2. Keys メニューからAPIキー発行 (`sk-or-v1-xxx`)

### 3. 設定ファイル編集

`config.py` を開いてAPIキーを設定。

### 4. 開発時実行

```bash
python main.py
```

### 5. ローカルビルド

```bash
pip install pyinstaller
pyinstaller shiharai-tool.spec
# → dist/shiharai-tool.exe が生成される
```

## AIモデルの追加/変更

`config.py` の `models` リストを書き換えるだけ:

```python
# 1個モード（現在）
"models": ["anthropic/claude-opus-4.6"],

# 2個モード（クロスチェック）
"models": [
    "anthropic/claude-opus-4.6",
    "google/gemini-3.1-pro-preview",
],

# 3個モード（多数決可能）
"models": [
    "anthropic/claude-opus-4.6",
    "google/gemini-3.1-pro-preview",
    "openai/gpt-5.4",
],
```

## リリース手順

`RELEASE_GUIDE.md` を参照。基本的には：

```bash
# コード修正後
# version.py の VERSION を上げる
git add .
git commit -m "fix: ..."
git tag v1.0.1
git push --tags
# ↑ これだけで GitHub Actions が自動で .exe をビルド＆リリース
# ↑ 山本さんのPCが次回起動時に自動更新
```

## ファイル構成

```
pdf_extractor/
├── main.py                        # エントリーポイント
├── version.py                     # バージョン情報（リリース時に編集）
├── config.py                      # 設定
├── updater.py                     # 自動アップデート機能
├── pdf_converter.py               # PDF → JPEG 変換
├── plumber_extractor.py           # pdfplumber抽出（テキストPDF用）
├── ai_extractor.py                # AI抽出（画像PDF用）
├── orchestrator.py                # 並列実行 & 結果採用ロジック
├── excel_writer.py                # Excel書き込み & 分類ルール
├── requirements.txt               # Python依存関係
├── shiharai-tool.spec             # PyInstaller設定
├── .github/
│   └── workflows/
│       └── build.yml              # GitHub Actions (自動ビルド)
├── .gitignore
├── README.md
└── RELEASE_GUIDE.md               # リリース手順書（陸くん用）
```

## 注意事項

- **AI APIの利用**: PDFの内容は一時的にOpenRouter経由でAIプロバイダーに送信されます。
  - AIの学習には使われません（API利用規約による）
  - 30日後にプロバイダー側で削除されます
  - 機密情報が含まれる場合は事前に依頼者と合意を取ってください
- **コスト**: Claude Opus 4.6 単体で1回あたり約45円。月1回なら年間500円程度。
- **処理時間**: pdfplumber成功なら数秒、AI使用時は30秒〜1分程度。
- **Windows向け**: 自動アップデート機能はWindowsのみ対応。Mac版は別途ビルドが必要。


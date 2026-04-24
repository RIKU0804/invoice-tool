# invoice-tool

支払通知書PDFから明細を自動抽出し、集計用Excelに反映するWindows GUIツール。

## 特徴

- **pdfplumberでローカル抽出**: テキストPDFを完全オフラインで解析
- **邸別自動集計**: 社保・生産課・材料費などの分類ルールを自動適用
- **動的レイアウト**: 1〜50邸まで邸数に応じて行を自動伸縮
- **振込金額/税込相殺の自動入力**: PDFから合計行を読み取り、根拠画像も視覚確認
- **支払日の自動抽出**: ExcelのK1セルに自動記載
- **自動アップデート**: GitHub Releases経由で新バージョンを自動検出・更新

## 処理フロー

```
[入力PDF]
  ↓ pdfplumber
[明細行の抽出]
  ↓ 分類ルール
[邸別集計 (D/E/F/G列)]
  ↓
[集計用.xlsx テンプレートに書き込み]
  ↓
[出力 Excel]
```

## 開発者向け

### セットアップ

```bash
pip install -r requirements.txt
python gui.py
```

### ローカルビルド

```bash
pip install pyinstaller
pyinstaller invoice-tool.spec
# → dist/invoice-tool.exe
```

### リリース

```bash
# version.py の VERSION を更新
git add .
git commit -m "fix: xxx"
git tag v1.0.XX
git push origin main
git push origin v1.0.XX
# → GitHub Actions が自動で invoice-tool.exe をビルド→Releases公開
# → 既存ユーザーのアプリが次回起動時に自動更新検出
```

## ファイル構成

```
invoice-tool/
├── gui.py                  # エントリーポイント (customtkinter GUI)
├── version.py              # VERSION 定数
├── config.py               # 分類ルール定義
├── updater.py              # GitHub Releases チェック + 自動更新
├── plumber_extractor.py    # pdfplumber 抽出 (明細/支払日/合計行)
├── excel_writer.py         # 集計ロジック + Excel 書き込み
├── requirements.txt
├── invoice-tool.spec       # PyInstaller (onefile, UPX=False)
├── template/
│   └── 集計用.xlsx         # Excel テンプレート (exeにバンドル)
└── .github/workflows/
    └── build.yml           # タグpushで自動ビルド&リリース
```

## システム要件

- Windows 10 (64bit) / Windows 11
- テキストPDF（コピペ可能なPDF）

## ライセンス

Private project. Source is visible for auto-update distribution purposes.

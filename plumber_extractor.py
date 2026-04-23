"""
pdfplumber 抽出モジュール

テキストPDFから表構造を直接抽出する。画像PDFでは失敗する（それが正常）。
成功時：高速・無料・高精度
失敗時：Noneを返し、呼び出し側でAI抽出にフォールバック
"""

import pdfplumber
import re


def extract_with_pdfplumber(pdf_path: str) -> dict | None:
    """
    pdfplumberで支払通知書から明細を抽出する。

    Returns:
        成功: {"source": "pdfplumber", "rows": [...], "totals": {...}}
        失敗: None（画像PDF、表認識失敗など）
    """
    try:
        all_rows = []

        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                # テキスト抽出で内容があるか判定（画像PDFだと空）
                text = page.extract_text() or ""
                if len(text.strip()) < 50:
                    print(f"  [pdfplumber] Page {page_num}: テキストなし → 画像PDFの可能性")
                    return None

                # 表を抽出
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        if not row or len(row) < 8:
                            continue
                        parsed = _parse_row(row)
                        if parsed:
                            all_rows.append(parsed)

        if not all_rows:
            print("  [pdfplumber] 明細行を検出できず")
            return None

        return {
            "source": "pdfplumber",
            "rows": all_rows,
            "row_count": len(all_rows),
        }

    except Exception as e:
        print(f"  [pdfplumber] エラー: {e}")
        return None


def _parse_row(row: list) -> dict | None:
    """
    1行のデータを構造化する。
    期待形式: [事業所, 契約NO, 邸名, 工種, 税抜, 消費税, 税込, 備考]
    """
    try:
        jigyosho, keiyaku_no, tei_mei, koushu, zeinuki, shohizei, zeikomi, bikou = row[:8]

        # 金額のマイナス(▲)を負数に正規化
        zeinuki_val = _parse_amount(zeinuki)
        if zeinuki_val is None:
            return None

        return {
            "事業所": (jigyosho or "").strip(),
            "契約NO": (keiyaku_no or "").strip(),
            "邸名": (tei_mei or "").strip(),
            "工種": (koushu or "").strip(),
            "税抜金額": zeinuki_val,
            "消費税": _parse_amount(shohizei) or 0,
            "税込金額": _parse_amount(zeikomi) or 0,
            "備考": (bikou or "").strip(),
        }
    except (ValueError, IndexError):
        return None


def _parse_amount(s) -> int | None:
    """金額文字列を整数に変換（▲や−を負号として扱う）"""
    if s is None:
        return None
    s = str(s).strip()
    if not s:
        return None
    is_negative = s.startswith("▲") or s.startswith("-") or s.startswith("−")
    s_clean = re.sub(r"[^\d]", "", s)
    if not s_clean:
        return None
    val = int(s_clean)
    return -val if is_negative else val


if __name__ == "__main__":
    from config import CONFIG
    result = extract_with_pdfplumber(CONFIG["input_pdf"])
    if result:
        print(f"抽出成功: {result['row_count']}行")
        for row in result['rows'][:3]:
            print(row)
    else:
        print("抽出失敗（画像PDFの可能性）")

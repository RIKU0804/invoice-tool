"""
pdfplumber 抽出モジュール

テキストPDFから表構造を直接抽出する。画像PDFでは失敗する（それが正常）。
成功時：高速・無料・高精度
失敗時：Noneを返し、呼び出し側でAI抽出にフォールバック
"""

import pdfplumber
import re
from typing import Optional


def extract_payment_date(pdf_path: str) -> Optional[str]:
    """PDFから「支払日」行を検出して "YYYY年MM月DD日" 形式で返す"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                m = re.search(r'支払日\s*(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日', text)
                if m:
                    return f"{m.group(1)}年{m.group(2).zfill(2)}月{m.group(3).zfill(2)}日"
    except Exception as e:
        print(f"  [支払日抽出] エラー: {e}")
    return None


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

        zeinuki_val = _parse_amount(zeinuki)
        if zeinuki_val is None:
            return None

        def _s(v):
            return str(v).strip() if v is not None else ""

        return {
            "事業所": _s(jigyosho),
            "契約NO": _s(keiyaku_no),
            "邸名": _s(tei_mei),
            "工種": _s(koushu),
            "税抜金額": zeinuki_val,
            "消費税": _parse_amount(shohizei) or 0,
            "税込金額": _parse_amount(zeikomi) or 0,
            "備考": _s(bikou),
        }
    except Exception as e:
        print(f"  [pdfplumber] row parse failed: {type(e).__name__}: {e} row={row!r}")
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

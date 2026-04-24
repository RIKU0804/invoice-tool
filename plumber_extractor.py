"""
pdfplumber 抽出モジュール

テキストPDFから表構造を直接抽出する。
画像PDFには対応しない方針。
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


def extract_totals_and_snippet(pdf_path: str, snippet_out_path: Optional[str] = None) -> dict:
    """
    PDFから振込金額(合計の税込)と税込相殺を抽出し、
    根拠となる画像スニペットをPNG保存する。

    Returns:
        {"furikomi": int|None, "sousai": int|None, "snippet_path": str|None}
    """
    result = {"furikomi": None, "sousai": None, "snippet_path": None}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            # 合計・相殺計が載っている最終ページを探す
            target_page = None
            for page in pdf.pages:
                text = page.extract_text() or ""
                if ("合計" in text) and ("相殺" in text or "工事代" in text):
                    target_page = page
            if target_page is None:
                return result

            text = target_page.extract_text() or ""
            # 合計行（末尾の大きな数字3つ）: 税抜, 消費税, 税込
            m = re.search(r'合計\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)', text)
            if m:
                result["furikomi"] = int(m.group(3).replace(",", ""))
            # ＜相殺 計＞行
            m2 = re.search(r'＜相殺\s*計＞\s*([▲▽\-−]?[\d,]+)\s+([▲▽\-−]?[\d,]*)\s*([▲▽\-−]?[\d,]+)', text)
            if m2:
                def _to_int(s):
                    s = s.strip()
                    neg = s.startswith(("▲", "-", "−", "▽"))
                    s = re.sub(r'[^\d]', '', s)
                    if not s:
                        return 0
                    return -int(s) if neg else int(s)
                result["sousai"] = _to_int(m2.group(3))

            # スニペット画像を保存（工事代計〜合計の範囲を切り出す）
            if snippet_out_path:
                words = target_page.extract_words()
                top_y = None
                bottom_y = None
                for w in words:
                    if '工事代' in w['text']:
                        top_y = w['top'] - 4
                    if w['text'] == '合計' and bottom_y is None:
                        bottom_y = w['bottom'] + 4
                if top_y is not None and bottom_y is not None and bottom_y > top_y:
                    try:
                        cropped = target_page.crop((0, top_y, target_page.width, bottom_y))
                        img = cropped.to_image(resolution=200)
                        img.save(snippet_out_path, format="PNG")
                        result["snippet_path"] = snippet_out_path
                    except Exception as e:
                        print(f"  [スニペット生成] エラー: {e}")
    except Exception as e:
        print(f"  [合計抽出] エラー: {e}")
    return result


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
    import sys
    pdf = sys.argv[1] if len(sys.argv) > 1 else "input/支払通知書.pdf"
    result = extract_with_pdfplumber(pdf)
    if result:
        print(f"抽出成功: {result['row_count']}行")
        for row in result['rows'][:3]:
            print(row)
    else:
        print("抽出失敗（画像PDFの可能性）")

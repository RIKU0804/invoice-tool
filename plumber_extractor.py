"""
pdfplumber 抽出モジュール

テキストPDFから表構造を直接抽出する。
画像PDFには対応しない方針。
"""

import pdfplumber
import re
from typing import Optional

# カラム名の別名マッピング（柔軟なヘッダー検出用）
_COLUMN_ALIASES: dict[str, list[str]] = {
    "事業所":   ["事業所"],
    "契約NO":   ["契約NO", "契約No", "契約番号", "契約no"],
    "邸名":     ["邸名", "物件名"],
    "工種":     ["工種"],
    "税抜金額": ["税抜金額", "税抜", "金額(税抜)", "金額（税抜）"],
    "消費税":   ["消費税", "税額"],
    "税込金額": ["税込金額", "税込", "金額(税込)", "金額（税込）"],
    "備考":     ["備考", "摘要"],
}

# ヘッダーなし時のデフォルト位置インデックス（従来の挙動）
_DEFAULT_COL_MAP: dict[str, int] = {
    "事業所":   0,
    "契約NO":   1,
    "邸名":     2,
    "工種":     3,
    "税抜金額": 4,
    "消費税":   5,
    "税込金額": 6,
    "備考":     7,
}

# ヘッダー判定に必要な最低限のキー列
_REQUIRED_HEADER_COLS = {"邸名", "工種", "税抜金額", "税込金額"}


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


def _to_int_amount(s: str) -> int:
    s = s.strip()
    neg = s.startswith(("▲", "-", "−", "▽"))
    digits = re.sub(r'[^\d]', '', s)
    if not digits:
        return 0
    return -int(digits) if neg else int(digits)


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
            target_page = None
            for page in pdf.pages:
                text = page.extract_text() or ""
                if ("合計" in text) and ("相殺" in text or "工事代" in text):
                    target_page = page
            if target_page is None:
                return result

            text = target_page.extract_text() or ""
            # 最後の合計マッチ（サブトータル誤認回避）
            all_goukei = re.findall(r'合計\s+([\d,]+)\s+([\d,]+)\s+([\d,]+)', text)
            if all_goukei:
                result["furikomi"] = int(all_goukei[-1][2].replace(",", ""))

            # 相殺計の形式は 2 通りある:
            #   3列形式: ＜相殺 計＞ -15,000 0 -15,000  (税抜 消費税 税込)
            #   2列形式: ＜相殺 計＞ -718,450 -718,450 (税抜=税込、消費税列省略=非課税)
            # 最後の数字を税込として採用する
            m_sousai = re.search(
                r'＜相殺\s*計＞\s*([▲▽\-−]?[\d,]+)(?:\s+([▲▽\-−]?[\d,]+))?(?:\s+([▲▽\-−]?[\d,]+))?',
                text,
            )
            if m_sousai:
                nums = [g for g in m_sousai.groups() if g]
                if nums:
                    result["sousai"] = _to_int_amount(nums[-1])

            if snippet_out_path:
                words = target_page.extract_words()
                top_y = bottom_y = None
                left_x = None
                for w in words:
                    if '工事代' in w['text']:
                        top_y = w['top'] - 4
                    if w['text'] == '合計' and bottom_y is None:
                        bottom_y = w['bottom'] + 4
                    if '工事代' in w['text'] or '相殺' in w['text'] or w['text'] == '合計':
                        if left_x is None or w['x0'] < left_x:
                            left_x = w['x0']
                left_x = max(0.0, (left_x or 230.0) - 6)
                if top_y is not None and bottom_y is not None and bottom_y > top_y:
                    try:
                        cropped = target_page.crop((left_x, top_y, target_page.width, bottom_y))
                        img = cropped.to_image(resolution=300)
                        img.save(snippet_out_path, format="PNG")
                        result["snippet_path"] = snippet_out_path
                    except Exception as e:
                        print(f"  [スニペット生成] エラー: {e}")
    except Exception as e:
        print(f"  [合計抽出] エラー: {e}")
    return result


def _detect_column_map(row: list) -> dict[str, int] | None:
    """
    ヘッダー行かどうかを判定し、カラムマップを返す。
    ヘッダーと判断できない場合は None を返す。
    """
    def _norm(v: object) -> str:
        return str(v).strip() if v is not None else ""

    cells = [_norm(c) for c in row]

    col_map: dict[str, int] = {}
    for canonical, aliases in _COLUMN_ALIASES.items():
        for i, cell in enumerate(cells):
            if cell in aliases:
                col_map[canonical] = i
                break

    found = set(col_map.keys())
    if _REQUIRED_HEADER_COLS.issubset(found):
        return col_map
    return None


def _cell(row: list, col_map: dict[str, int], key: str) -> object:
    """col_map からセルを取得。インデックス範囲外は None"""
    idx = col_map.get(key)
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def extract_with_pdfplumber(pdf_path: str) -> dict | None:
    """
    pdfplumberで支払通知書から明細を抽出する。

    Returns:
        成功: {
            "source": "pdfplumber",
            "rows": [...],
            "row_count": int,
            "pdf_koujidai_zeinuki": int|None,  # PDF記載の工事代計（税抜）
            "col_map_used": "header"|"positional",
        }
        失敗: None（画像PDF、表認識失敗など）
    """
    try:
        all_rows: list[dict] = []
        pdf_koujidai_zeinuki: int | None = None
        _pdf_koujidai_zeikomi: int | None = None
        col_map_mode = "positional"

        with pdfplumber.open(pdf_path) as pdf:
            any_text_page = False
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text() or ""
                if len(text.strip()) < 50:
                    print(f"  [pdfplumber] Page {page_num}: テキスト少 → スキップ")
                    continue
                any_text_page = True

                # ページテキストから工事代計を抽出
                # 「＜工事代 計＞ 税抜 消費税 税込」の形式
                m_koujidai = re.search(
                    r'＜工事代\s*計＞\s*([\d,]+)\s+([\d,]+)\s+([\d,]+)', text
                )
                if m_koujidai and pdf_koujidai_zeinuki is None:
                    pdf_koujidai_zeinuki = int(m_koujidai.group(1).replace(",", ""))
                    _pdf_koujidai_zeikomi = int(m_koujidai.group(3).replace(",", ""))

                tables = page.extract_tables()
                for table in tables:
                    if not table:
                        continue

                    # ヘッダー行の検出
                    active_col_map: dict[str, int] = _DEFAULT_COL_MAP
                    start_idx = 0
                    detected = _detect_column_map(table[0])
                    if detected is not None:
                        active_col_map = detected
                        col_map_mode = "header"
                        start_idx = 1  # ヘッダー行自体はスキップ

                    for row in table[start_idx:]:
                        if not row:
                            continue

                        # 列数チェック（足りなくても処理続行、警告のみ）
                        max_idx = max(active_col_map.values())
                        if len(row) <= max_idx:
                            print(
                                f"  [pdfplumber] 列数不足 (期待>={max_idx+1}, 実際={len(row)})"
                                f" row={row!r} → スキップ"
                            )
                            continue

                        parsed = _parse_row_mapped(row, active_col_map)
                        if parsed:
                            all_rows.append(parsed)

            if not any_text_page:
                print("  [pdfplumber] 全ページでテキスト検出失敗 → 画像PDFの可能性")
                return None

        if not all_rows:
            print("  [pdfplumber] 明細行を検出できず")
            return None

        # 邸名キャリーフォワード
        # PDFが邸名を最初の行にしか書かない場合、pdfplumberは後続行の邸名を空白で読む。
        # 工種があり集計行でない空邸名行には直前の邸名を引き継ぐ。
        last_valid_tei = ""
        carry_count = 0
        for row in all_rows:
            tei = row["邸名"]
            if tei:
                last_valid_tei = tei
            elif last_valid_tei and row["工種"] and not row["工種"].startswith("＜"):
                row["邸名"] = last_valid_tei
                carry_count += 1
        if carry_count:
            print(f"  [pdfplumber] 邸名キャリーフォワード: {carry_count}行 補完")

        # 照合チェック
        # 集計行(邸名='', '合計', '消費税 対象外' 等)は除外 — classify_and_aggregate と同じ条件
        # PDFは「税込合計 / 1.1」で税抜を逆算するため税抜での照合は誤差が出る。
        # 税込で照合することで計算方式の違いに左右されない正確な照合ができる。
        if _pdf_koujidai_zeikomi is not None:
            meisai = [
                r for r in all_rows
                if r["邸名"]
                and r["邸名"] not in ("計", "合計")
                and "消費税" not in r["邸名"]
                and "対象外" not in r["邸名"]
            ]
            extracted_zeikomi = sum(r["税込金額"] for r in meisai)
            diff_zeikomi = extracted_zeikomi - _pdf_koujidai_zeikomi
            extracted_zeinuki = sum(r["税抜金額"] for r in meisai)
            if abs(diff_zeikomi) > 10:
                print(
                    f"  [pdfplumber] [WARNING] 照合差異(税込): 抽出={extracted_zeikomi:,} / "
                    f"PDF={_pdf_koujidai_zeikomi:,} / 差={diff_zeikomi:+,}円 "
                    f"→ フォーマット変化の可能性"
                )
            else:
                print(
                    f"  [pdfplumber] [OK] 照合(税込): 差={diff_zeikomi:+,}円 "
                    f"(抽出税込={extracted_zeikomi:,} / PDF税込={_pdf_koujidai_zeikomi:,})"
                )
        elif pdf_koujidai_zeinuki is not None:
            meisai = [
                r for r in all_rows
                if r["邸名"]
                and r["邸名"] not in ("計", "合計")
                and "消費税" not in r["邸名"]
                and "対象外" not in r["邸名"]
            ]
            extracted_zeinuki = sum(r["税抜金額"] for r in meisai)
            diff = extracted_zeinuki - pdf_koujidai_zeinuki
            if abs(diff) > 10:
                print(
                    f"  [pdfplumber] [WARNING] 照合差異(税抜): 抽出={extracted_zeinuki:,} / "
                    f"PDF={pdf_koujidai_zeinuki:,} / 差={diff:+,}円"
                )
            else:
                print(
                    f"  [pdfplumber] [OK] 照合(税抜): 差={diff:+,}円"
                )

        return {
            "source": "pdfplumber",
            "rows": all_rows,
            "row_count": len(all_rows),
            "pdf_koujidai_zeinuki": pdf_koujidai_zeinuki,
            "col_map_used": col_map_mode,
        }

    except Exception as e:
        print(f"  [pdfplumber] エラー: {e}")
        return None


def _parse_row_mapped(row: list, col_map: dict[str, int]) -> dict | None:
    """
    カラムマップを使って1行のデータを構造化する。
    col_map は _DEFAULT_COL_MAP（位置ベース）またはヘッダー検出結果のどちらでも可。
    """
    def _s(v: object) -> str:
        return str(v).strip() if v is not None else ""

    try:
        tei_mei = _s(_cell(row, col_map, "邸名"))
        zeinuki_raw = _cell(row, col_map, "税抜金額")
        zeinuki_val = _parse_amount(zeinuki_raw)

        if zeinuki_val is None:
            return None

        return {
            "事業所":   _s(_cell(row, col_map, "事業所")),
            "契約NO":   _s(_cell(row, col_map, "契約NO")),
            "邸名":     tei_mei,
            "工種":     _s(_cell(row, col_map, "工種")),
            "税抜金額": zeinuki_val,
            "消費税":   _parse_amount(_cell(row, col_map, "消費税")) or 0,
            "税込金額": _parse_amount(_cell(row, col_map, "税込金額")) or 0,
            "備考":     _s(_cell(row, col_map, "備考")),
        }
    except Exception as e:
        print(f"  [pdfplumber] row parse failed: {type(e).__name__}: {e} row={row!r}")
        return None


# 後方互換のため残す（gui.py 等から直接呼んでいる場合に備えて）
def _parse_row(row: list) -> dict | None:
    return _parse_row_mapped(row, _DEFAULT_COL_MAP)


def _parse_amount(s: object) -> int | None:
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
        print(f"抽出成功: {result['row_count']}行 (カラムマップ: {result['col_map_used']})")
        if result["pdf_koujidai_zeinuki"] is not None:
            print(f"PDF工事代計(税抜): {result['pdf_koujidai_zeinuki']:,}円")
        for row in result['rows'][:3]:
            print(row)
    else:
        print("抽出失敗（画像PDFの可能性）")

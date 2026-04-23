"""
Excel反映モジュール

抽出した明細データを、既存の集計用.xlsxに新シートとして追加する。
分類ルール（社保/生産課/材料費）も自動適用する。
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
from collections import defaultdict


def classify_and_aggregate(rows: list[dict]) -> list[dict]:
    """
    明細行を邸ごとに集計し、D/E/F/G列に振り分ける。

    ルール:
    - D列(税抜): プラス金額の合計
    - E列(社保): マイナス×防水(社保)×備考「生産課中口分」の絶対値
    - F列(生産課): マイナス×社保以外×備考「生産課中口分」の絶対値
    - G列(材料費): 防水シート相殺 + その他のマイナス
    """
    by_tei = defaultdict(lambda: {
        "邸名": "",
        "契約NO": set(),
        "工事名称": set(),
        "D_items": [],  # プラス金額リスト（後で加算式に）
        "E": 0,
        "F": 0,
        "G_items": [],  # マイナス絶対値リスト
    })

    for row in rows:
        tei = row["邸名"]
        amount = row["税抜金額"]
        koushu = row["工種"]
        bikou = row.get("備考", "")

        # 集計行・特殊行はスキップ
        if not tei or tei in ("計", "合計") or "消費税" in tei or "対象外" in tei:
            continue

        agg = by_tei[tei]
        agg["邸名"] = tei
        agg["契約NO"].add(row.get("契約NO", ""))

        # 工事名称の推定（工種からベース名を抽出）
        base_name = _extract_koji_base(koushu)
        if base_name:
            agg["工事名称"].add(base_name)

        is_seisanka = "生産課中口分" in bikou
        is_shaho = "社保" in koushu
        is_bousuisheet = "防水シート" in koushu

        if amount >= 0:
            # プラス → D列
            agg["D_items"].append(amount)
        else:
            abs_amount = abs(amount)
            if is_seisanka and is_shaho:
                agg["E"] += abs_amount
            elif is_seisanka:
                agg["F"] += abs_amount
            else:
                agg["G_items"].append(abs_amount)

    # dict → list に変換（邸名順）
    result = []
    for tei, agg in by_tei.items():
        koji_names = list(agg["工事名称"])
        koji_label = "・".join(sorted(set(koji_names))) if koji_names else ""
        result.append({
            "邸名": tei,
            "契約NO": list(agg["契約NO"])[0] if agg["契約NO"] else "",
            "工事名称": koji_label,
            "D_items": agg["D_items"],
            "E": agg["E"],
            "F": agg["F"],
            "G_items": agg["G_items"],
        })
    return result


def _extract_koji_base(koushu: str) -> str | None:
    """工種から工事名称ベースを抽出する（集計用Excelの表記に合わせる）"""
    if "防水" in koushu:
        return "防水"
    if "柱脚" in koushu:
        return "柱脚"
    return None


def write_to_template(
    template_path: str,
    output_path: str,
    sheet_name: str,
    aggregated: list[dict],
    furikomi_kingaku: int = None,
    pdf_koujidai_zeikomi: int = None,
    pdf_sousai_zeikomi: int = None,
):
    """
    集計用テンプレートに新シートを追加してデータを書き込む。
    既存シートの書式・計算式パターンを踏襲する。
    """
    wb = load_workbook(template_path)

    # コピー元シートを決める（最新の月次シートをベースに）
    src_sheet_name = wb.sheetnames[0]
    src = wb[src_sheet_name]
    new_ws = wb.copy_worksheet(src)
    new_ws.title = sheet_name

    # 巨大マージセルがあれば解除（以前確認済みの挙動）
    merges_to_unmerge = [str(m) for m in new_ws.merged_cells.ranges if 'B29:J35' in str(m)]
    for m in merges_to_unmerge:
        new_ws.unmerge_cells(m)

    # unmerge後にB29:J35の外枠(赤)を再描画。内側セルの既存罫線は保持
    red_side = Side(style='medium', color='FFC00000')
    for row in range(29, 36):
        for col_idx in range(2, 11):  # B=2, J=10
            is_top = row == 29
            is_bottom = row == 35
            is_left = col_idx == 2
            is_right = col_idx == 10
            if not (is_top or is_bottom or is_left or is_right):
                continue  # 内側はスキップ（既存罫線を保持）
            cell = new_ws.cell(row=row, column=col_idx)
            existing = cell.border
            cell.border = Border(
                left=red_side if is_left else existing.left,
                right=red_side if is_right else existing.right,
                top=red_side if is_top else existing.top,
                bottom=red_side if is_bottom else existing.bottom,
            )

    # タイトル書き換え
    new_ws['C2'] = f'{sheet_name}　着工=受注　ベース'

    # C列（工事名称）の赤塗りをクリア
    no_fill = PatternFill(fill_type=None)
    for r in range(5, 24):
        new_ws.cell(row=r, column=3).fill = no_fill

    # 明細書き込み（5行目から）
    for i, item in enumerate(aggregated):
        r = 5 + i
        new_ws.cell(row=r, column=1, value=i + 1)
        new_ws.cell(row=r, column=2, value=item["邸名"])
        new_ws.cell(row=r, column=3, value=item["工事名称"])
        # D列: プラス金額の加算式
        d_formula = '=' + '+'.join(str(x) for x in item["D_items"]) if item["D_items"] else 0
        new_ws.cell(row=r, column=4, value=d_formula)
        # E, F列
        new_ws.cell(row=r, column=5, value=item["E"] if item["E"] else None)
        new_ws.cell(row=r, column=6, value=item["F"] if item["F"] else None)
        # G列: マイナス絶対値の加算式
        if item["G_items"]:
            g_formula = '=' + '+'.join(str(x) for x in item["G_items"])
            new_ws.cell(row=r, column=7, value=g_formula)
        else:
            new_ws.cell(row=r, column=7, value=None)
        # H, I列: 外注(空欄)
        new_ws[f'H{r}'].value = None
        new_ws[f'I{r}'].value = None
        # K列: 班長(空欄)
        new_ws[f'K{r}'].value = None

    # 空行のクリア（18行より少ない場合）
    for i in range(len(aggregated), 18):
        r = 5 + i
        for col in range(2, 13):
            new_ws.cell(row=r, column=col).value = None

    # 担当者別集計をSUMIFに置き換え
    new_ws.cell(row=29, column=12, value='=SUMIF(K5:K23,K29,J5:J23)')
    new_ws.cell(row=30, column=12, value='=SUMIF(K5:K23,K30,J5:J23)')
    new_ws.cell(row=31, column=12, value='=SUMIF(K5:K23,K31,J5:J23)')

    # 振込金額照合欄
    _write_furikomi_verification(new_ws, furikomi_kingaku, pdf_koujidai_zeikomi, pdf_sousai_zeikomi)

    # 使いやすさ機能（プルダウン・条件付き書式・担当邸数）
    _add_usability_features(new_ws)

    wb.save(output_path)


def _write_furikomi_verification(ws, furikomi, koujidai, sousai):
    """振込金額照合欄（税抜⇔税込の二重計算）"""
    # 既存欄クリア
    for r in range(34, 42):
        for c in range(2, 6):
            ws.cell(r, c).value = None

    ws['B34'] = '【振込金額照合（税抜⇔税込の二重計算）】'
    ws['B34'].font = copy(ws['C2'].font)

    ws['B35'] = '① 振込金額（税込・実際に振り込まれた額）'
    ws['D35'] = furikomi if furikomi else None
    ws['B36'] = '② 税込相殺（PDF・手入力）'
    ws['D36'] = sousai if sousai else None
    ws['B37'] = '③ 税込工事代計（① − ②）'
    ws['D37'] = '=D35-D36'
    ws['B38'] = '④ 税抜逆算（③ ÷ 1.1）'
    ws['D38'] = '=ROUND(D37/1.1,0)'
    ws['B39'] = '⑤ Excel税抜合計（J24）'
    ws['D39'] = '=J24'
    ws['B40'] = '⑥ 差額（⑤ − ④）'
    ws['D40'] = '=D39-D38'
    ws['B41'] = '※ ±数円→インボイス端数差(正常) / 大きな差→PDF読取エラーの可能性'

    for r in [35, 36, 37, 38, 39, 40]:
        ws[f'D{r}'].number_format = '#,##0;[Red]▲#,##0'

    ws['B40'].font = Font(name=ws['B35'].font.name, size=11, bold=True)
    ws['D40'].font = Font(name=ws['D35'].font.name, size=11, bold=True)


def _add_usability_features(ws):
    """プルダウン／条件付き書式／担当邸数カウント"""
    # 班長プルダウン
    dv = DataValidation(
        type='list',
        formula1='"山本,熱田,安保"',
        allow_blank=True,
        showErrorMessage=True,
        errorTitle='班長名エラー',
        error='山本 / 熱田 / 安保 から選んでください',
    )
    dv.add('K5:K23')
    ws.add_data_validation(dv)

    # 差額赤/緑
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    green_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
    ws.conditional_formatting.add('D40', FormulaRule(formula=['ABS(D40)>10'], fill=red_fill))
    ws.conditional_formatting.add('D40', FormulaRule(formula=['AND(ABS(D40)<=10,D40<>"")'], fill=green_fill))

    # 班長未入力黄色
    light_yellow = PatternFill(start_color='FFF9E6', end_color='FFF9E6', fill_type='solid')
    ws.conditional_formatting.add('K5:K23', FormulaRule(formula=['AND(K5="",B5<>"")'], fill=light_yellow))

    # 担当邸数
    ws['N3'] = '【担当邸数】'
    ws['N3'].font = copy(ws['C2'].font)
    ws['N4'] = '班長'
    ws['O4'] = '邸数'
    header_font = Font(name=ws['B5'].font.name, bold=True)
    ws['N4'].font = header_font
    ws['O4'].font = header_font
    ws['N5'] = '山本'
    ws['O5'] = '=COUNTIF(K5:K23,N5)'
    ws['N6'] = '熱田'
    ws['O6'] = '=COUNTIF(K5:K23,N6)'
    ws['N7'] = '安保'
    ws['O7'] = '=COUNTIF(K5:K23,N7)'
    ws['N8'] = '未入力'
    ws['O8'] = '=COUNTBLANK(K5:K23)-COUNTBLANK(B5:B23)'
    ws['N9'] = '合計'
    ws['O9'] = '=SUM(O5:O8)'
    ws.conditional_formatting.add('O8', CellIsRule(operator='greaterThan', formula=['0'], fill=light_yellow))
    ws.column_dimensions['N'].width = 12
    ws.column_dimensions['O'].width = 8
    # 数値列が####にならないよう幅を確保
    for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[col].width = 14

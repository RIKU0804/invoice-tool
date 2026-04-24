"""
Excel反映モジュール

抽出した明細データを、既存の集計用.xlsxに新シートとして追加する。
分類ルール（社保/生産課/材料費）も自動適用する。

邸数によって動的に行を挿入することで18邸を超える場合でも対応する。
基準レイアウト:
    5-22: 明細データ行 (18行)
    23:   スペーサー（空）
    24:   合計行 (SUM式)
    29-31: 班長集計 (SUMIF)
    29-35: 赤枠エリア
    37-44: 振込金額照合
邸数が18を超える場合、行23の前に行を挿入することで範囲を拡張する。
"""

from typing import Optional
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import FormulaRule, CellIsRule, Rule
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.worksheet.datavalidation import DataValidation
from copy import copy
from collections import defaultdict


DEFAULT_DATA_ROWS = 18  # テンプレ既定の明細行数(5-22)
MAX_TEI = 50             # これ以上は明示的にエラー


def classify_and_aggregate(rows: list[dict]) -> list[dict]:
    """明細行を邸ごとに集計し、D/E/F/G列に振り分ける。"""
    by_tei = defaultdict(lambda: {
        "邸名": "",
        "契約NO": set(),
        "工事名称": set(),
        "D_items": [],
        "E": 0,
        "F": 0,
        "G_items": [],
    })

    for row in rows:
        tei = row["邸名"]
        amount_raw = row["税抜金額"]
        koushu = row["工種"]
        bikou = row.get("備考", "")

        if not tei or tei in ("計", "合計") or "消費税" in tei or "対象外" in tei:
            print(f"  [skip] 集計行: 邸名={tei}")
            continue

        try:
            amount = int(round(float(amount_raw)))
        except (TypeError, ValueError):
            print(f"  [skip] 金額パース失敗: 邸={tei} 金額={amount_raw!r}")
            continue

        agg = by_tei[tei]
        agg["邸名"] = tei
        agg["契約NO"].add(row.get("契約NO", ""))

        base_name = _extract_koji_base(koushu)
        if base_name:
            agg["工事名称"].add(base_name)

        is_seisanka = "生産課中口分" in bikou
        is_shaho = "社保" in koushu

        if amount >= 0:
            agg["D_items"].append(amount)
        else:
            abs_amount = abs(amount)
            if is_seisanka and is_shaho:
                agg["E"] += abs_amount
            elif is_seisanka:
                agg["F"] += abs_amount
            else:
                agg["G_items"].append(abs_amount)

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


def _extract_koji_base(koushu: str) -> Optional[str]:
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
    furikomi_kingaku: Optional[int] = None,
    pdf_koujidai_zeikomi: Optional[int] = None,
    pdf_sousai_zeikomi: Optional[int] = None,
    payment_date: Optional[str] = None,
):
    """集計用テンプレートに書き込む。邸数に応じて動的に行挿入する。"""
    n_tei = len(aggregated)
    if n_tei < 1:
        raise ValueError("抽出された邸数が0です。PDFの内容を確認してください。")
    if n_tei > MAX_TEI:
        raise ValueError(f"邸数が{MAX_TEI}を超えています: {n_tei}邸")

    wb = load_workbook(template_path)
    ws = wb[wb.sheetnames[0]]
    ws.title = sheet_name

    # 全mergeを解除（read-only エラー回避）
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))

    # 合計行を常に 5+n_tei 行に統一（邸数に関わらずデータ行の直下）
    # 元テンプレ配置: data 5-22, spacer 23, 合計 24
    # n_tei=18: 合計を行23に → 行23(スペーサー)削除
    # n_tei<18: 余り行+スペーサー削除
    # n_tei>18: 行挿入+元スペーサー削除
    if n_tei > DEFAULT_DATA_ROWS:
        extra = n_tei - DEFAULT_DATA_ROWS
        ws.insert_rows(23, amount=extra)
        # 行22の書式をコピー
        src_row = 22
        src_height = ws.row_dimensions[src_row].height
        for new_r in range(23, 23 + extra):
            ws.row_dimensions[new_r].height = src_height
            for col_idx in range(1, 15):
                src_cell = ws.cell(row=src_row, column=col_idx)
                dst_cell = ws.cell(row=new_r, column=col_idx)
                if src_cell.has_style:
                    dst_cell.font = copy(src_cell.font)
                    dst_cell.border = copy(src_cell.border)
                    dst_cell.fill = copy(src_cell.fill)
                    dst_cell.number_format = src_cell.number_format
                    dst_cell.alignment = copy(src_cell.alignment)
                    dst_cell.protection = copy(src_cell.protection)
                if col_idx == 10:
                    dst_cell.value = f'=ROUNDDOWN(D{new_r}-E{new_r}-F{new_r}-G{new_r}-H{new_r}-I{new_r},0)'
                elif col_idx == 12:
                    dst_cell.value = f'=J{new_r}/D{new_r}'
        # 元スペーサー(行23 が extra 分下にシフト)を削除
        ws.delete_rows(23 + extra, amount=1)
        print(f"  [insert] {extra}行追加+スペーサー削除 ({n_tei}邸)")
    else:
        # n_tei <= 18 → 余りデータ行+スペーサーを削除して合計を 5+n_tei に詰める
        delete_count = 19 - n_tei  # 最小1(n=18の時)〜最大14(n=5の時)
        if delete_count > 0:
            ws.delete_rows(5 + n_tei, amount=delete_count)
            print(f"  [delete] {delete_count}行削除 ({n_tei}邸)")

    # 行番号ヘルパー（合計行は常に 5 + n_tei）
    data_last_row = 4 + n_tei
    sum_row = 5 + n_tei
    hancho_row_start = sum_row + 5      # 元24→29なので+5
    furikomi_start = sum_row + 13       # 元24→37なので+13
    extra = 0  # 後続のレガシーなオフセット計算を無効化

    # 赤枠再描画（合計行=5+n_tei 基準、元テンプレB29:J35 → B(sum_row+5):J(sum_row+11)）
    red_top = sum_row + 5
    red_bottom = sum_row + 11
    _draw_red_border(ws, top=red_top, bottom=red_bottom, left=2, right=10)

    # タイトル
    ws['C2'] = f'{sheet_name}　着工=受注　ベース'

    # 支払日 (PDFから抽出) を K1 に表示。既存の"YYYY/MM/DD 更新"は上書き
    if payment_date:
        ws['K1'] = f'支払日: {payment_date}'

    # 旧レイアウトの注釈セル(K27:L27相当)をクリア（sum_row + 3）
    note_row = sum_row + 3
    ws.cell(row=note_row, column=11).value = None
    ws.cell(row=note_row, column=12).value = None

    # C列赤塗りクリア（全データ行）
    no_fill = PatternFill(fill_type=None)
    for r in range(5, data_last_row + 1):
        ws.cell(row=r, column=3).fill = no_fill

    # 明細書き込み
    _write_rows(ws, aggregated, data_last_row)

    # 合計行の数式を挿入後の範囲で書き換え
    _rewrite_sum_row(ws, sum_row, data_last_row)

    # 付帯数式（原材料経費合計・生産課支払）: sum_row + 1, sum_row + 2
    i_zairyo_row = sum_row + 1
    e_seisanka_row = sum_row + 2
    ws[f'I{i_zairyo_row}'] = f'=SUM(E5:I{data_last_row})'
    ws[f'E{e_seisanka_row}'] = f'=SUM(E{sum_row}:F{sum_row})'
    # 「生産課 支払 "数値」のセルはMeiryo 20なので E+F マージで表示幅を確保
    try:
        ws.merge_cells(start_row=e_seisanka_row, start_column=5, end_row=e_seisanka_row, end_column=6)
    except Exception:
        pass

    # 担当者別集計(SUMIF) - 挿入後の行/範囲
    data_range_J = f"J5:J{data_last_row}"
    data_range_K = f"K5:K{data_last_row}"
    ws.cell(row=hancho_row_start,     column=12, value=f'=SUMIF({data_range_K},K{hancho_row_start},{data_range_J})')
    ws.cell(row=hancho_row_start + 1, column=12, value=f'=SUMIF({data_range_K},K{hancho_row_start + 1},{data_range_J})')
    ws.cell(row=hancho_row_start + 2, column=12, value=f'=SUMIF({data_range_K},K{hancho_row_start + 2},{data_range_J})')

    # 振込金額照合
    _write_furikomi_verification(
        ws, furikomi_kingaku, pdf_sousai_zeikomi,
        start_row=furikomi_start, sum_row=sum_row,
    )

    # 使いやすさ機能
    _add_usability_features(ws, data_last_row=data_last_row, furikomi_start=furikomi_start)

    try:
        wb.save(output_path)
    finally:
        wb.close()


def _draw_red_border(ws, top: int, bottom: int, left: int, right: int):
    """指定範囲の外周に赤枠(medium)を描画。内側エッジはクリア。"""
    red_side = Side(style='medium', color='FFC00000')
    no_side = Side(style=None)
    for row in range(top, bottom + 1):
        for col_idx in range(left, right + 1):
            is_top = row == top
            is_bottom = row == bottom
            is_left = col_idx == left
            is_right = col_idx == right
            if not (is_top or is_bottom or is_left or is_right):
                continue
            ws.cell(row=row, column=col_idx).border = Border(
                left=red_side if is_left else no_side,
                right=red_side if is_right else no_side,
                top=red_side if is_top else no_side,
                bottom=red_side if is_bottom else no_side,
            )


def _write_rows(ws, aggregated: list[dict], data_last_row: int):
    """明細データをワークシートに書き込む。余ったデータ行はクリア。"""
    for i, item in enumerate(aggregated):
        r = 5 + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=item["邸名"])
        ws.cell(row=r, column=3, value=item["工事名称"])
        d_formula = '=' + '+'.join(str(x) for x in item["D_items"]) if item["D_items"] else 0
        ws.cell(row=r, column=4, value=d_formula)
        ws.cell(row=r, column=5, value=item["E"] if item["E"] else None)
        ws.cell(row=r, column=6, value=item["F"] if item["F"] else None)
        if item["G_items"]:
            ws.cell(row=r, column=7, value='=' + '+'.join(str(x) for x in item["G_items"]))
        else:
            ws.cell(row=r, column=7, value=None)
        ws[f'H{r}'].value = None
        ws[f'I{r}'].value = None
        ws[f'K{r}'].value = None

    # 余った行をクリア（A列の行番号 + B-L列すべて）
    n = len(aggregated)
    for r in range(5 + n, data_last_row + 1):
        for col in range(1, 13):  # A..L
            ws.cell(row=r, column=col).value = None


def _rewrite_sum_row(ws, sum_row: int, data_last_row: int):
    """合計行の数式を新しい範囲で書き換える。"""
    # D〜I: SUM
    for col_letter in ['D', 'E', 'F', 'G', 'H', 'I']:
        ws[f'{col_letter}{sum_row}'] = f'=SUM({col_letter}5:{col_letter}{data_last_row})'
    # J: 粗利合計（プラス連結式で精度確保）
    j_terms = '+'.join(f'J{r}' for r in range(5, data_last_row + 1))
    ws[f'J{sum_row}'] = f'=ROUNDDOWN({j_terms},0)'
    # L: 粗利率
    ws[f'L{sum_row}'] = f'=J{sum_row}/D{sum_row}'


def _write_furikomi_verification(ws, furikomi, sousai, start_row: int, sum_row: int):
    """振込金額照合欄（税抜⇔税込の二重計算）"""
    r_header = start_row
    r_furikomi = start_row + 1
    r_sousai = start_row + 2
    r_zeikomi_total = start_row + 3
    r_zeinuki_calc = start_row + 4
    r_excel_total = start_row + 5
    r_sagaku = start_row + 6
    r_note = start_row + 7

    # 新位置のみクリア（旧位置はテンプレ側で既に削除済み）
    for r in range(r_header, r_note + 1):
        for c in range(2, 6):
            ws.cell(row=r, column=c).value = None

    ws.cell(row=r_header, column=2, value='【振込金額照合（税抜⇔税込の二重計算）】')
    ws.cell(row=r_header, column=2).font = copy(ws['C2'].font)

    ws.cell(row=r_furikomi, column=2, value='① 振込金額(税込)')
    ws.cell(row=r_furikomi, column=4, value=furikomi if furikomi is not None else None)
    ws.cell(row=r_sousai, column=2, value='② 税込相殺(PDF・手入力)')
    ws.cell(row=r_sousai, column=4, value=sousai if sousai is not None else None)
    ws.cell(row=r_zeikomi_total, column=2, value='③ 税込工事代計(① − ②)')
    ws.cell(row=r_zeikomi_total, column=4, value=f'=D{r_furikomi}-D{r_sousai}')
    # ③ 行のboldを解除（テンプレ由来で太字になってるため）
    b_font = ws.cell(row=r_zeikomi_total, column=2).font
    d_font = ws.cell(row=r_zeikomi_total, column=4).font
    ws.cell(row=r_zeikomi_total, column=2).font = Font(
        name=b_font.name, size=b_font.size or 17, bold=False, color=b_font.color
    )
    ws.cell(row=r_zeikomi_total, column=4).font = Font(
        name=d_font.name, size=d_font.size or 17, bold=False, color=d_font.color
    )
    ws.cell(row=r_zeinuki_calc, column=2, value='④ 税抜逆算(③ ÷ 1.1)')
    ws.cell(row=r_zeinuki_calc, column=4, value=f'=ROUND(D{r_zeikomi_total}/1.1,0)')
    ws.cell(row=r_excel_total, column=2, value=f'⑤ Excel税抜合計(J{sum_row})')
    ws.cell(row=r_excel_total, column=4, value=f'=J{sum_row}')
    ws.cell(row=r_sagaku, column=2, value='⑥ 差額(⑤ − ④)')
    ws.cell(row=r_sagaku, column=4, value=f'=D{r_excel_total}-D{r_zeinuki_calc}')
    ws.cell(row=r_note, column=2, value='※ ±数円→インボイス端数差(正常) / 大きな差→PDF読取エラーの可能性')

    for r in [r_furikomi, r_sousai, r_zeikomi_total, r_zeinuki_calc, r_excel_total, r_sagaku]:
        ws.cell(row=r, column=4).number_format = '#,##0;[Red]▲#,##0'

    # ⑥差額 行は他の振込金額照合行と同じサイズ(17)、太字にしない
    ws.cell(row=r_sagaku, column=2).font = Font(
        name=ws.cell(row=r_furikomi, column=2).font.name, size=17, bold=False
    )
    ws.cell(row=r_sagaku, column=4).font = Font(
        name=ws.cell(row=r_furikomi, column=4).font.name, size=17, bold=False
    )


def _add_usability_features(ws, data_last_row: int, furikomi_start: int):
    """プルダウン／条件付き書式／担当邸数カウント"""
    k_range = f"K5:K{data_last_row}"
    b_range = f"B5:B{data_last_row}"

    # 班長プルダウン
    dv = DataValidation(
        type='list', formula1='"山本,熱田,安保"', allow_blank=True,
        showErrorMessage=True, errorTitle='班長名エラー',
        error='山本 / 熱田 / 安保 から選んでください',
    )
    dv.add(k_range)
    ws.add_data_validation(dv)

    # 既存の条件付き書式をリセット（過去ルールが挿入で壊れている可能性があるため）
    ws.conditional_formatting = ConditionalFormattingList()

    # 差額赤/緑 (D{sagaku_row})
    sagaku_row = furikomi_start + 6  # ⑥ 差額 の行
    red_fill = PatternFill(start_color='FFCCCC', end_color='FFCCCC', fill_type='solid')
    green_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
    ws.conditional_formatting.add(
        f'D{sagaku_row}', FormulaRule(formula=[f'ABS(D{sagaku_row})>10'], fill=red_fill)
    )
    ws.conditional_formatting.add(
        f'D{sagaku_row}', FormulaRule(formula=[f'AND(ABS(D{sagaku_row})<=10,D{sagaku_row}<>"")'], fill=green_fill)
    )

    # 班長未入力黄色
    light_yellow = PatternFill(start_color='FFF9E6', end_color='FFF9E6', fill_type='solid')
    ws.conditional_formatting.add(
        k_range, FormulaRule(formula=[f'AND(K5="",B5<>"")'], fill=light_yellow)
    )

    # 班長名による配置と色分け
    _add_hancho_styling(ws, k_range)

    # L列(粗利率)右揃え統一
    for r in range(5, data_last_row + 1):
        ws[f'L{r}'].alignment = Alignment(horizontal='right', vertical='center')

    # 担当邸数カウント(N3:O9) - 固定位置（挿入された行に影響されない想定）
    ws['N3'] = '【担当邸数】'
    ws['N3'].font = copy(ws['C2'].font)
    ws['N4'] = '班長'
    ws['O4'] = '邸数'
    header_font = Font(name=ws['B5'].font.name, bold=True)
    ws['N4'].font = header_font
    ws['O4'].font = header_font
    ws['N5'] = '山本'
    ws['O5'] = f'=COUNTIF({k_range},N5)'
    ws['N6'] = '熱田'
    ws['O6'] = f'=COUNTIF({k_range},N6)'
    ws['N7'] = '安保'
    ws['O7'] = f'=COUNTIF({k_range},N7)'
    ws['N8'] = '未入力'
    ws['O8'] = f'=COUNTBLANK({k_range})-COUNTBLANK({b_range})'
    ws['N9'] = '合計'
    ws['O9'] = '=SUM(O5:O8)'
    ws.conditional_formatting.add('O8', CellIsRule(operator='greaterThan', formula=['0'], fill=light_yellow))
    ws.column_dimensions['N'].width = 14
    ws.column_dimensions['O'].width = 10
    for col in ['D', 'E', 'F', 'G', 'H', 'I', 'J']:
        ws.column_dimensions[col].width = 28


def _add_hancho_styling(ws, k_range: str):
    """班長名に応じて配置と色を自動変更する条件付き書式。"""
    styles = [
        ("山本", "left", "FF006100"),
        ("熱田", "center", "FFC65911"),
        ("安保", "right", "FF2E75B6"),
    ]
    for name, halign, color in styles:
        dxf = DifferentialStyle(
            font=Font(color=color, bold=True),
            alignment=Alignment(horizontal=halign, vertical="center"),
        )
        rule = Rule(type="cellIs", operator="equal", formula=[f'"{name}"'], dxf=dxf)
        ws.conditional_formatting.add(k_range, rule)

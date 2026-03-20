#!/usr/bin/env python3
"""
毎月の経費レポートを自動生成するスクリプト
生成ファイル: expense_report_YYYY_MM.xlsx
"""

import argparse
import calendar
from datetime import date, datetime
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import (
        Alignment, Border, Font, PatternFill, Side
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("openpyxl が見つかりません。インストールしてください:")
    print("  pip install openpyxl")
    raise


# ─── カテゴリ定義 ─────────────────────────────────────────────
CATEGORIES = [
    "交通費",
    "宿泊費",
    "飲食費",
    "消耗品",
    "通信費",
    "接待費",
    "研修・書籍",
    "その他",
]

# ─── スタイル定数 ─────────────────────────────────────────────
COLOR_HEADER_BG   = "2F5496"   # 濃い青
COLOR_HEADER_FONT = "FFFFFF"   # 白
COLOR_TOTAL_BG    = "D9E1F2"   # 薄い青
COLOR_ALT_ROW     = "F2F7FF"   # 交互行の薄い青
COLOR_BORDER      = "BFBFBF"   # グレー

THIN  = Side(style="thin",   color=COLOR_BORDER)
THICK = Side(style="medium", color="2F5496")

BORDER_THIN  = Border(left=THIN,  right=THIN,  top=THIN,  bottom=THIN)
BORDER_OUTER = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)


def _cell_border(top=THIN, bottom=THIN, left=THIN, right=THIN) -> Border:
    return Border(top=top, bottom=bottom, left=left, right=right)


def build_workbook(year: int, month: int, sample: bool = False) -> openpyxl.Workbook:
    """ワークブックを組み立てて返す。"""
    wb = openpyxl.Workbook()

    _build_input_sheet(wb.active, year, month, sample)
    _build_summary_sheet(wb, year, month)
    _build_category_sheet(wb)

    return wb


# ─── 入力シート ───────────────────────────────────────────────
def _build_input_sheet(ws, year: int, month: int, sample: bool):
    ws.title = "経費入力"

    # ── タイトル行 ──
    ws.merge_cells("A1:F1")
    title_cell = ws["A1"]
    title_cell.value = f"{year}年{month}月 経費報告書"
    title_cell.font      = Font(name="游ゴシック", size=16, bold=True, color="2F5496")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # ── 作成日 ──
    ws["E2"] = "作成日:"
    ws["E2"].font      = Font(bold=True)
    ws["E2"].alignment = Alignment(horizontal="right")
    ws["F2"] = date.today()
    ws["F2"].number_format = "YYYY/MM/DD"

    # ── ヘッダー行 ──
    headers = ["No.", "日付", "金額（円）", "カテゴリ", "備考", "承認"]
    col_widths = [5, 13, 14, 14, 35, 8]

    for col, (h, w) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=4, column=col, value=h)
        cell.font      = Font(name="游ゴシック", bold=True, color=COLOR_HEADER_FONT, size=11)
        cell.fill      = PatternFill("solid", fgColor=COLOR_HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = _cell_border()
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.row_dimensions[4].height = 22

    # ── データ行 ──
    data_start = 5
    num_rows   = 30  # データ入力行数

    sample_data = _sample_data(year, month) if sample else []

    for i in range(num_rows):
        row = data_start + i
        fill = PatternFill("solid", fgColor=COLOR_ALT_ROW) if i % 2 == 0 else None

        # No.
        no_cell = ws.cell(row=row, column=1, value=i + 1)
        no_cell.alignment = Alignment(horizontal="center")
        no_cell.border    = BORDER_THIN
        if fill:
            no_cell.fill = fill

        # 日付
        date_cell = ws.cell(row=row, column=2)
        date_cell.number_format = "YYYY/MM/DD"
        date_cell.alignment     = Alignment(horizontal="center")
        date_cell.border        = BORDER_THIN
        if fill:
            date_cell.fill = fill

        # 金額
        amt_cell = ws.cell(row=row, column=3)
        amt_cell.number_format = "#,##0"
        amt_cell.alignment     = Alignment(horizontal="right")
        amt_cell.border        = BORDER_THIN
        if fill:
            amt_cell.fill = fill

        # カテゴリ（ドロップダウン）
        cat_cell = ws.cell(row=row, column=4)
        cat_cell.alignment = Alignment(horizontal="center")
        cat_cell.border    = BORDER_THIN
        if fill:
            cat_cell.fill = fill

        # 備考
        rem_cell = ws.cell(row=row, column=5)
        rem_cell.alignment = Alignment(horizontal="left")
        rem_cell.border    = BORDER_THIN
        if fill:
            rem_cell.fill = fill

        # 承認チェック欄
        apv_cell = ws.cell(row=row, column=6)
        apv_cell.alignment = Alignment(horizontal="center")
        apv_cell.border    = BORDER_THIN
        if fill:
            apv_cell.fill = fill

        # サンプルデータを埋め込む
        if i < len(sample_data):
            d = sample_data[i]
            date_cell.value = d["date"]
            amt_cell.value  = d["amount"]
            cat_cell.value  = d["category"]
            rem_cell.value  = d["note"]

    # カテゴリのドロップダウンバリデーション
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(
        type="list",
        formula1='"' + ",".join(CATEGORIES) + '"',
        allow_blank=True,
        showDropDown=False,
    )
    dv.sqref = f"D{data_start}:D{data_start + num_rows - 1}"
    ws.add_data_validation(dv)

    # ── 合計行 ──
    total_row = data_start + num_rows
    ws.merge_cells(f"A{total_row}:B{total_row}")
    total_label = ws[f"A{total_row}"]
    total_label.value     = "合計"
    total_label.font      = Font(bold=True, color="2F5496")
    total_label.alignment = Alignment(horizontal="center")
    total_label.fill      = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
    total_label.border    = _cell_border(top=THICK)

    total_val = ws.cell(row=total_row, column=3)
    total_val.value        = f"=SUM(C{data_start}:C{data_start + num_rows - 1})"
    total_val.number_format = "#,##0"
    total_val.font         = Font(bold=True, color="2F5496")
    total_val.alignment    = Alignment(horizontal="right")
    total_val.fill         = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
    total_val.border       = _cell_border(top=THICK)

    for col in range(4, 7):
        c = ws.cell(row=total_row, column=col)
        c.fill   = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
        c.border = _cell_border(top=THICK)

    # ── 件数行 ──
    count_row = total_row + 1
    ws.merge_cells(f"A{count_row}:B{count_row}")
    ws[f"A{count_row}"].value     = "件数"
    ws[f"A{count_row}"].font      = Font(bold=True)
    ws[f"A{count_row}"].alignment = Alignment(horizontal="center")

    cnt_cell = ws.cell(row=count_row, column=3)
    cnt_cell.value        = f"=COUNTA(C{data_start}:C{data_start + num_rows - 1})"
    cnt_cell.number_format = "#,##0"
    cnt_cell.alignment    = Alignment(horizontal="right")

    # ── 印刷設定 ──
    ws.print_title_rows = "4:4"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToPage   = True
    ws.page_setup.fitToWidth  = 1
    ws.sheet_view.showGridLines = True


# ─── 集計シート ───────────────────────────────────────────────
def _build_summary_sheet(wb: openpyxl.Workbook, year: int, month: int):
    ws = wb.create_sheet("カテゴリ集計")

    ws.merge_cells("A1:C1")
    ws["A1"].value     = f"{year}年{month}月 カテゴリ別集計"
    ws["A1"].font      = Font(bold=True, size=13, color="2F5496")
    ws["A1"].alignment = Alignment(horizontal="center")

    headers = ["カテゴリ", "金額（円）", "割合（%）"]
    widths  = [16, 16, 14]
    for col, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font      = Font(bold=True, color=COLOR_HEADER_FONT)
        cell.fill      = PatternFill("solid", fgColor=COLOR_HEADER_BG)
        cell.alignment = Alignment(horizontal="center")
        cell.border    = BORDER_THIN
        ws.column_dimensions[get_column_letter(col)].width = w

    input_ws = "経費入力"
    data_start = 5
    num_rows   = 30

    for i, cat in enumerate(CATEGORIES):
        row  = 4 + i
        fill = PatternFill("solid", fgColor=COLOR_ALT_ROW) if i % 2 == 0 else None

        cat_cell = ws.cell(row=row, column=1, value=cat)
        cat_cell.alignment = Alignment(horizontal="center")
        cat_cell.border    = BORDER_THIN
        if fill:
            cat_cell.fill = fill

        # SUMIF で経費入力シートを参照
        amt_cell = ws.cell(row=row, column=2)
        amt_cell.value = (
            f"=SUMIF('{input_ws}'!D{data_start}:D{data_start + num_rows - 1},"
            f'"{cat}",'
            f"'{input_ws}'!C{data_start}:C{data_start + num_rows - 1})"
        )
        amt_cell.number_format = "#,##0"
        amt_cell.alignment     = Alignment(horizontal="right")
        amt_cell.border        = BORDER_THIN
        if fill:
            amt_cell.fill = fill

        pct_cell = ws.cell(row=row, column=3)
        pct_cell.value = (
            f"=IF(B{4 + len(CATEGORIES) + 1}=0,0,"
            f"B{row}/B{4 + len(CATEGORIES) + 1})"
        )
        pct_cell.number_format = "0.0%"
        pct_cell.alignment     = Alignment(horizontal="right")
        pct_cell.border        = BORDER_THIN
        if fill:
            pct_cell.fill = fill

    # 合計行
    total_row = 4 + len(CATEGORIES)
    ws.cell(row=total_row, column=1, value="合計").font      = Font(bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="center")
    ws.cell(row=total_row, column=1).fill      = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
    ws.cell(row=total_row, column=1).border    = _cell_border(top=THICK)

    total_amt = ws.cell(row=total_row, column=2)
    total_amt.value        = f"=SUM(B4:B{total_row - 1})"
    total_amt.number_format = "#,##0"
    total_amt.font         = Font(bold=True)
    total_amt.alignment    = Alignment(horizontal="right")
    total_amt.fill         = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
    total_amt.border       = _cell_border(top=THICK)

    ws.cell(row=total_row, column=3).fill   = PatternFill("solid", fgColor=COLOR_TOTAL_BG)
    ws.cell(row=total_row, column=3).border = _cell_border(top=THICK)


# ─── カテゴリ一覧シート ───────────────────────────────────────
def _build_category_sheet(wb: openpyxl.Workbook):
    ws = wb.create_sheet("カテゴリ一覧")
    ws["A1"].value = "カテゴリ"
    ws["A1"].font  = Font(bold=True)
    for i, cat in enumerate(CATEGORIES, start=2):
        ws.cell(row=i, column=1, value=cat)
    ws.column_dimensions["A"].width = 16
    ws.sheet_state = "hidden"  # 補助シートなので非表示


# ─── サンプルデータ ───────────────────────────────────────────
def _sample_data(year: int, month: int) -> list[dict]:
    _, last_day = calendar.monthrange(year, month)
    return [
        {"date": date(year, month, 5),  "amount": 1500,  "category": "交通費", "note": "東京→大阪 新幹線"},
        {"date": date(year, month, 7),  "amount": 3200,  "category": "飲食費", "note": "チーム昼食"},
        {"date": date(year, month, 10), "amount": 5500,  "category": "接待費", "note": "取引先 ○○様"},
        {"date": date(year, month, 12), "amount": 2800,  "category": "消耗品", "note": "文具・用紙"},
        {"date": date(year, month, 15), "amount": 12000, "category": "宿泊費", "note": "出張 ビジネスホテル"},
        {"date": date(year, month, 18), "amount": 980,   "category": "通信費", "note": "Wi-Fi"},
        {"date": date(year, month, 20), "amount": 3500,  "category": "研修・書籍", "note": "技術書籍"},
        {"date": date(year, month, last_day), "amount": 800, "category": "その他", "note": ""},
    ]


# ─── メイン ───────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="月次経費レポート Excel 自動生成")
    parser.add_argument("--year",   type=int, default=date.today().year,  help="対象年 (デフォルト: 今年)")
    parser.add_argument("--month",  type=int, default=date.today().month, help="対象月 (デフォルト: 今月)")
    parser.add_argument("--output", type=str, default="",                 help="出力ディレクトリ (デフォルト: カレントディレクトリ)")
    parser.add_argument("--sample", action="store_true",                  help="サンプルデータを挿入する")
    args = parser.parse_args()

    if not (1 <= args.month <= 12):
        parser.error("--month は 1〜12 の範囲で指定してください。")

    out_dir = Path(args.output) if args.output else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = out_dir / f"expense_report_{args.year}_{args.month:02d}.xlsx"

    print(f"生成中: {filename}")
    wb = build_workbook(args.year, args.month, sample=args.sample)
    wb.save(filename)
    print(f"完了: {filename}")
    print()
    print("シート構成:")
    print("  ① 経費入力     … 日付・金額・カテゴリ・備考を入力するメインシート")
    print("  ② カテゴリ集計 … カテゴリ別の合計・割合を自動集計")
    print("  ③ カテゴリ一覧 … ドロップダウン用補助シート（非表示）")


if __name__ == "__main__":
    main()

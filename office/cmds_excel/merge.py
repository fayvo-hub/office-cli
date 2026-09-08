"""merge — 合并 / 取消合并单元格。"""

from __future__ import annotations

import argparse

from openpyxl.styles import Alignment
from openpyxl.worksheet.cell_range import CellRange

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "merge"
HELP = "合并 / 取消合并单元格区域"
DESCRIPTION = """合并 / 取消合并单元格区域。

用法示例:
  office merge -f demo.xlsx --range A1:B2                # 合并
  office merge -f demo.xlsx --range A1:B2 --center       # 合并并让内容水平垂直居中
  office merge -f demo.xlsx --range A1:B2 --unmerge      # 取消合并

注意:
- 合并后数据只保留左上角格:openpyxl 执行合并时会自动清除区域内其它格的值,
  因此默认会在执行前警告并列出将被清除的单元格(--discard-others 可跳过警告)
- 新合并区不能与现有合并区重叠
输出 JSON: {"ok": true, "file": "...", "sheet": "...", "action": "merge|unmerge",
            "range": "A1:B2", "warnings": [...], "merged_cells": [...]}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    sp.add_argument("--range", required=True, metavar="REF", help="区域,如 A1:B2")
    sp.add_argument("--unmerge", action="store_true", help="取消合并而不是合并")
    sp.add_argument("--center", action="store_true",
                    help="合并后把左上角内容设为水平垂直居中(纯视觉,不改数据)")
    sp.add_argument("--discard-others", action="store_true",
                    help="确认丢弃非左上角的值,不输出警告(openpyxl 合并时本来就会清除这些值)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)
    new_range = CellRange(xlutil.area_label(r1, c1, r2, c2))
    warnings: list[str] = []

    if args.unmerge:
        existing = {str(r).upper(): r for r in ws.merged_cells.ranges}
        key = str(new_range).upper()
        if key not in existing:
            raise CliError("not_merged", f"区域 {args.range} 不是现有合并区,无法取消合并。"
                                         f"现有合并区: {sorted(existing)}")
        ws.unmerge_cells(str(existing[key]))
        action = "unmerge"
    else:
        for rng in ws.merged_cells.ranges:
            if _intersect(rng, new_range):
                raise CliError("overlap", f"新合并区 {args.range} 与现有合并区 {rng} 重叠;"
                                          f"请先取消原合并(--unmerge)")
        # 非左上格有值 -> 合并前先警告(openpyxl 合并会清除这些值)
        lost = []
        for row in xlutil.iter_area(ws, r1, c1, r2, c2):
            for cell in row:
                if (cell.row, cell.column) == (r1, c1):
                    continue
                if cell.value is not None:
                    lost.append(cell.coordinate)
        if lost and not args.discard_others:
            warnings.append(
                f"合并 {new_range} 时以下格的值会被清除(仅保留左上角): {lost}。"
                f"如需保留请先 read 出来备份;确认丢弃可加 --discard-others 跳过本警告")

        ws.merge_cells(str(new_range))
        action = "merge"

    if args.center and not args.unmerge:
        top_left = ws.cell(row=r1, column=c1)
        old = top_left.alignment
        top_left.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=old.wrap_text)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    merged_now = [str(r) for r in ws.merged_cells.ranges]
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title, "action": action,
            "range": str(new_range), "warnings": warnings, "merged_cells": merged_now}


def _intersect(a: CellRange, b: CellRange) -> bool:
    return not (a.max_row < b.min_row or b.max_row < a.min_row or
                a.max_col < b.min_col or b.max_col < a.min_col)

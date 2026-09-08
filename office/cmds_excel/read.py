"""read — 读取单元格/区域,输出 JSON。"""

from __future__ import annotations

import argparse
from typing import Optional

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "read"
HELP = "读取单元格/区域/整表,输出 JSON(默认行数组,公式格显示公式文本)"
DESCRIPTION = """读取单元格/区域/整表并输出 JSON。

用法示例:
  office read -f demo.xlsx                        # 读激活表全部数据
  office read -f demo.xlsx --sheet 销售 --range A1:F13
  office read -f demo.xlsx --range B2             # 单个单元格
  office read -f demo.xlsx --range A1:C10 --cached   # 公式格显示计算结果(需文件有缓存)

输出结构(rows 模式):
{
  "file": "...", "sheet": "Sheet1", "range": "A1:F13",
  "row_count": 13, "col_count": 6,
  "rows": [ [ "月份", 100, 1.5, ... ], ... ],   // 每行与列对齐;空表为 []
  "merged_cells": ["A15:B15"],                   // 合并区列表(数据只在左上角格)
  "warnings": []                                 // 截断/缓存缺失等提示
}

值编码约定:
- 日期/时间 -> ISO 字符串,如 "2024-03-05T10:00:00"、"2024-03-05"
- 公式格默认输出公式文本(如 "=B2*C2");加 --cached 则输出计算缓存值
- 错误值(如 #DIV/0!)以字符串原样输出,类型见 --cells 模式
- 合并区只在左上角格有值,其余格为 null —— 要完整铺开请自行按 merged_cells 填充

--cells 模式(逐格输出类型与格式元数据):
{"ref": "B2", "row": 2, "col": 2, "type": "number|str|formula|bool|date|error|blank",
 "value": ..., "number_format": "General", "formula": "=B2*C2"?, ...}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    sp.add_argument("--range", metavar="REF",
                    help="要读取的范围:B2 / A1:B3 / A:C(整列) / 1:3(整行);缺省读全表数据区")
    sp.add_argument("--cells", action="store_true",
                    help="逐格输出(含类型/number_format/公式文本等元数据)")
    sp.add_argument("--cached", action="store_true",
                    help="公式格输出计算缓存值而不是公式文本"
                        "(缓存是文件保存时留下的;openpyxl 等程序写出的文件可能无缓存)")
    sp.add_argument("--limit", type=int, default=10000, metavar="N",
                    help="行数上限保护,超出截断并警告(默认 10000;设 0 不限制)")
    sp.add_argument("--drop-empty-rows", action="store_true",
                    help="过滤整行为空的记录(默认保留,保持行列对齐)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file, write=False)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)

    if args.range:
        r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)
    else:
        dims = xlutil.data_dimensions(ws)
        if dims is None:
            return _empty_result(args, ws, [])
        r1, c1, r2, c2 = dims

    warnings: list[str] = []

    # 行数上限保护
    if args.limit and args.limit > 0 and (r2 - r1 + 1) > args.limit:
        r2 = r1 + args.limit - 1
        warnings.append(
            f"该区域共 {c2 - c1 + 1} 列 × 超过 {args.limit} 行,已按 --limit {args.limit} 截断;"
            f"如需全部数据请加 --limit 0 或缩小 --range"
        )

    # 需要缓存值时二次以 data_only 加载(坐标对齐取缓存)
    wb_cached = xlutil.open_workbook(plan.read_path, data_only=True) if args.cached else None
    ws_cached = wb_cached[ws.title] if wb_cached is not None else None

    merged = [str(r) for r in ws.merged_cells.ranges]

    if args.cells:
        cells = []
        for row in xlutil.iter_area(ws, r1, c1, r2, c2):
            for cell in row:
                cells.append(_cell_dict(cell, ws_cached))
        result: dict = {
            "ok": True, "file": args.file, "sheet": ws.title,
            "range": xlutil.area_label(r1, c1, r2, c2),
            "mode": "cells", "cell_count": len(cells),
            "cells": cells, "merged_cells": merged,
            "warnings": warnings,
        }
    else:
        rows = []
        empty_hits = 0
        for row in xlutil.iter_area(ws, r1, c1, r2, c2):
            vals = []
            row_empty = True
            for cell in row:
                if cell.data_type == "f":
                    v: Optional[str] = cell.value
                    if ws_cached is not None:
                        cv = ws_cached.cell(row=cell.row, column=cell.column).value
                        v = xlutil.serialize_value(cv)
                        if cv is not None:
                            row_empty = False
                else:
                    v = xlutil.serialize_value(cell.value)
                    if v is not None and v != "":
                        row_empty = False
                vals.append(v)
            if args.drop_empty_rows and row_empty:
                empty_hits += 1
                continue
            rows.append(vals)
        result = {
            "ok": True, "file": args.file, "sheet": ws.title,
            "range": xlutil.area_label(r1, c1, r2, c2),
            "mode": "rows", "row_count": len(rows), "col_count": c2 - c1 + 1,
            "rows": rows, "merged_cells": merged,
            "warnings": warnings,
        }
        if empty_hits:
            result["dropped_empty_rows"] = empty_hits

    return result


def _cell_dict(cell, ws_cached=None) -> dict:
    d = {
        "ref": cell.coordinate,
        "row": cell.row,
        "col": cell.column,
        "type": xlutil.cell_type(cell),
        "number_format": cell.number_format,
    }
    if cell.data_type == "f":
        d["formula"] = cell.value
        if ws_cached is not None:
            cv = ws_cached.cell(row=cell.row, column=cell.column).value
            d["value"] = xlutil.serialize_value(cv)
            d["cached_available"] = cv is not None
        else:
            d["value"] = None
    else:
        d["value"] = xlutil.serialize_value(cell.value)

    link = getattr(cell, "hyperlink", None)
    if link is not None and getattr(link, "target", None):
        d["hyperlink"] = link.target
    comment = getattr(cell, "comment", None)
    if comment is not None and getattr(comment, "text", None):
        d["comment"] = comment.text
    return d


def _empty_result(args, ws, warnings) -> dict:
    return {
        "ok": True, "file": args.file, "sheet": ws.title,
        "range": "A1", "mode": "rows",
        "row_count": 0, "col_count": 0,
        "rows": [], "merged_cells": [],
        "warnings": ["该表无任何数据"] + warnings,
    }

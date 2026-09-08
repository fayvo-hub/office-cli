"""list — 列出工作簿结构:所有工作表、尺寸、合并区、表头预览。"""

from __future__ import annotations

import argparse

from .. import ioplan, xlutil
from ..errors import CliError
from ..cli import add_file_arg

NAME = "list"
HELP = "列出工作簿的全部工作表与结构信息(含前 N 行预览)"
DESCRIPTION = """列出工作簿的全部工作表与结构信息。

输出 JSON 结构:
{
  "file": "...",
  "sheets": [
    {
      "name": "Sheet1", "index": 0, "state": "visible",
      "max_row": 12, "max_column": 5, "dimensions": "A1:E13",
      "merged_cells": ["A1:B1"],
      "tables": [],
      "charts": 0, "images": 0,
      "preview_rows": [ [第一行的值], [第二行的值], ... ]   // 默认前 3 行
    }
  ]
}

说明:
- 公式单元格在预览中显示公式文本(如 "=B2*C2")
- 想拿全部数据请用 read 命令
- 常见用途:AI 先 list 了解表结构,再决定后续 read/write 参数
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    sp.add_argument("--preview", type=int, default=3, metavar="N",
                    help="每个表预览前 N 行(默认 3,设 0 关闭)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file, write=False)
    wb = xlutil.open_workbook(plan.read_path)
    sheets = []
    for idx, ws in enumerate(wb.worksheets):
        info = _sheet_info(ws, idx, args.preview)
        sheets.append(info)
    return {"ok": True, "file": args.file, "sheet_count": len(sheets), "sheets": sheets}


def _sheet_info(ws, index: int, preview: int) -> dict:
    dims = xlutil.data_dimensions(ws)
    if dims is None:
        max_row = max_col = 0
        dimensions = "A1"
    else:
        _, _, max_row, max_col = dims
        dimensions = xlutil.area_label(*dims)

    merged = [str(r) for r in ws.merged_cells.ranges]
    tables = list(getattr(ws, "tables", {}).keys())
    charts = len(getattr(ws, "_charts", []))
    images = len(getattr(ws, "_images", []))

    info = {
        "name": ws.title,
        "index": index,
        "state": ws.sheet_state,
        "max_row": max_row,
        "max_column": max_col,
        "dimensions": dimensions,
        "merged_cells": merged,
        "tables": tables,
        "charts": charts,
        "images": images,
    }

    if preview > 0 and dims is not None:
        r1, c1, r2, c2 = dims
        r2 = min(r2, r1 + preview - 1)
        rows = []
        for row in ws.iter_rows(min_row=r1, min_col=c1, max_row=r2, max_col=c2):
            vals = []
            for cell in row:
                if cell.data_type == "f":
                    vals.append(cell.value)  # 公式文本
                else:
                    vals.append(xlutil.serialize_value(cell.value))
            rows.append(vals)
        info["preview_rows"] = rows

    return info

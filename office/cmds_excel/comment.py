"""comment — 单元格批注(注释)设置/清除。"""

from __future__ import annotations

import argparse

from openpyxl.comments import Comment

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "comment"
HELP = "单元格批注:set 添加/覆盖 / clear 删除"
DESCRIPTION = """单元格批注(Excel 中右上角红三角,悬停显示)。

用法示例:
  office comment set -f f.xlsx --cell B2 --text '此处需人工复核'
  office comment set -f f.xlsx --cell B2 --text '说明' --author 张三
  office comment clear -f f.xlsx --cell B2

说明:
- 同一格重复 set 会覆盖旧批注
- 合并区只能给左上角格加批注,其余格报 merged_cell
输出 JSON: {"ok": true, "file": "...", "cell": "B2", "comment": "..."|null}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="cm_action", metavar="动作", required=True)

    s = sub.add_parser("set", help="添加/覆盖批注")
    add_file_arg(s)
    add_sheet_arg(s)
    s.add_argument("--cell", required=True, metavar="REF", help="单元格,如 B2")
    s.add_argument("--text", required=True, metavar="TEXT", help="批注内容")
    s.add_argument("--author", metavar="NAME", default="office-cli",
                   help="作者名(默认 office-cli)")
    s.set_defaults(cm_func="set")

    s2 = sub.add_parser("clear", help="删除批注")
    add_file_arg(s2)
    add_sheet_arg(s2)
    s2.add_argument("--cell", required=True, metavar="REF", help="单元格,如 B2")
    s2.set_defaults(cm_func="clear")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    r1, c1, r2, c2 = xlutil.parse_range(args.cell, ws)
    if not (r1 == r2 and c1 == c2):
        raise CliError("bad_args", f"--cell 需单个单元格,收到区域 {args.cell}")
    anchor = xlutil.merged_anchor(ws, r1, c1)
    if anchor is not None and (anchor[0], anchor[1]) != (r1, c1):
        raise CliError("merged_cell",
                       f"{args.cell} 位于合并区 {xlutil.area_label(*anchor)} 内且不是左上角;"
                       f"批注只能加在左上角格")
    cell = ws.cell(row=r1, column=c1)

    if args.cm_func == "set":
        cell.comment = Comment(args.text, args.author)
        result: str | None = args.text
    else:
        cell.comment = None
        result = None

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "cell": args.cell.upper(), "comment": result}

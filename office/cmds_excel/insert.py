"""insert — 在工作表插入整行/整列(其后数据自动下移/右移)。"""

from __future__ import annotations

import argparse

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "insert"
HELP = "插入整行/整列(--rows 3 或 --cols C,可 --count)"
DESCRIPTION = """在指定位置插入整行或整列,其后内容自动移位(Excel 打开后同效果)。

用法示例:
  office insert -f f.xlsx --rows 3 --count 2    # 在第 3 行前插入 2 个空行
  office insert -f f.xlsx --cols B --count 1    # 在 B 列前插入 1 个空列
  office insert -f f.xlsx --sheet 销售 --rows 10

说明与限制:
- 公式引用不会随插入更新(openpyxl 无重算引擎),涉及公式请人工核对
- 若插入位置会错位已有合并单元格/表格对象/自动筛选区域,会先报错拒绝,
  请先用 merge --unmerge / layout --unfilter 清理后再插入
- 操作不可撤销,建议先备份
输出 JSON: {"ok": true, "file": "...", "inserted": "rows|cols", "at": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    g = sp.add_mutually_exclusive_group(required=True)
    g.add_argument("--rows", metavar="N",
                   help="在该行之前插入(如 3 = 新行成为第 3 行)")
    g.add_argument("--cols", metavar="COL",
                   help="在该列之前插入(如 C = 新列成为 C 列)")
    sp.add_argument("--count", type=int, default=1, metavar="N",
                    help="插入数量(默认 1)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    axis, at, label = _axis(args)

    xlutil.guard_rowcol_shift(ws, axis, at, "插入")
    if axis == "rows":
        ws.insert_rows(at, args.count)
    else:
        ws.insert_cols(at, args.count)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "inserted": axis, "at": label, "count": args.count,
            "warnings": ["公式引用不会随插入自动更新,请核对涉及公式的区域"]}


# ---------------------------------------------------------------------------

def _axis(args) -> tuple[str, int, str]:
    if args.rows is not None:
        n = int(args.rows)
        if not 1 <= n <= xlutil.MAX_ROWS:
            raise CliError("bad_args", f"--rows 需在 1..{xlutil.MAX_ROWS},收到 {n}")
        return "rows", n, str(n)
    text = args.cols.strip().upper()
    c = xlutil.col_to_idx(text)
    if not 1 <= c <= xlutil.MAX_COLS:
        raise CliError("bad_args", f"--cols 超出范围: {text}")
    return "cols", c, text

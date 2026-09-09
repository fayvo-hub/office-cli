"""validate — 数据有效性(下拉列表)设置/清除。"""

from __future__ import annotations

import argparse

from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.cell import range_boundaries

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "validate"
HELP = "数据验证:add 下拉列表 / clear 移除"
DESCRIPTION = """单元格下拉列表(数据验证)。

用法示例:
  office validate add -f f.xlsx --range A2:A100 --list '男,女'
  office validate add -f f.xlsx --range B2:B100 --source '选项表!A1:A5'
  office validate clear -f f.xlsx --range A2:A100

说明:
- --list 内联枚举(逗号分隔,选项本身含逗号时用 --source 更稳)
- --source 引用其他区域的选项,如 '选项表!A1:A5'(sheet 名含空格需加引号)
- 默认允许留空(--allow-blank 关闭可禁止空值)
- 不匹配选项时 Excel 会拒绝输入并提示
输出 JSON: {"ok": true, "file": "...", "validations": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="dv_action", metavar="动作", required=True)

    s = sub.add_parser("add", help="添加下拉列表验证")
    add_file_arg(s)
    add_sheet_arg(s)
    s.add_argument("--range", required=True, metavar="REF",
                   help="应用区域,如 A2:A100(通常不含表头行)")
    g = s.add_mutually_exclusive_group(required=True)
    g.add_argument("--list", metavar="ITEMS", help="选项列表,逗号分隔,如 '男,女'")
    g.add_argument("--source", metavar="REF",
                   help="选项所在区域,如 '选项表!A1:A5'")
    s.add_argument("--allow-blank", action="store_true",
                   help="允许留空(默认已允许;此参数仅用于显式说明)")
    s.set_defaults(dv_func="add")

    s2 = sub.add_parser("clear", help="移除与区域相交的下拉验证")
    add_file_arg(s2)
    add_sheet_arg(s2)
    s2.add_argument("--range", required=True, metavar="REF", help="应用区域")
    s2.set_defaults(dv_func="clear")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)

    if args.dv_func == "add":
        if args.list is not None:
            formula = f'"{args.list}"'
        else:
            src = args.source.strip()
            if src.startswith("="):
                formula = src
            else:
                formula = "=" + src
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.error = "请从下拉列表中选择"
        dv.errorTitle = "输入无效"
        ws.add_data_validation(dv)
        dv.add(xlutil.area_label(r1, c1, r2, c2))
    else:
        # 移除与清除区(行列区间)相交的所有验证
        box = (c1, r1, c2, r2)  # range_boundaries 顺序: 列,行,列,行
        for d in list(ws.data_validations.dataValidation):
            overlap = False
            for r in d.sqref.ranges:
                mc, mr, Mc, Mr = range_boundaries(str(r))
                if not (Mc < box[0] or Mr < box[1]
                        or mc > box[2] or mr > box[3]):
                    overlap = True
                    break
            if overlap:
                ws.data_validations.dataValidation.remove(d)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    n = len(ws.data_validations.dataValidation)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "validations": n}

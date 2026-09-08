"""sheet — 工作表增删改(不触碰单元格内容)。"""

from __future__ import annotations

import argparse

from .. import ioplan, xlutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "sheet"
HELP = "工作表管理:add / rename / remove / copy / active"
DESCRIPTION = """工作表管理(不触碰单元格内容)。

子命令:
  office sheet add    -f f.xlsx --name 新表 [--index 1]
  office sheet rename -f f.xlsx --name 旧名 --new-name 新名
  office sheet remove -f f.xlsx --name 表名        (最后一个表不可删除)
  office sheet copy   -f f.xlsx --name 源表 --new-name 副本名
  office sheet active -f f.xlsx --name 表名        (设置打开时默认显示的表)

输出 JSON: {"ok": true, "file": "...", "action": "...", "sheets": ["当前全部表名"]}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="sheet_action", metavar="动作", required=True)
    for act in ("add", "rename", "remove", "copy", "active"):
        s = sub.add_parser(act, help=_ACT_HELP[act])
        add_file_arg(s)
        if act in ("rename", "remove", "copy", "active"):
            s.add_argument("--name", required=True, metavar="OLD",
                           help="现有工作表名")
        if act == "add":
            s.add_argument("--name", required=True, metavar="NAME",
                           help="新工作表名(不能与现有表重名)")
            s.add_argument("--index", type=int, default=None, metavar="N",
                           help="插入位置,0 起(默认追加到末尾)")
        if act in ("rename", "copy"):
            s.add_argument("--new-name", required=True, metavar="NEW",
                           help="新表名(不能与现有表重名)")
        s.set_defaults(sheet_func=act)


_ACT_HELP = {
    "add": "新建工作表",
    "rename": "重命名工作表",
    "remove": "删除工作表(内容一并删除,不可恢复)",
    "copy": "复制工作表(含内容与样式)",
    "active": "设为打开文件时默认显示的表",
}


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    action = args.sheet_func

    if action == "add":
        _check_new_name(wb, args.name)
        wb.create_sheet(title=args.name, index=args.index)
    elif action == "rename":
        ws = xlutil.choose_sheet(wb, args.name)
        _check_new_name(wb, args.new_name)
        ws.title = args.new_name
    elif action == "remove":
        ws = xlutil.choose_sheet(wb, args.name)
        if len(wb.sheetnames) <= 1:
            raise CliError("last_sheet", f"不能删除最后一个工作表 '{args.name}'"
                                         f"(Excel 工作簿至少保留一个表)")
        del wb[ws.title]
    elif action == "copy":
        ws = xlutil.choose_sheet(wb, args.name)
        _check_new_name(wb, args.new_name)
        wb.copy_worksheet(ws).title = args.new_name
    elif action == "active":
        ws = xlutil.choose_sheet(wb, args.name)
        wb.active = wb.sheetnames.index(ws.title)
    else:  # pragma: no cover
        raise CliError("internal", f"未知动作 {action}")

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "action": action,
            "sheets": xlutil.sheet_names_snapshot(wb)}


def _check_new_name(wb, name: str) -> None:
    name = name.strip()
    if not name:
        raise CliError("bad_name", "表名不能为空")
    if name in wb.sheetnames:
        raise CliError("dup_sheet",
                       f"表名 '{name}' 已存在;Excel 表名必须唯一。现有表: {wb.sheetnames}")

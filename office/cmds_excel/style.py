"""style — 给区域设置字体/填充/对齐/边框/数字格式。"""

from __future__ import annotations

import argparse

from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "style"
HELP = "设置区域样式:字体/填充/对齐/边框/数字格式"
DESCRIPTION = """给指定区域设置样式(只改样式,不动数据)。未给出的选项保持原样。

用法示例:
  xcli style -f demo.xlsx --range A1:F1 --bold --fill FFFF00 --align center
  xcli style -f demo.xlsx --sheet 销售 --range A2:A13 --num-format '0.00%'
  xcli style -f demo.xlsx --range A1:F13 --border thin --border-color 999999

选项:
  字体:   --font-name 微软雅黑 | --font-size 12 | --bold [true] | --italic | --underline
          --font-color FF0000 (6 位 hex)
  填充:   --fill FFFF00 (6 位 hex 或 none 清除填充)
  对齐:   --align left|center|right | --valign top|center|bottom | --wrap [true]
  边框:   --border thin|medium|thick|dashed|dotted|double|hair|none
          --border-color 333333(需与 --border 一起用)
  数字:   --num-format '0.00' | 'yyyy-mm-dd' | '#,##0' | 'general'(恢复常规)

说明:
- 布尔选项可写 --bold 或 --bold true / --bold false
- 区域最大 50000 格(openpyxl 会为区域内每格创建样式对象,过大会拖垮文件)
- 合并区域请直接对该区域操作(样式会应用到所有格)
输出 JSON: {"ok": true, "file": "...", "sheet": "...", "range": "...",
            "cells_affected": 12, "applied": ["bold", "fill", ...]}
"""

MAX_STYLE_CELLS = 50_000

_FONT_SIZES = (8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48, 72)
_ALIGNS = ("left", "center", "right", "fill", "justify", "centerContinuous", "distributed")
_VALIGNS = ("top", "center", "bottom", "justify", "distributed")
_BORDERS = ("thin", "medium", "thick", "dashed", "dotted", "double", "hair", "none")


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    sp.add_argument("--range", required=True, metavar="REF",
                    help="样式作用区域,如 A1:F1(必填)")
    sp.add_argument("--font-name", metavar="NAME", help="字体名,如 微软雅黑 / Calibri")
    sp.add_argument("--font-size", type=float, metavar="PT",
                    help=f"字号(常用: {', '.join(str(s) for s in _FONT_SIZES)})")
    sp.add_argument("--bold", nargs="?", const="true", metavar="BOOL", type=_boolish,
                    help="粗体(可写 --bold 或 --bold true/false)")
    sp.add_argument("--italic", nargs="?", const="true", metavar="BOOL", type=_boolish,
                    help="斜体")
    sp.add_argument("--underline", nargs="?", const="true", metavar="BOOL", type=_boolish,
                    help="下划线")
    sp.add_argument("--font-color", metavar="HEX", help="字体颜色,6 位 hex 如 FF0000")
    sp.add_argument("--fill", metavar="HEX|none",
                    help="单元格填充色(6 位 hex);none 清除填充")
    sp.add_argument("--align", choices=_ALIGNS, help="水平对齐")
    sp.add_argument("--valign", choices=_VALIGNS, help="垂直对齐")
    sp.add_argument("--wrap", nargs="?", const="true", metavar="BOOL", type=_boolish,
                    help="自动换行")
    sp.add_argument("--border", choices=_BORDERS, help="四边边框样式;none 清除边框")
    sp.add_argument("--border-color", metavar="HEX", help="边框颜色(与 --border 配合)")
    sp.add_argument("--num-format", metavar="FMT",
                    help="数字格式代码;general 恢复常规(引号包裹含特殊字符的格式)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)

    if xlutil.area_size(r1, c1, r2, c2) > MAX_STYLE_CELLS:
        raise CliError("too_large", f"区域 {args.range} 超过 {MAX_STYLE_CELLS} 格上限"
                                    f"(openpyxl 会为每格创建样式对象,过大会拖垮文件);"
                                    f"请缩小范围分批设置")

    if args.border_color and not args.border:
        raise CliError("bad_args", "--border-color 需要与 --border 一起使用")

    border_style = None if args.border is None else (args.border if args.border != "none" else None)
    border_color = _color(args.border_color) if (args.border_color and border_style) else None
    fill_color = _color(args.fill) if args.fill and args.fill.lower() != "none" else None
    font_color = _color(args.font_color) if args.font_color else None

    applied = [k for k, v in vars(args).items()
               if v is not None and k.startswith(("font", "bold", "italic", "underline",
                                                  "fill", "align", "valign", "wrap",
                                                  "border", "num_format"))]

    n = 0
    for row in xlutil.iter_area(ws, r1, c1, r2, c2):
        for cell in row:
            _apply_one(cell, args, border_style, border_color, fill_color, font_color)
            n += 1

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "range": xlutil.area_label(r1, c1, r2, c2),
            "cells_affected": n, "applied": sorted(set(applied))}


def _apply_one(cell, args, border_style, border_color, fill_color, font_color) -> None:
    # 字体:保留未指定项的原值,避免重置其他字体属性
    if any(v is not None for v in (args.font_name, args.font_size, args.bold,
                                   args.italic, args.underline, args.font_color)):
        old = cell.font
        cell.font = Font(
            name=args.font_name or old.name,
            size=args.font_size if args.font_size is not None else old.size,
            bold=args.bold if args.bold is not None else old.bold,
            italic=args.italic if args.italic is not None else old.italic,
            underline=("single" if args.underline else None) if args.underline is not None
                      else old.underline,
            color=Color(rgb=font_color) if font_color else old.color,
        )

    # 填充
    if args.fill is not None:
        if fill_color is None:
            cell.fill = PatternFill(fill_type=None)
        else:
            cell.fill = PatternFill(fill_type="solid", fgColor=Color(rgb=fill_color))

    # 对齐:合并已有设置
    if any(v is not None for v in (args.align, args.valign, args.wrap)):
        old = cell.alignment
        cell.alignment = Alignment(
            horizontal=args.align or old.horizontal,
            vertical=args.valign or old.vertical,
            wrap_text=args.wrap if args.wrap is not None else old.wrap_text,
        )

    # 边框:四边统一
    if args.border is not None:
        if border_style is None:
            cell.border = Border()
        else:
            side = Side(style=border_style,
                        color=Color(rgb=border_color) if border_color else None)
            cell.border = Border(left=side, right=side, top=side, bottom=side)

    # 数字格式
    if args.num_format is not None:
        if args.num_format.lower() == "general":
            cell.number_format = "General"
        else:
            cell.number_format = args.num_format


def _boolish(s: str) -> bool:
    s = s.strip().lower()
    if s in ("true", "1", "yes", "on"):
        return True
    if s in ("false", "0", "no", "off"):
        return False
    raise argparse.ArgumentTypeError(f"期望 true/false,收到 '{s}'")


def _color(hex_str: str) -> str:
    s = hex_str.strip().lstrip("#").upper()
    if len(s) == 6:
        return "FF" + s
    if len(s) == 8:
        return s
    raise CliError("bad_color", f"颜色 '{hex_str}' 不合法:应为 6 位 hex(如 FF0000)")

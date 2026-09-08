"""chart — 在表中插入基础图表(bar/col/line/pie)。"""

from __future__ import annotations

import argparse

from openpyxl.chart import BarChart, LineChart, PieChart, Reference

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "chart"
HELP = "插入图表:bar(横条)/col(柱状)/line(折线)/pie(饼图)"
DESCRIPTION = """在表中插入基础图表(openpyxl 创建,Excel/WPS 可直接查看编辑)。

用法示例:
  xcli chart add -f demo.xlsx --type col --data A1:C8 --at E2 --title "月度销量"
  xcli chart add -f demo.xlsx --type pie --data A1:B13 --title "占比"

数据区约定(--data):
- 第一列 = 类别(如月份/姓名),其余每列 = 一个数据系列
- 若数据区第一行整行都是文本,视为表头:系列名取表头文字
- pie 图只支持 2 列(类别 + 数值)

--at 指定图表左上角锚点(默认放在数据区右上方)。
输出 JSON: {"ok": true, "file": "...", "sheet": "...", "type": "col",
            "title": "...", "anchor": "E2", "series": 2, "categories": 6}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="chart_action", metavar="动作", required=True)
    s = sub.add_parser("add", help="添加图表")
    add_file_arg(s)
    add_sheet_arg(s)
    s.add_argument("--type", required=True, choices=("bar", "col", "line", "pie"),
                   help="bar=横向条形 col=竖向柱状 line=折线 pie=饼图")
    s.add_argument("--data", required=True, metavar="REF",
                   help="数据区(含类别列),如 A1:B13")
    s.add_argument("--at", metavar="REF", help="图表左上角锚点格,如 E2(默认数据区右侧)")
    s.add_argument("--title", metavar="TEXT", help="图表标题")
    s.set_defaults(chart_func="add")


def run(args: argparse.Namespace) -> dict:
    if args.chart_func != "add":  # pragma: no cover
        raise CliError("internal", f"未知动作 {args.chart_func}")
    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    r1, c1, r2, c2 = xlutil.parse_range(args.data, ws)

    if c2 < c1 + 1:
        raise CliError("bad_args", "图表数据至少需要两列:第一列是类别,其余列是数据系列")
    if args.type == "pie" and c2 > c1 + 1:
        raise CliError("bad_args", "饼图只支持 2 列数据(类别 + 数值);"
                                   "当前数据区有多列,请缩小 --data")

    first_row_is_text = _row_is_text(ws, r1, c1, c2)
    data_start_row = r1

    chart = _make_chart(args.type)
    if args.title:
        chart.title = args.title

    # 系列区:类别列之后的每一列是一个系列
    if c2 > c1:
        series = Reference(ws, min_col=c1 + 1, min_row=r1, max_col=c2, max_row=r2)
        chart.add_data(series, titles_from_data=first_row_is_text)
    # 类别:数据区第一列(有表头时从第 2 行开始)
    cat_start = r1 + 1 if first_row_is_text else r1
    cats = Reference(ws, min_col=c1, min_row=cat_start, max_col=c1, max_row=r2)
    chart.set_categories(cats)

    anchor = args.at or f"{xlutil.idx_to_col(c2 + 2)}{r1}"
    xlutil.parse_range(anchor, ws)  # 校验锚点格式
    ws.add_chart(chart, anchor)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "type": args.type, "title": args.title,
            "anchor": anchor, "series": c2 - c1, "categories": r2 - cat_start + 1}


def _row_is_text(ws, row: int, c1: int, c2: int) -> bool:
    for col in range(c1, c2 + 1):
        v = ws.cell(row=row, column=col).value
        if not (isinstance(v, str) and not xlutil.is_formula_text(v)):
            return False
    return True


def _make_chart(ctype: str):
    if ctype == "bar":
        return BarChart(type="bar")
    if ctype == "col":
        return BarChart(type="col")
    if ctype == "line":
        return LineChart()
    if ctype == "pie":
        return PieChart()
    raise CliError("bad_args", f"未知图表类型 {ctype}")  # pragma: no cover

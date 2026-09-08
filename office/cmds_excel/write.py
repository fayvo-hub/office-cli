"""write — 写入单元格/区域(支持公式),原子保存。"""

from __future__ import annotations

import argparse
import json
import os

from openpyxl import Workbook

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "write"
HELP = "向单元格/区域写入数据(支持公式),原子保存"
DESCRIPTION = """向单元格/区域写入数据(支持公式),写入成功后原子替换原文件。

用法示例:
  xcli write -f demo.xlsx --cell B2 --data 42
  xcli write -f demo.xlsx --sheet 销售 --cell A1 --data-file values.json
  xcli write -f demo.xlsx --cell A1 --data '[[\"a\",1],[\"b\",2]]'
  xcli write -f new.xlsx --cell A1 --data '[[\"x\",\"=B1*2\"]]' --create

--data 的 JSON 形状约定(三种):
  标量    42 / "hello"               -> 写单个单元格(--cell 指定)
  一维    [1, 2, 3]                  -> 自上而下写入一列(从 --cell 开始)
  二维    [[1, 2], [3, 4]]           -> 按行列写入矩形(从 --cell 开始)

值类型约定:
- null        -> 清空该格
- true/false  -> Excel 布尔
- 数字        -> 数字
- 字符串以 "=" 开头 -> 按公式写入(如 "=SUM(A1:A9)");想写 = 开头的字面量请加 --literal
- 其余字符串  -> 文本;日期请给 ISO 字符串(如 "2024-03-05T00:00:00"),
  需要日期格式时可用 style 命令设置 number_format

注意:本命令是覆盖式写入,不保留原格式以外的任何撤销能力;写入大表前建议先备份。
输出 JSON: {"ok": true, "file": "...", "sheet": "...", "range": "A1:B3",
            "values_written": 6, "sheets": [...]}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp)
    add_sheet_arg(sp)
    sp.add_argument("--cell", default="A1", metavar="REF",
                    help="写入起点单元格,如 A1 或 B2(默认 A1)")
    src = sp.add_mutually_exclusive_group(required=True)
    src.add_argument("--data", metavar="JSON",
                     help="内联 JSON 数据(见上;Shell 转义麻烦时推荐 --data-file)")
    src.add_argument("--data-file", metavar="PATH",
                     help="从 UTF-8 文件读取 JSON 数据(AI 首选,避免转义问题)")
    sp.add_argument("--create", action="store_true",
                    help="文件不存在时新建;若文件已存在则报错(防误覆盖,可加 --overwrite)")
    sp.add_argument("--overwrite", action="store_true",
                    help="配合 --create:允许覆盖已存在的文件")
    sp.add_argument("--literal", action="store_true",
                    help="把以 = 开头的字符串当作文本写入,而不是公式")


def run(args: argparse.Namespace) -> dict:
    data = _load_data(args)
    _validate_shape(data)

    exists = os.path.exists(args.file)
    if not exists:
        if not args.create:
            raise CliError("no_file", f"文件不存在: {args.file}"
                                      f"(新建文件请加 --create)")
        plan = None
        wb = Workbook()
        ws = wb.active
        ws.title = args.sheet or "Sheet1"
    else:
        if args.create and not args.overwrite:
            raise CliError("file_exists", f"文件已存在: {args.file}"
                                          f"(确认要覆盖请加 --overwrite)")
        plan = ioplan.excel_plan(args.file)
        wb = xlutil.open_workbook(plan.read_path)
        ws = xlutil.choose_sheet(wb, args.sheet)

    r0, c0, _, _ = xlutil.parse_range(args.cell, ws)
    values = _normalize(data)
    if values is None:
        if not os.path.exists(args.file):
            # --create + 空数据:仍落盘一个空表,保证文件被创建
            xlutil.save_workbook_atomic(wb, args.file)
        return {"ok": True, "file": args.file, "sheet": ws.title,
                "range": "A1", "values_written": 0, "sheets": xlutil.sheet_names_snapshot(wb)}

    rows_n, cols_n = len(values), max(len(r) for r in values)
    if r0 + rows_n - 1 > xlutil.MAX_ROWS or c0 + cols_n - 1 > xlutil.MAX_COLS:
        raise CliError("too_large",
                       f"写入范围超出 Excel 边界(行上限 {xlutil.MAX_ROWS},列上限 {xlutil.MAX_COLS});"
                       f"请缩小数据或分多次写入")

    written = 0
    for i, row_vals in enumerate(values):
        for j, v in enumerate(row_vals):
            row, col = r0 + i, c0 + j
            cell = ws.cell(row=row, column=col)
            anchor = xlutil.merged_anchor(ws, row, col)
            if anchor is not None:
                raise CliError(
                    "merged_cell",
                    f"目标格 {cell.coordinate} 位于合并区 {anchor} 内且不是左上角,无法单独写入;"
                    f"请写入左上角 {anchor},或先执行 merge --unmerge 拆分该区域")
            if v is None:
                cell.value = None
            elif isinstance(v, str) and xlutil.is_formula_text(v) and not args.literal:
                cell.value = v  # openpyxl 自动标记为公式
            else:
                cell.value = v
                if isinstance(v, str) and xlutil.is_formula_text(v):
                    # --literal:强制存为文本
                    cell.data_type = "s"
            written += 1

    out_path = plan.write_path if plan else args.file
    xlutil.save_workbook_atomic(wb, out_path)
    r2, c2 = r0 + rows_n - 1, c0 + cols_n - 1
    _up = {"upgraded_from": plan.upgraded_from} if plan and plan.upgraded_from else {}
    return {"ok": True, "file": out_path, **_up, "sheet": ws.title,
            "range": xlutil.area_label(r0, c0, r2, c2),
            "values_written": written,
            "sheets": xlutil.sheet_names_snapshot(wb)}


# ---------------------------------------------------------------------------
# 辅助
# ---------------------------------------------------------------------------

def _load_data(args) -> object:
    if args.data_file is not None:
        if not os.path.exists(args.data_file):
            raise CliError("no_file", f"--data-file 不存在: {args.data_file}")
        try:
            with open(args.data_file, encoding="utf-8") as fh:
                raw = fh.read()
        except OSError as e:
            raise CliError("read_failed", f"读取 {args.data_file} 失败: {e}") from e
    else:
        raw = args.data
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        raise CliError("bad_json", f"--data 不是合法 JSON(第 {e.lineno} 行: {e.msg});"
                                   f"推荐把数据写入文件后用 --data-file 传入") from e


def _validate_shape(data) -> None:
    if data is None or isinstance(data, (bool, int, float, str)):
        return
    if isinstance(data, list):
        if all(isinstance(x, list) for x in data):
            return
        if all(not isinstance(x, (list, dict)) for x in data):
            return
    raise CliError("bad_data",
                   "数据形状不支持:只接受标量 / 一维数组(写一列)/ 二维数组(写矩形),"
                   "元素只能是 null/布尔/数字/字符串")


def _normalize(data) -> list[list] | None:
    """-> 二维 list;空输入返回 None"""
    if data is None:
        return None
    if isinstance(data, list):
        if not data:
            return None
        if isinstance(data[0], list):
            return [list(r) for r in data]
        return [[v] for v in data]
    return [[data]]

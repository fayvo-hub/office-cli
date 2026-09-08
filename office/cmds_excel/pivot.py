"""pivot — 数据透视表:列出/修改数据源范围(openpyxl 3.1 支持读写已有透视表)。"""

from __future__ import annotations

import argparse

from .. import xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "pivot"
HELP = "数据透视表:list 列出 / set-source 修改数据源范围"
DESCRIPTION = """数据透视表操作(openpyxl ≥3.1 可读可写已有透视表,但**不能凭空创建**)。

用法示例:
  xcli pivot list -f demo.xlsx
  xcli pivot set-source -f demo.xlsx --name 数据透视表1 --ref A1:D500

子命令:
  pivot list        列出工作簿内所有透视表:名称、所在表、位置、数据源范围
  pivot set-source  修改某个透视表的数据源范围(区域需覆盖表头 + 全部数据行)

说明与限制:
- 只能操作文件中**已存在**的透视表;要新建透视表请先用 Excel 做好模板,
  再用本命令改数据源(推荐做法:模板表 + set-source 更新范围)
- set-source 默认同时设置 refreshOnLoad(Excel 打开时自动按新范围重算),
  加 --no-refresh 可关闭
- 数据源必须来自工作表区域(不支持命名区域/外部连接等源)
输出 JSON:
  list:       {"ok": true, "pivots": [{"name","sheet","location","source_range", ...}]}
  set-source: {"ok": true, "name": "...", "source_range": "$A$1:$D$500",
               "refresh_on_load": true}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="pivot_action", metavar="动作", required=True)
    s = sub.add_parser("list", help="列出所有透视表")
    add_file_arg(s)
    s.set_defaults(pivot_func="list")

    s2 = sub.add_parser("set-source", help="修改数据源范围")
    add_file_arg(s2)
    s2.add_argument("--name", required=True, metavar="NAME",
                    help="透视表名称(pivot list 可查)")
    s2.add_argument("--ref", required=True, metavar="REF",
                    help="新数据源范围,如 A1:D500(必须含表头行)")
    s2.add_argument("--no-refresh", action="store_true",
                    help="不设置 refreshOnLoad(默认设置,Excel 打开自动重算)")
    s2.set_defaults(pivot_func="set-source")


def run(args: argparse.Namespace) -> dict:
    if str(args.file).lower().endswith(".xls"):
        raise CliError(
            "unsupported_format",
            f"旧版 .xls 文件不含透视表,无法操作。"
            f"请先用 office convert -f {args.file} --out xxx.xlsx 升级后再试")
    wb = xlutil.open_workbook(args.file)

    if args.pivot_func == "list":
        pivots = _collect(wb, _sheetid_map(args.file))
        return {"ok": True, "file": args.file, "pivot_count": len(pivots),
                "pivots": pivots}

    if args.pivot_func == "set-source":
        return _set_source(wb, args)

    raise CliError("internal", f"未知动作 {args.pivot_func}")  # pragma: no cover


# ---------------------------------------------------------------------------

def _collect(wb, sheet_map: dict[int, str]):
    out = []
    for ws in wb.worksheets:
        for p in getattr(ws, "_pivots", []):
            info = {"name": p.name, "sheet": ws.title}
            loc = getattr(p, "location", None)
            if loc is not None and loc.ref:
                info["location"] = loc.ref
            cache = getattr(p, "cache", None)
            src = None
            if cache is not None:
                info["refresh_on_load"] = bool(cache.refreshOnLoad)
                info["record_count"] = cache.recordCount
                cs = getattr(cache, "cacheSource", None)
                if cs is not None:
                    src = getattr(cs, "worksheetSource", None)
            if src is not None:
                if src.ref:
                    info["source_range"] = src.ref
                if src.sheet is not None:
                    sid = int(src.sheet)
                    info["source_sheet"] = sheet_map.get(sid, f"sheetId={sid}")
                if src.name:
                    info["source_name"] = src.name
            out.append(info)
    return out


def _set_source(wb, args) -> dict:
    target = None
    for ws in wb.worksheets:
        for p in getattr(ws, "_pivots", []):
            if p.name == args.name:
                target = (ws, p)
                break
        if target:
            break
    if target is None:
        existing = [p.name for ws in wb.worksheets for p in getattr(ws, "_pivots", [])]
        raise CliError("no_pivot", f"找不到名为 '{args.name}' 的透视表;"
                                   f"现有透视表: {existing or '无(openpyxl 不能凭空创建透视表)'}")

    ws, p = target
    cache = getattr(p, "cache", None)
    cs = getattr(cache, "cacheSource", None) if cache is not None else None
    src = getattr(cs, "worksheetSource", None) if cs is not None else None
    if src is None or not src.ref:
        raise CliError("unsupported",
                       f"透视表 '{args.name}' 的数据源不是普通工作表区域"
                       f"(命名区域/外部连接等),无法修改范围")

    # 数据源所在表:通常同表;跨表时按 sheetId 定位,用于展开整列范围
    src_sheet = None
    if src.sheet is not None:
        t = _sheetid_map(args.file).get(int(src.sheet))
        if t is None or t not in wb.sheetnames:
            raise CliError("unsupported", f"透视表 '{args.name}' 引用了无法定位的源表"
                                           f"(sheetId={src.sheet}),无法修改范围")
        src_sheet = wb[t]
    else:
        src_sheet = ws  # 缺省视为本表(Excel 常用形式)

    r1, c1, r2, c2 = xlutil.parse_range(args.ref, src_sheet)
    if r1 == r2 and c1 == c2:
        raise CliError("bad_args", "数据源必须是矩形区域(表头 + 数据),不能是单格")
    new_ref = _to_abs(r1, c1, r2, c2)

    src.ref = new_ref
    if not args.no_refresh:
        cache.refreshOnLoad = True

    xlutil.save_workbook_atomic(wb, args.file)
    return {"ok": True, "file": args.file, "name": args.name,
            "sheet": ws.title, "source_range": new_ref,
            "refresh_on_load": bool(cache.refreshOnLoad),
            "tip": "已设置 refreshOnLoad:Excel 打开文件时会按新范围自动重算透视表"}


def _to_abs(r1: int, c1: int, r2: int, c2: int) -> str:
    """把区域转为 $A$1:$D$500 绝对引用形式(Excel 数据源惯例)"""
    a = f"${xlutil.idx_to_col(c1)}${r1}"
    b = f"${xlutil.idx_to_col(c2)}${r2}"
    return a if a == b else f"{a}:{b}"


def _sheetid_map(path: str) -> dict[int, str]:
    """解析 xl/workbook.xml 的 <sheet name=... sheetId=...> 映射(openpyxl 不暴露 sheetId)。"""
    import re
    import zipfile

    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/workbook.xml").decode("utf-8")
        mapping = {}
        for tag in re.findall(r"<sheet\b[^>]*>", xml):
            name_m = re.search(r'\bname="([^"]+)"', tag)
            sid_m = re.search(r'\bsheetId="(\d+)"', tag)
            if name_m and sid_m:
                mapping[int(sid_m.group(1))] = name_m.group(1)
        return mapping
    except Exception:
        return {}

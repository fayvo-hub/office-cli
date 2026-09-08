"""convert — 跨格式转换(xlsx/csv/json/xls/doc/docx/pdf/md),按扩展名自动路由。

数据类转换(xlsx<->csv/json)只搬数据、不保留样式;文档类转换按引擎能力保真。
老格式输入(.xls/.doc)自动经本机 WPS 升级,原文件不动。

用法示例:
  office convert -f a.xlsx --sheet Sheet1 --out a.csv
  office convert -f a.csv --out a.xlsx
  office convert -f 老表.xls --out 新表.xlsx      # WPS 后台升级
  office convert -f a.docx --out a.pdf            # WPS 导出(排版保真)
  office convert -f 老文档.doc --out 新文档.docx
  office convert -f a.pdf --out a.docx            # pdf2docx 引擎
  office convert -f a.md --out a.pdf              # 默认样式(精细控制用 office md to-pdf)
  office convert -f a.md --out a.docx
  office convert -f a.docx --out a.md

支持方向矩阵:
  xlsx/xlsm/xls -> csv / json / xlsx        (xls 自动升级;--sheet/--range/--cached)
  csv -> xlsx / json                        (自动识别编码 UTF-8/GBK、分隔符;可 --delimiter)
  json -> xlsx
  doc/docx -> pdf / doc/docx -> docx->md     (doc 自动升级)
  docx -> md
  pdf -> docx
  md -> pdf / docx / html

说明:
- --to 可省略,由 --out 后缀推断
- 输出文件已存在时直接覆盖;输入输出不能是同一路径
- 公式格默认输出公式文本,加 --cached 输出计算缓存值(需文件有缓存)
输出 JSON: {"ok": true, "from": "...", "to": "...", "rows": 13, "cols": 6, "warnings": []}
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import tempfile

from openpyxl import Workbook

from . import xlutil
from .cli import add_sheet_arg
from .errors import CliError

NAME = "convert"
HELP = "跨格式转换:xlsx/csv/json/xls/doc/docx/pdf/md(按扩展名自动路由)"
DESCRIPTION = __doc__

_INT_RE = re.compile(r"^[+-]?\d+$")
_FLOAT_RE = re.compile(r"^[+-]?(?:\d+\.\d*|\.\d+)(?:[eE][+-]?\d+)?$")


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PATH",
                    help="输入文件(支持 .xlsx/.xlsm/.xls/.csv/.json/.docx/.doc/.pdf/.md/.html)")
    add_sheet_arg(sp)
    sp.add_argument("--out", required=True, metavar="PATH", help="输出文件路径(后缀决定目标格式)")
    sp.add_argument("--to", metavar="FMT",
                    help="目标格式;缺省由 --out 后缀推断")
    sp.add_argument("--range", metavar="REF",
                    help="(xlsx 源)转换区域,如 A1:F13;缺省全表")
    sp.add_argument("--cached", action="store_true",
                    help="(xlsx 源)公式格输出计算缓存值而不是公式文本")
    sp.add_argument("--no-infer", action="store_true",
                    help="(csv 源)不做数字/布尔推断,全部按文本写入")
    sp.add_argument("--delimiter", metavar="CHAR", default=None,
                    help="(csv 源)强制分隔符(默认自动识别 , ; \\t |)")
    sp.add_argument("--encoding", metavar="ENC", default=None,
                    help="(csv 源)指定编码,如 gbk;默认自动尝试 UTF-8/GBK")
    sp.add_argument("--css", metavar="PATH", default=None,
                    help="(md -> pdf)附加 CSS 文件(覆盖默认样式尾部追加)")
    sp.add_argument("--engine", metavar="ENG", default=None,
                    help="(docx/doc -> pdf)转换引擎:默认 wps(本机 WPS Office)")


def run(args: argparse.Namespace) -> dict:
    src_ext = os.path.splitext(args.file)[1].lower().lstrip(".")
    dst_ext = args.to or os.path.splitext(args.out)[1].lower().lstrip(".")
    if not dst_ext:
        raise CliError("bad_args", f"无法推断输出格式,请给 --out 加后缀或加 --to")
    if src_ext in ("xlsx", "xlsm", "xls", "csv", "json"):
        if os.path.abspath(args.file) == os.path.abspath(args.out):
            raise CliError("same_file", "输入与输出是同一路径,拒绝执行")

    warnings: list[str] = []

    # ---------------- Excel/CSV/JSON 数据类 ----------------
    if src_ext in ("xlsx", "xlsm", "xls"):
        if src_ext == "xls":
            if dst_ext == "xlsx":
                from . import wps

                wps.convert(args.file, args.out)
                return {"ok": True, "from": args.file, "to": args.out,
                        "rows": None, "cols": None, "warnings":
                            [f"{args.file} 为旧版 .xls,已由 WPS 无损升级(原文件未改动)"]}
            # 其他目标:升级到临时 xlsx 再走常规流程
            fd, tmp = tempfile.mkstemp(prefix=".office-cvt-", suffix=".xlsx")
            os.close(fd)
            os.remove(tmp)
            try:
                from . import wps

                wps.convert(args.file, tmp)
                warnings.append(f"{args.file} 为旧版 .xls,已由 WPS 自动升级后转换")
                return _run_xlsx_conv(args, tmp, dst_ext)
            finally:
                try:
                    os.remove(tmp)
                except OSError:
                    pass
        return _run_xlsx_conv(args, args.file, dst_ext)

    if src_ext == "csv":
        if dst_ext not in ("xlsx", "json"):
            raise _unsupported(src_ext, dst_ext)
        rows = _read_csv(args.file, delimiter=args.delimiter, encoding=args.encoding)
        if dst_ext == "xlsx":
            _write_xlsx_rows(args.out, rows, args.sheet, infer=not args.no_infer)
        else:
            doc = {"file": args.file, "rows": rows}
            _write_text(args.out, json.dumps(doc, ensure_ascii=False, indent=2))
        return _ok(args, rows, warnings)

    if src_ext == "json":
        if dst_ext != "xlsx":
            raise _unsupported(src_ext, dst_ext)
        rows, w = _load_json_rows(args.file)
        warnings += w
        _write_xlsx_rows(args.out, rows, args.sheet, infer=False)
        return _ok(args, rows, warnings)

    # ---------------- Word 文档类 ----------------
    if src_ext in ("docx", "doc"):
        if dst_ext == "pdf":
            from . import wps

            wps.convert(args.file, args.out)
            up = [f"{args.file} 为旧版 .doc,已由 WPS 自动升级后转换"] if src_ext == "doc" else []
            return {"ok": True, "from": args.file, "to": args.out,
                    "engine": "wps", "warnings": up}
        if dst_ext == "docx" and src_ext == "doc":
            from . import wps

            wps.convert(args.file, args.out)
            return {"ok": True, "from": args.file, "to": args.out, "engine": "wps",
                    "warnings": ["旧版 .doc 已由 WPS 无损升级,原文件未改动"]}
        if dst_ext == "md":
            from .docx2md import docx_to_md

            docx_to_md(args.file, args.out)
            return {"ok": True, "from": args.file, "to": args.out, "engine": "builtin",
                    "warnings": []}
        raise _unsupported(src_ext, dst_ext)

    # ---------------- PDF ----------------
    if src_ext == "pdf":
        if dst_ext != "docx":
            raise _unsupported(src_ext, dst_ext)
        _pdf_to_docx(args.file, args.out)
        return {"ok": True, "from": args.file, "to": args.out,
                "engine": "pdf2docx",
                "warnings": ["pdf2docx 近似还原:复杂版式/特殊字体可能偏差,建议抽查"]}

    # ---------------- Markdown ----------------
    if src_ext in ("md", "markdown"):
        from .mdutil import md_to_docx, md_to_html, md_to_pdf

        if dst_ext == "pdf":
            md_to_pdf(args.file, args.out, css_path=args.css)
            return {"ok": True, "from": args.file, "to": args.out,
                    "engine": "playwright+chrome", "warnings": []}
        if dst_ext == "docx":
            md_to_docx(args.file, args.out)
            return {"ok": True, "from": args.file, "to": args.out,
                    "engine": "builtin", "warnings": []}
        if dst_ext == "html":
            md_to_html(args.file, args.out)
            return {"ok": True, "from": args.file, "to": args.out,
                    "engine": "builtin", "warnings": []}
        raise _unsupported(src_ext, dst_ext)

    raise _unsupported(src_ext, dst_ext)


# ---------------------------------------------------------------------------
# 内部:xlsx 系转换
# ---------------------------------------------------------------------------

def _run_xlsx_conv(args: argparse.Namespace, src_xlsx: str, dst_ext: str) -> dict:
    """xlsx(或已升级的临时 xlsx)-> csv/json/xlsx 原 excel convert 语义。"""
    warnings: list[str] = []
    if dst_ext not in ("csv", "json", "xlsx"):
        raise _unsupported("xlsx", dst_ext)
    if dst_ext == "xlsx":
        # xlsm -> xlsx / xlsx -> xlsx(规范化另存)
        wb = xlutil.open_workbook(src_xlsx)
        try:
            wb.save(args.out)
        except PermissionError:
            raise CliError("file_busy", f"无法写入 {args.out}: 文件可能被 Excel/WPS 打开") from None
        return {"ok": True, "from": args.file, "to": args.out,
                "rows": None, "cols": None, "warnings": warnings}

    # 构造与旧版 excel convert 一致的参数视图
    inner = argparse.Namespace(file=src_xlsx, sheet=args.sheet, out=args.out,
                               cached=args.cached, range=args.range,
                               to=dst_ext, no_infer=args.no_infer)
    # 兼容原实现:直接调用其内部逻辑
    if dst_ext == "csv":
        rows, extra, w = _read_xlsx_rows(inner)
        _write_csv(args.out, rows)
    elif dst_ext == "json":
        rows, extra, w = _read_xlsx_rows(inner)
        doc = {"file": args.file, "sheet": extra["sheet"],
               "range": extra["range"], "rows": rows,
               "merged_cells": extra["merged"]}
        _write_text(args.out, json.dumps(doc, ensure_ascii=False, indent=2))
    else:  # pragma: no cover
        raise _unsupported("xlsx", dst_ext)
    warnings += w
    return _ok(args, rows, warnings)


def _ok(args, rows: list[list], warnings: list[str]) -> dict:
    return {"ok": True, "from": args.file, "to": args.out,
            "rows": len(rows), "cols": (max((len(r) for r in rows), default=0)),
            "warnings": warnings}


def _unsupported(src_ext: str, dst_ext: str) -> CliError:
    return CliError(
        "unsupported",
        f"不支持 {src_ext} -> {dst_ext} 转换。支持方向见 office convert --help:"
        f"xlsx/csv/json 互转、xls/xlsx、doc/docx/pdf、docx/md、pdf/docx、md/pdf|docx|html")


def _read_xlsx_rows(args) -> tuple[list[list], dict, list[str]]:
    """xlsx -> 二维值(公式格输出公式文本或缓存值)"""
    wb = xlutil.open_workbook(args.file)
    ws = xlutil.choose_sheet(wb, args.sheet)
    if getattr(args, "range", None):
        r1, c1, r2, c2 = xlutil.parse_range(args.range, ws)
    else:
        dims = xlutil.data_dimensions(ws)
        if dims is None:
            return [], {"sheet": ws.title, "range": "A1", "merged": []}, ["该表无数据"]
        r1, c1, r2, c2 = dims

    wb2 = xlutil.open_workbook(args.file, data_only=True) if args.cached else None
    ws2 = wb2[ws.title] if wb2 is not None else None

    rows = []
    for row in xlutil.iter_area(ws, r1, c1, r2, c2):
        vals = []
        for cell in row:
            if cell.data_type == "f":
                if ws2 is not None:
                    vals.append(xlutil.serialize_value(
                        ws2.cell(row=cell.row, column=cell.column).value))
                else:
                    vals.append(cell.value)
            else:
                vals.append(xlutil.serialize_value(cell.value))
        rows.append(vals)
    merged = [str(r) for r in ws.merged_cells.ranges]
    return rows, {"sheet": ws.title, "range": xlutil.area_label(r1, c1, r2, c2),
                  "merged": merged}, []


def _write_csv(path: str, rows: list[list]) -> None:
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh, lineterminator="\n")
            for row in rows:
                writer.writerow(["" if v is None else _csv_scalar(v) for v in row])
    except OSError as e:
        raise CliError("write_failed", f"写入 {path} 失败: {e}") from e


def _csv_scalar(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    return str(v)


def _read_csv(path: str, delimiter: str | None = None,
              encoding: str | None = None) -> list[list]:
    """读取 CSV:自动识别编码(UTF-8/GBK)与分隔符(, ; \\t |)。"""
    if not os.path.exists(path):
        raise CliError("no_file", f"文件不存在: {path}")
    raw = None
    encodings = [encoding] if encoding else ["utf-8-sig", "utf-8", "gb18030"]
    last_err = None
    for enc in encodings:
        try:
            with open(path, encoding=enc, newline="") as fh:
                raw = fh.read()
            break
        except (UnicodeDecodeError, OSError) as e:
            last_err = e
            continue
    if raw is None:
        raise CliError("cannot_open", f"无法解码 {path}(尝试了 {','.join(encodings)}): {last_err}")
    if delimiter is None:
        delimiter = _detect_delimiter(raw)
    try:
        rows = [row for row in csv.reader(io.StringIO(raw), delimiter=delimiter)]
        # 剔除全空行(常见于文件尾换行)
        rows = [r for r in rows if any(c.strip() != "" for c in r)]
        return rows
    except Exception as e:
        raise CliError("cannot_open", f"解析 CSV {path} 失败: {e}") from e


def _detect_delimiter(raw: str) -> str:
    """统计引号外候选分隔符出现次数,取最多者(默认逗号)。"""
    candidates = [",", ";", "\t", "|"]
    counts = {c: 0 for c in candidates}
    in_quote = False
    for ch in raw[:200_000]:
        if ch == '"':
            in_quote = not in_quote
        elif ch in counts and not in_quote:
            counts[ch] += 1
    best = max(candidates, key=lambda c: counts[c])
    return best if counts[best] > 0 else ","


def _write_xlsx_rows(path: str, rows: list[list], sheet_name, *, infer: bool) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name or "Sheet1"
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j).value = _infer(v) if infer else (v if v != "" else None)
    try:
        wb.save(path)
    except PermissionError as e:
        raise CliError("file_busy", f"无法写入 {path}: 文件可能正被 Excel/WPS 打开") from e
    except OSError as e:
        raise CliError("write_failed", f"写入 {path} 失败: {e}") from e


def _infer(v) -> object:
    """csv 字符串 -> 值推断;空串 -> None"""
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return None
        if s.upper() in ("TRUE", "FALSE"):
            return s.upper() == "TRUE"
        if _INT_RE.fullmatch(s):
            try:
                return int(s)
            except ValueError:
                return v
        if _FLOAT_RE.fullmatch(s):
            try:
                return float(s)
            except ValueError:
                return v
    return v


def _load_json_rows(path: str) -> tuple[list[list], list[str]]:
    if not os.path.exists(path):
        raise CliError("no_file", f"文件不存在: {path}")
    try:
        with open(path, encoding="utf-8") as fh:
            doc = json.load(fh)
    except (OSError, json.JSONDecodeError) as e:
        raise CliError("bad_json", f"读取/解析 {path} 失败: {e}") from e
    warnings = []
    if isinstance(doc, dict) and "rows" in doc:
        if doc.get("merged_cells"):
            warnings.append("源 json 含 merged_cells,转换不重建合并区(仅搬数据)")
        rows = doc["rows"]
    elif isinstance(doc, list):
        rows = doc
    else:
        raise CliError("bad_json", f"{path} 顶层必须是二维数组或 {{\"rows\": [...]}}")
    if not all(isinstance(r, list) for r in rows):
        raise CliError("bad_json", f"{path} 的 rows 必须是二维数组([[col, col, ...], ...])")
    return rows, warnings


def _write_text(path: str, text: str) -> None:
    try:
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(text)
    except OSError as e:
        raise CliError("write_failed", f"写入 {path} 失败: {e}") from e


# ---------------------------------------------------------------------------
# 内部:PDF -> docx
# ---------------------------------------------------------------------------

def _pdf_to_docx(src: str, dst: str) -> None:
    if not os.path.exists(src):
        raise CliError("no_file", f"文件不存在: {src}")
    try:
        from pdf2docx import Converter
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf->docx 需要 pdf2docx 库,请安装: pip install pdf2docx") from None
    conv = None
    try:
        conv = Converter(src)
        conv.convert(dst, multi_paragraphs=False)
    except Exception as e:
        raise CliError("convert_failed",
                       f"pdf2docx 转换失败: {type(e).__name__}: {e}"
                       f"(扫描件 PDF(无文字层)无法转换,可先 pdf to-image 走 OCR)") from e
    finally:
        if conv is not None:
            try:
                conv.close()
            except Exception:
                pass

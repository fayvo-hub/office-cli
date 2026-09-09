# -*- coding: utf-8 -*-
"""rag prep — RAG 知识库摄取前的文档清洗与语义抽取(MVP)。

v0.1 MVP 定位(依据 rag-sample 实证缺口):
- xlsx/xlsm/xls: 多 sheet 全量导出、合并单元格展开、多行表头合成、公式缓存值回退、
  输出 LLM 友好 md + 结构化 JSON + 每表质检 warnings
- docx/doc: 走内置 docx→md 引擎(标题层级/表格保留),输出 md + 统计 JSON
- pdf: 逐页文本 + 表格抽取(复用 pdf read),页脚/页码去噪,输出 md + 结构化 JSON
- 目录模式: 批量处理并输出 qa.json 质检汇总(坏文件/公式未解析等上库前暴露)

设计原则: 纯确定性逻辑,零模型依赖;单文件无副作用只读(老格式升级产物在临时区)。
"""

from __future__ import annotations

import argparse
import json
import os
import re

from .. import ioplan, xlutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "prep"
HELP = "RAG 清洗:原生 Office/PDF → 结构完整 md + JSON(合并展开/表头合成/多sheet全出/去噪)"
DESCRIPTION = """RAG 知识库摄取前的文档清洗与语义抽取(MVP)。

用法示例:
  office rag prep -f 报表.xlsx                    # 单文件,输出到同目录 <名>.md + <名>.json
  office rag prep -f 报表.xls --out-dir clean/    # 老格式自动升级后处理
  office rag prep -f 文档目录/ --out-dir clean/    # 批量目录(扩展名 xlsx/xlsm/xls/docx/doc/pdf),
                                                   #   另写 clean/qa.json 质检汇总
  office rag prep -f 表.xlsx --header-rows 2      # 手工指定表头行数(auto=自动判定,默认)

xlsx 语义重建(相对普通转换的关键差异):
- 多 sheet 全部导出(不做静默截断);纵向合并单元格向下展开,横向合并向右展开,
  展开后每个数据行都带完整归属维度(如"华东"不丢)
- 顶部整行合并的标题识别为文档/小节标题,不混入表头
- 多行表头按列合成层级标签("2023年 / 上半年")
- 公式格:有缓存值用缓存值,无缓存保留公式原文并记入 formula_unresolved 质检项
- 纯文本 sheet(编制说明等)输出为文本块而非强行表格

输出(每个输入文件):
- <名>.md   : LLM/分块友好(Markdown 表格、层级标题)
- <名>.json : 结构化(逐表 headers/columns/rows + 行号 + 质检 warnings)
- 目录模式附加 qa.json: 各文件行数/表数/公式未解析数/告警,坏文件上库前暴露

stdout(单文件): {file, ok, md, json, format, sheets, tables, rows, warnings} —— 详细
结构见 .json 文件。目录模式 stdout: {total, succeeded, failed, warnings_total, qa}。
"""

_SUPPORTED = {".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf"}
_HEADER_AUTO_LIMIT = 4          # auto 判定时表头行数上限(超过视为无表头/全数据)
_TITLE_COVER_RATIO = 0.6        # 顶部单格合并块覆盖 ≥ 该比例列宽 → 视为标题行
_MAX_UNRESOLVED = 10            # formula_unresolved 单表最多记录条数

_PAGEFOOT_RE = re.compile(r"^\s*第\s*\d+\s*页\s*[／/]\s*共\s*\d+\s*页\s*$")
_PAGENO_RE = re.compile(r"^\s*第\s*\d+\s*页\s*$")
_LONELY_NUM_RE = re.compile(r"^\s*[-–—]?\s*\d+\s*[-–—]?\s*$")


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="输入:文件(.xlsx/.xlsm/.xls/.docx/.doc/.pdf)或目录(批量)")
    sp.add_argument("--out-dir", metavar="DIR",
                    help="输出目录;缺省=输入文件所在目录(目录批量时建议显式指定)")
    sp.add_argument("--header-rows", metavar="N|auto", default="auto",
                    help="xlsx 表头行数:auto 自动判定(默认)/ 数字固定 N / 0 无表头")
    sp.add_argument("--recursive", action="store_true",
                    help="目录模式递归子目录(默认仅当前目录)")


# ---------------------------------------------------------------------------
# 通用小工具
# ---------------------------------------------------------------------------

def _atomic_write(path: str, text: str) -> None:
    """临时文件 + os.replace 原子写(避免半截文件污染 RAG 库)。"""
    import tempfile

    directory = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp = tempfile.mkstemp(prefix=".rag-tmp-", suffix=os.path.basename(path), dir=directory)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(text)
        os.replace(tmp, path)
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise


def _cell_text(v) -> str:
    """值 → md 单元格文本(转义竖线与换行)。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "true" if v else "false"
    s = str(v)
    s = s.replace("|", "\\|").replace("\n", " ").replace("\r", " ")
    return s


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


# ---------------------------------------------------------------------------
# xlsx 语义重建
# ---------------------------------------------------------------------------

def _grid_of(ws, mr: int, mc: int) -> list[list]:
    """构建 [1..mr][1..mc] 网格(下标 0 起),并在合并区展开后填充。

    展开规则: 纵向(rows>1,cols=1)向下填;横向(rows=1,cols>1)向右填;
    二维块只保留左上角(渲染层会给出占位),返回 (grid, block2d_warn)。
    """
    grid = [[None] * mc for _ in range(mr)]
    for row in ws.iter_rows(min_row=1, min_col=1, max_row=mr, max_col=mc):
        for cell in row:
            if cell.value is not None:
                grid[cell.row - 1][cell.column - 1] = cell.value
    warn2d: list[str] = []
    for rng in ws.merged_cells.ranges:
        r1, r2, c1, c2 = rng.min_row, rng.max_row, rng.min_col, rng.max_col
        anchor = grid[r1 - 1][c1 - 1]
        if anchor is None:  # 合并块无内容(罕见),跳过
            continue
        if r1 == r2 and c1 < c2:            # 横向 → 向右展开
            for c in range(c1, c2 + 1):
                grid[r1 - 1][c - 1] = anchor
        elif c1 == c2 and r1 < r2:          # 纵向 → 向下展开
            for r in range(r1, r2 + 1):
                grid[r - 1][c1 - 1] = anchor
        else:                               # 二维块 → 仅左上
            warn2d.append(f"{ws.cell(row=r1, column=c1).coordinate}:{ws.cell(row=r2, column=c2).coordinate}")
    return grid, warn2d


def _is_title_row(grid_row: list, mc: int) -> bool:
    """顶部整行合并的标题行判定: 非空格均落于同一横向合并块且覆盖足够宽。"""
    non_empty = [(i, v) for i, v in enumerate(grid_row) if v is not None]
    if not non_empty:
        return False
    # 若所有非空值来自唯一单元格且它在一个横向合并块(网格展开后同值同宽)…
    # 展开后横向块内各格同值 → 统计非空跨度是否覆盖 mc*ratio
    vals = {v for _, v in non_empty}
    if len(non_empty) >= max(2, int(mc * _TITLE_COVER_RATIO)) and len(vals) == 1:
        return True
    return False


def _split_rows(sheet_name: str, grid: list, mc: int, header_arg: str,
                warn: list) -> tuple[list[str], list[list], list[int], list[int]]:
    """返回 (title_texts, headers[行文本数组], data[行文本数组], data_row_numbers)。

    title_texts: sheet 顶部识别出的标题行文本(可能多个,通常 1)
    headers: 表头区每行的文本序列(原始顺序,渲染时按列合成)
    data: 数据行(每行定长 mc)
    """
    mr = len(grid)
    r = 0
    titles: list[str] = []
    while r < mr and _is_title_row(grid[r], mc):
        non_empty = [v for v in grid[r] if v is not None]
        if non_empty:
            titles.append(_cell_text(non_empty[0]))
        r += 1

    def is_value_row(row: list) -> bool:
        """数据行特征: 含数值格或公式格(公式通常产生数值,如 =SUMIF)。"""
        return any(_is_number(v) or (isinstance(v, str) and xlutil.is_formula_text(v))
                   for v in row if v is not None)

    headers: list[list[str]] = []
    data_rows: list[int] = []
    if header_arg == "auto":
        # 自动判定: 表头 = 顶部连续"不含数值/公式"的行;数据起点 = 首个含数值/公式的行
        limit = _HEADER_AUTO_LIMIT
        hr = r
        gave_up = False
        while hr < mr:
            row = grid[hr]
            if not any(v is not None for v in row):      # 空行: 表头段内允许,越过
                hr += 1
                continue
            if is_value_row(row):                        # 首个数值/公式行 → 数据起点
                break
            headers.append([_cell_text(v) for v in row])
            hr += 1
            if len(headers) >= limit:                    # 连续文本行过多: 不猜了
                headers = []
                gave_up = True
                warn.append(f"前 {limit} 行以上均为文本,无法自动判定表头,整表按无表头输出"
                            f"(可用 --header-rows N 手工指定)")
                break
        while hr < mr and not any(v is not None for v in grid[hr]):
            hr += 1
        if hr < mr and is_value_row(grid[hr]):
            data_rows = list(range(hr, mr))              # 找到数据起点
        else:
            # 无数值/公式行,或已放弃表头猜测: 内容行一律保留(r 起全收)
            headers = []
            if not gave_up and r < mr:
                warn.append("该表无数值/公式行,按纯文本内容行输出")
            data_rows = [i for i in range(r, mr) if any(v is not None for v in grid[i])]
    else:
        n = int(header_arg)
        hr = r + max(0, n)
        for i in range(r, min(hr, mr)):
            headers.append([_cell_text(v) for v in grid[i]])
        data_rows = list(range(hr, mr))
    return titles, headers, data_rows, [i + 1 for i in range(len(data_rows))]


def _column_labels(headers: list[list[str]], cols: int) -> list[str]:
    """按列合成表头标签: 各行文本去连续重复后 ' / ' 连接;全空列用 列X 占位。"""
    labels: list[str] = []
    for c in range(cols):
        parts: list[str] = []
        for hrow in headers:
            t = (hrow[c] if c < len(hrow) else "").strip()
            if not t:
                continue
            if parts and parts[-1] == t:    # 纵向合并展开造成的重复
                continue
            parts.append(t)
        labels.append(" / ".join(parts) if parts else f"列{xlutil.idx_to_col(c + 1)}")
    return labels


def _prep_xlsx(src: str) -> dict:
    plan = ioplan.excel_plan(src, write=False)
    path = plan.read_path
    wb = xlutil.open_workbook(path)
    cached_wb = None  # 懒加载: 遇到公式需要缓存值时才打开

    doc_title = os.path.splitext(os.path.basename(src))[0]
    out_sheets: list[dict] = []
    file_warn: list[str] = []

    for ws in wb.worksheets:
        swarn: list[str] = []
        mr, mc = ws.max_row or 1, ws.max_column or 1
        for rng in ws.merged_cells.ranges:          # 合并区可能超出 max_row 感知
            mr = max(mr, rng.max_row)
            mc = max(mc, rng.max_col)
        if mr <= 1 and mc <= 1 and ws["A1"].value is None:
            swarn.append("空表,已跳过")
            out_sheets.append({"name": ws.title, "title": None, "columns": [],
                               "headers": [], "rows": [], "row_numbers": [],
                               "formula_unresolved": [], "warnings": swarn})
            continue

        grid, warn2d = _grid_of(ws, mr, mc)
        if warn2d:
            swarn.append("二维合并块无法展开,仅保留左上角值:" + ",".join(warn2d[:5])
                         + ("…" if len(warn2d) > 5 else ""))

        # 行宽裁剪: 取表头+数据实际非空最大列,右侧空列裁掉
        max_used = 0
        for row in grid:
            for c in range(mc - 1, -1, -1):
                if row[c] is not None:
                    max_used = max(max_used, c + 1)
                    break
        cols = max_used or 1
        grid = [row[:cols] for row in grid]

        titles, headers, data_rows, data_num = _split_rows(ws.title, grid, cols,
                                                           _header_mode, swarn)
        if not data_rows and not headers:
            swarn.append("无可抽取数据")
            out_sheets.append({"name": ws.title, "title": None, "columns": [],
                               "headers": [], "rows": [], "row_numbers": [],
                               "formula_unresolved": [], "warnings": swarn})
            continue

        # 公式缓存回退: 逐格判定公式文本
        rows_out: list[list] = []
        unresolved: list[dict] = []
        for gi in data_rows:
            row = grid[gi]
            out_row: list = []
            for c, v in enumerate(row):
                if v is None:
                    out_row.append(None)
                elif xlutil.is_formula_text(v):
                    if cached_wb is None:
                        cached_wb = xlutil.open_workbook(path, data_only=True)
                    cv = cached_wb[ws.title].cell(row=gi + 1, column=c + 1).value
                    if cv is None:
                        out_row.append(_cell_text(v))       # 公式原文(带 = 前缀)
                        if len(unresolved) < _MAX_UNRESOLVED:
                            unresolved.append({"row": gi + 1, "col": xlutil.idx_to_col(c + 1),
                                               "formula": v})
                    else:
                        out_row.append(xlutil.serialize_value(cv))
                else:
                    out_row.append(xlutil.serialize_value(v))
            rows_out.append(out_row)
        if unresolved:
            swarn.append(f"{len(unresolved)} 个公式无缓存计算值,已保留公式原文"
                         f"(用 Excel/WPS 打开另存可生成缓存)")
        if len(data_rows) != sum(1 for r in rows_out if any(v is not None for v in r)):
            dropped = len(data_rows) - sum(1 for r in rows_out if any(v is not None for v in r))
            if dropped:
                swarn.append(f"跳过 {dropped} 个全空行")
                keep = [(r_, n_) for r_, n_ in zip(rows_out, data_num) if any(v is not None for v in r_)]
                rows_out = [k[0] for k in keep]
                data_num = [k[1] for k in keep]

        labels = _column_labels(headers, cols)
        if not headers:
            swarn.append("未识别表头行,md 表头为占位列名(列A/列B…)")
        out_sheets.append({
            "name": ws.title,
            "title": titles[0] if titles else None,
            "columns": labels,
            "headers": [[_cell_text(v) for v in h] for h in headers],
            "rows": rows_out,
            "row_numbers": data_num,
            "formula_unresolved": unresolved,
            "warnings": swarn,
        })
        if titles:
            file_warn.append(f"sheet「{ws.title}」标题行已识别为标题文本,未混入表头")

    if not out_sheets:
        raise CliError("cannot_open", f"{src} 没有可读取的工作表")

    md_text = _render_xlsx_md(doc_title, out_sheets)
    return {"format": "xlsx", "doc_title": doc_title, "md": md_text,
            "sheets": out_sheets, "warnings": file_warn}


def _render_xlsx_md(doc_title: str, sheets: list[dict]) -> str:
    out: list[str] = [f"# {doc_title}"]
    for s in sheets:
        out.append("")
        out.append(f"## {s['name']}")
        if s.get("title"):
            out.append("")
            out.append(f"### {s['title']}")
        if s["rows"]:
            out.append("")
            out.append("| " + " | ".join(_cell_text(x) for x in s["columns"]) + " |")
            out.append("| " + " | ".join(["---"] * len(s["columns"])) + " |")
            for row in s["rows"]:
                padded = row + [None] * (len(s["columns"]) - len(row))
                out.append("| " + " | ".join(_cell_text(x) for x in padded) + " |")
        elif not s.get("headers") and not s.get("rows"):
            out.append("")
            out.append("*(空表)*")
    out.append("")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# docx / pdf
# ---------------------------------------------------------------------------

def _clean_pdf_text(text: str) -> str:
    """去页脚/页码噪音行(仅整行匹配;孤立页码仅当处于页首/页尾行)。"""
    # pdfminer 可能提出兼容区字符 ⻚(U+2EDA),先归一化
    text = text.replace("\u2eda", "页")
    lines = text.splitlines()
    out: list[str] = []
    for i, ln in enumerate(lines):
        if _PAGEFOOT_RE.match(ln) or _PAGENO_RE.match(ln):
            continue
        if _LONELY_NUM_RE.match(ln) and (i == 0 or i == len(lines) - 1):
            continue  # 页首/页尾孤立数字(页码)
        out.append(ln)
    return "\n".join(out)


def _prep_pdf(src: str) -> dict:
    from ..cmds_pdf import read as pdf_read  # 复用 cmds_pdf.read 的抽取管线(同进程)

    args = argparse.Namespace(file=src, pages=None, tables=True, max_chars=0)
    res = pdf_read.run(args)
    warn: list[str] = []
    cleaned_pages = []
    md_chunks: list[str] = [f"# {os.path.splitext(os.path.basename(src))[0]}"]
    for pg in res.get("pages", []):
        text = _clean_pdf_text(pg.get("text") or "")
        if text != (pg.get("text") or ""):
            warn.append(f"第 {pg['index']} 页已去除页脚/页码噪音行")
        cleaned_pages.append({"index": pg["index"], "text": text,
                              "tables": pg.get("tables", [])})
        md_chunks.append("")
        if text.strip():
            md_chunks.append(text.strip())
        for t in pg.get("tables", []):
            rows = t.get("rows_data") or []
            if not rows:
                continue
            md_chunks.append("")
            md_chunks.append("| " + " | ".join(_cell_text(x) for x in rows[0]) + " |")
            md_chunks.append("| " + " | ".join(["---"] * len(rows[0])) + " |")
            for r in rows[1:]:
                md_chunks.append("| " + " | ".join(_cell_text(x) for x in r) + " |")
        if res.get("no_text_pages"):
            warn.append("存在无文本层页面(可能是扫描件),请另行 OCR")
    return {"format": "pdf", "md": "\n".join(md_chunks) + "\n",
            "pages": cleaned_pages,
            "no_text_pages": res.get("no_text_pages", []),
            "warnings": warn + res.get("warnings", [])}


def _prep_docx(src: str, md_path: str) -> dict:
    from .. import docx2md

    plan = ioplan.word_plan(src, write=False)
    stat = docx2md.docx_to_md(plan.read_path, md_path)
    return {"format": "docx", "md_path": md_path, "stats": stat, "warnings": []}


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

_header_mode = "auto"


def _collect_inputs(src: str, recursive: bool) -> list[str]:
    if os.path.isfile(src):
        ext = os.path.splitext(src)[1].lower()
        if ext not in _SUPPORTED:
            raise CliError("unsupported_format",
                           f"rag prep 暂不支持 '{ext}',仅支持: {', '.join(sorted(_SUPPORTED))}")
        return [src]
    if not os.path.isdir(src):
        raise CliError("no_file", f"路径不存在: {src}")
    files: list[str] = []
    for name in sorted(os.listdir(src)):
        p = os.path.join(src, name)
        if os.path.isfile(p) and os.path.splitext(name)[1].lower() in _SUPPORTED:
            files.append(p)
        elif recursive and os.path.isdir(p) and not name.startswith("."):
            files.extend(_collect_inputs(p, recursive))
    if not files:
        raise CliError("no_file", f"目录 {src} 中没有支持的文档"
                                  f"({', '.join(sorted(_SUPPORTED))})")
    return files


def run(args: argparse.Namespace) -> dict:
    global _header_mode
    if args.header_rows not in ("auto",):
        try:
            n = int(args.header_rows)
            if n < 0:
                raise ValueError
            _header_mode = str(n)
        except ValueError:
            raise CliError("bad_args", "--header-rows 需为 auto 或 ≥0 的整数") from None
    else:
        _header_mode = "auto"

    files = _collect_inputs(args.file, args.recursive)
    out_dir = args.out_dir
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    results = []
    for f in files:
        ext = os.path.splitext(f)[1].lower()
        stem = os.path.splitext(os.path.basename(f))[0]
        d = out_dir or os.path.dirname(os.path.abspath(f))
        md_path = os.path.join(d, stem + ".md")
        json_path = os.path.join(d, stem + ".json")
        try:
            if ext in (".xlsx", ".xlsm", ".xls"):
                r = _prep_xlsx(f)
            elif ext == ".pdf":
                r = _prep_pdf(f)
            else:  # .docx / .doc
                r = _prep_docx(f, md_path)
            if "md" in r:  # xlsx/pdf 的 md 由本命令落地; docx 由 docx2md 落地
                _atomic_write(md_path, r.pop("md"))
            r.pop("md_path", None)
            with open(json_path, "w", encoding="utf-8") as fo:
                json.dump({"source": f, "ok": True, **r}, fo,
                          ensure_ascii=False, indent=1)
            if "sheets" in r:
                rows_total = sum(len(s.get("rows", [])) for s in r["sheets"])
                tables_total = sum(1 for s in r["sheets"] if s.get("rows"))
                stat = {"sheets": len(r["sheets"]), "tables": tables_total,
                        "rows": rows_total,
                        "formula_unresolved": sum(len(s.get("formula_unresolved", []))
                                                   for s in r["sheets"]),
                        "warnings": r["warnings"] + [w for s in r["sheets"]
                                                       for w in s.get("warnings", [])]}
            else:
                stat = {"sheets": 0, "tables": 0, "rows": 0,
                        "formula_unresolved": 0, "warnings": r.get("warnings", [])}
            results.append({"file": f, "ok": True, "md": md_path, "json": json_path,
                            "format": r["format"], **stat})
        except CliError as e:
            results.append({"file": f, "ok": False, "error": {"code": e.code,
                                                              "message": e.message}})
        except Exception as e:  # noqa: BLE001 单个文件失败不中断批量
            results.append({"file": f, "ok": False,
                            "error": {"code": "internal", "message": f"{type(e).__name__}: {e}"}})

    if len(files) == 1:
        return results[0]
    # 目录模式: 汇总 + qa.json
    qa_path = os.path.join(out_dir or os.path.dirname(os.path.abspath(args.file)),
                           "qa.json")
    summary = {"ok": True, "total": len(results),
               "succeeded": sum(1 for r in results if r["ok"]),
               "failed": sum(1 for r in results if not r["ok"]),
               "warnings_total": sum(len(r.get("warnings", [])) for r in results if r["ok"])}
    with open(qa_path, "w", encoding="utf-8") as fo:
        json.dump({"files": results, "summary": summary}, fo, ensure_ascii=False, indent=1)
    summary["qa"] = qa_path
    return summary

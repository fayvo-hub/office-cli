# -*- coding: utf-8 -*-
"""rag prep — RAG 知识库摄取前的文档清洗与语义抽取。

能力(v2):
- xlsx/xlsm/xls/csv: 多 sheet 全量导出、合并单元格展开(含二维块)、多级表头合成、
  多表分块(一个 sheet 多张表/备注行)、内置公式求值器(无缓存公式直接算,
  算不了回退缓存值,再无则保留原文记 unresolved)、百分比/日期按格式渲染、
  隐藏 sheet 剔除、纯文本 sheet 输出为 notes 块
- docx/doc: 内置 docx→md 引擎(标题层级/表格保留),输出 md + 统计 JSON
- pdf: 逐页文本 + 表格抽取(复用 pdf read),页脚/页码去噪,输出 md + 结构化 JSON
- 目录模式: 批量处理并输出 qa.json 质检汇总(坏文件/公式未解析等上库前暴露)

设计原则: 纯确定性逻辑,零模型依赖;单文件无副作用只读(老格式升级产物在临时区)。
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import time

from .. import ioplan, xlutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "prep"
HELP = "RAG 清洗:原生 Office/PDF/CSV → 结构完整 md + JSON(合并展开/公式求值/分块/去噪)"
DESCRIPTION = """RAG 知识库摄取前的文档清洗与语义抽取。

用法示例:
  office rag prep -f 报表.xlsx                    # 单文件,输出到同目录 <名>.md + <名>.json
  office rag prep -f 报表.xls --out-dir clean/    # 老格式自动升级后处理
  office rag prep -f 文档目录/ --out-dir clean/    # 批量目录(xlsx/xlsm/xls/csv/docx/doc/pdf/txt),
                                                   #   另写 clean/qa.json 质检汇总
  office rag prep -f 表.xlsx --header-rows 2      # 手工指定表头行数(auto=自动判定,默认)
  office rag prep -f 台账.csv                     # CSV 直接摄取(自动识别编码/分隔符)

xlsx 语义重建(相对普通转换的关键差异):
- 多 sheet 全部导出(隐藏 sheet 剔除);合并单元格展开: 纵向向下、横向向右、
  二维块整块铺开并在 warnings 注明坐标
- 一个 sheet 里可有多张表(空行分隔)与备注/说明段: 输出为 blocks 数组,
  每张表独立表头;纯文本段输出 notes 文本块
- 顶部整行合并的标题识别为 title,顶部短合并标题(<60% 列宽)抽为 subtitle 不混入表头
- 多级表头按列合成层级标签("2023年 / 上半年")
- 公式格: 内置求值器直接计算(跨表/条件聚合/文本/日期等);算不了回退缓存值;
  再无缓存保留公式原文并记入 formula_unresolved
- 数值格按 number_format 渲染: 百分比("24.1%")、日期(ISO)与 Excel 显示一致

输出(每个输入文件):
- <名>.md   : LLM/分块友好(Markdown 表格、层级标题)
- <名>.json : 结构化(sheets[].blocks[]: table/notes + 行号 + 质检 warnings)
- 目录模式附加 qa.json: 各文件行数/表数/公式未解析数/告警/耗时/大小

stdout(单文件): {file, ok, md, json, format, sheets, tables, rows, warnings} —— 详细
结构见 .json 文件。目录模式 stdout: {total, succeeded, failed, warnings_total, qa}。
"""

_SUPPORTED = {".xlsx", ".xlsm", ".xls", ".csv", ".docx", ".doc", ".pdf", ".txt"}
_HEADER_AUTO_LIMIT = 4          # 自动判定时表头行数上限(超过视为无表头/全数据)
_TITLE_COVER_RATIO = 0.6        # 顶部单值合并块覆盖 ≥ 该比例列宽 → 标题行
_MAX_UNRESOLVED = 10            # formula_unresolved 单块最多记录条数
_MAX_BYTES = 100 * 1024 * 1024  # 超过 100MB 拒绝处理
_PCT_RE = re.compile(r"0%|%$")          # number_format 含百分号形态(0%,0.0%,#%)
_DATE_FMT_RE = re.compile(r"[ymdhis]", re.I)  # number_format 含日期占位

_PAGEFOOT_RE = re.compile(r"^\s*第\s*\d+\s*页\s*[／/]\s*共\s*\d+\s*页\s*$")
_PAGENO_RE = re.compile(r"^\s*第\s*\d+\s*页\s*$")
_LONELY_NUM_RE = re.compile(r"^\s*[-–—]?\s*\d+\s*[-–—]?\s*$")


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="输入:文件(.xlsx/.xlsm/.xls/.csv/.docx/.doc/.pdf/.txt)或目录(批量)")
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
    return s.replace("|", "\\|").replace("\n", " ").replace("\r", " ")


def _is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _numish(v) -> bool:
    """文本行是否可视为数值行(csv/文本路径用)。"""
    if v is None:
        return False
    if _is_number(v):
        return True
    if isinstance(v, str):
        s = v.strip().rstrip("%").replace(",", "")
        try:
            float(s)
            return True
        except ValueError:
            return False
    return False


def _display(v, fmt: str | None):
    """值 → RAG 输出形态: 百分比按格式转 '24.1%';日期格式数值 → ISO;浮点修尾噪。

    文本/bool/None 原样返回。供 json rows 与 md 共用。
    """
    if v is None or isinstance(v, (bool, str)):
        return v
    if isinstance(v, (int, float)):
        f = fmt or ""
        if not isinstance(v, bool) and _PCT_RE.search(f) and "%" in f:
            return f"{v * 100:g}%"
        if isinstance(v, float) and not math.isfinite(v):
            return None  # 非法数值(如除零结果)以空呈现
        if _DATE_FMT_RE.search(f):
            # 日期格式下的裸数值 = Excel 日期序列 → ISO
            try:
                from openpyxl.utils.datetime import from_excel
                return from_excel(v).isoformat()
            except (ValueError, OverflowError):
                pass
        if isinstance(v, float):
            return round(v, 12)
        return v
    # datetime/date/time → ISO(openpyxl 已给对象);无时间成分的格式只出日期
    import datetime as _dtm
    if isinstance(v, _dtm.datetime):
        f = fmt or ""
        if not re.search(r"[hHsS]", f) and v.hour == v.minute == v.second == 0:
            return v.date().isoformat()
        return xlutil.serialize_value(v)
    return xlutil.serialize_value(v)


def _text_rows(lines: list[str]) -> str:
    """多行文本(notes)合并: 去首尾空行,内部空行保留。"""
    while lines and not lines[0].strip():
        lines.pop(0)
    while lines and not lines[-1].strip():
        lines.pop()
    return "\n".join(line.rstrip() for line in lines)


# ---------------------------------------------------------------------------
# xlsx 语义重建
# ---------------------------------------------------------------------------

def _grid_of(ws, mr: int, mc: int) -> tuple[list[list], dict, list[tuple]]:
    """构建 [1..mr][1..mc] 网格(下标 0 起)并在合并区展开。

    展开规则: 纵向(rows>1,cols=1)向下填;横向(rows=1,cols>1)向右填;
    二维块(rows>1,cols>1)整块铺满锚值(Excel 合并显示语义),坐标记入 2d_blocks。
    返回 (grid, fmt_at, 2d_blocks);fmt_at: {(r,c): number_format} 仅记录有值格。
    """
    grid = [[None] * mc for _ in range(mr)]
    fmt_at: dict = {}
    for row in ws.iter_rows(min_row=1, min_col=1, max_row=mr, max_col=mc):
        for cell in row:
            if cell.value is not None:
                grid[cell.row - 1][cell.column - 1] = cell.value
                nf = cell.number_format
                if nf and nf != "General":
                    fmt_at[(cell.row, cell.column)] = nf
    two_d: list[tuple] = []
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
        else:                               # 二维块 → 整块铺满
            two_d.append((r1, r2, c1, c2))
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    grid[r - 1][c - 1] = anchor
    return grid, fmt_at, two_d


def _span_of_row(grid_row: list, mc: int) -> int:
    """行内唯一非空值的连续跨度(展开后;0=非唯一/多段)。"""
    non_empty = [(i, v) for i, v in enumerate(grid_row) if v is not None]
    if not non_empty:
        return 0
    vals = {v for _, v in non_empty}
    if len(vals) != 1:
        return 0
    lo, hi = non_empty[0][0], non_empty[-1][0]
    # 同值必须连续(中间不能有空)
    if hi - lo + 1 != len(non_empty):
        return 0
    return hi - lo + 1


def _row_is_blank(grid_row: list) -> bool:
    return not any(v is not None for v in grid_row)


def _row_is_value(grid_row: list) -> bool:
    """数据行特征: 含数值/日期/公式的格。"""
    import datetime as _dt
    for v in grid_row:
        if v is None:
            continue
        if _is_number(v):
            return True
        if isinstance(v, (_dt.date, _dt.datetime, _dt.time)):
            return True
        if isinstance(v, str) and xlutil.is_formula_text(v):
            return True
    return False


class _SheetEnv:
    """公式求值器 FormulaEnv 适配: 直接读 openpyxl 工作簿(懒),带公式结果缓存。"""

    def __init__(self, wb, mr_of: dict, visiting: set):
        self.wb = wb
        self.mr_of = mr_of          # sheet → 逻辑 max_row(已按合并修正)
        self.cache: dict = {}       # (sheet,r,c) → 求值结果
        self.visiting = visiting

    def cell(self, sheet: str | None, row: int, col: int):
        from .formulas import FormulaError, evaluate

        title = sheet or self.wb.sheetnames[0]
        ws = self.wb[title]
        key = (title, row, col)
        if key in self.cache:
            return self.cache[key]
        v = ws.cell(row=row, column=col).value
        if isinstance(v, str) and xlutil.is_formula_text(v):
            if key in self.visiting:
                raise FormulaError("循环引用")
            self.visiting.add(key)
            try:
                r = evaluate(self, title, row, col, v)
            finally:
                self.visiting.discard(key)
            self.cache[key] = r
            return r
        return v

    def sheet_max_row(self, sheet: str | None) -> int:
        return self.mr_of.get(sheet, self.mr_of.get(next(iter(self.mr_of), ""), 1048576))

    def sheet_exists(self, sheet: str | None) -> bool:
        return (sheet or self.wb.sheetnames[0]) in self.wb.sheetnames


class _Ctx:
    """sheet 上下文: 值网格 + 格式 + 合并块(锚行横向跨度) + 行列裁剪信息。"""

    __slots__ = ("ws", "grid", "fmt_at", "two_d", "hspans", "cols", "mr")

    def __init__(self, ws, grid, fmt_at, two_d, cols):
        self.ws = ws
        self.grid = grid
        self.fmt_at = fmt_at
        self.two_d = two_d
        self.cols = cols
        self.mr = len(grid)
        # hspans: 行 → 横向合并块 [(c1,c2)];2D 块的每一行都记(铺满语义,
        # 供单值合并行剥离,避免 2D 标题底行文本污染表头)
        self.hspans = {}
        for rng in ws.merged_cells.ranges:
            r1, r2, c1, c2 = rng.min_row, rng.max_row, rng.min_col, rng.max_col
            if r1 == r2 and c2 > c1:
                self.hspans.setdefault(r1, []).append((c1, c2))
            elif r2 > r1 and c2 > c1:       # 2D 块: 顶行跨全宽(铺满语义)
                for r in range(r1, r2 + 1):
                    self.hspans.setdefault(r, []).append((c1, c2))

    def is_merged_wide(self, r: int, need_ratio: float) -> bool:
        """行 r(1 基)的首值是否来自覆盖 ≥need_ratio 列宽的横向/2D 合并块。"""
        spans = self.hspans.get(r)
        if not spans:
            return False
        row = self.grid[r - 1]
        first_c = next((i for i, v in enumerate(row) if v is not None), None)
        if first_c is None:
            return False
        for (c1, c2) in spans:
            if c1 - 1 <= first_c <= c2 - 1:
                return (c2 - c1 + 1) >= self.cols * need_ratio
        return False


def _analyze_rows(ctx: _Ctx) -> list[str]:
    """逐行分类: blank / text / value(值行可能同时含文本,按是否含数值/公式分)。"""
    cls = []
    for row in ctx.grid:
        if _row_is_blank(row):
            cls.append("blank")
        elif _row_is_value(row):
            cls.append("value")
        else:
            cls.append("text")
    return cls


def _segment_sheet(ctx: _Ctx, classes: list[str], header_arg: str,
                   warn: list) -> tuple[list[dict], list[str], list[str]]:
    """把一个 sheet 切成 blocks(不含公式求值/渲染,纯结构)。

    规则:
    - 空行是段界;段内 [text 头(≤4 行, 多则弃表头,文本行降为数据行) + value 行]
    - 纯文本段 → notes 块
    - 单值合并行剥离: 首个宽合并(≥60% 列宽)→ 块 title;其后宽合并行与顶部短合并行
      → subtitle 列表(不进表头);最多剥 3 行
    - 无表头且紧接前 table 块的纯 value 段 → 并入前块(数据续接)
    返回 (blocks, title_rows, subtitle_rows)。
    """
    mr, mc = ctx.mr, ctx.cols
    segs: list[tuple[int, int]] = []
    i = 0
    while i < mr:
        while i < mr and classes[i] == "blank":
            i += 1
        if i >= mr:
            break
        s = i
        while i < mr and classes[i] != "blank":
            i += 1
        segs.append((s, i - 1))

    blocks: list[dict] = []
    title_rows: list[str] = []
    subtitle_rows: list[str] = []

    def single_merged_row(r: int):
        """行内非空值仅来自一个合并块(连续同值 ≥2 格且在合并跨度内)时返回该值。"""
        span = _span_of_row(ctx.grid[r], mc)
        if span < 2:
            return None
        first_c = next(i for i, v in enumerate(ctx.grid[r]) if v is not None)
        for (c1, c2) in ctx.hspans.get(r + 1, ()):   # r 0 基 → 1 基行号
            if c1 - 1 <= first_c <= c2 - 1:
                return next(v for v in ctx.grid[r] if v is not None)
        return None

    for si, (s, e) in enumerate(segs):
        head_until = s
        while head_until <= e and classes[head_until] == "text":
            head_until += 1
        has_value = any(classes[r] == "value" for r in range(s, e + 1))
        seg_title = None
        seg_subs: list[str] = []
        # ---- 剥离单值合并行(title/subtitle),最多 3 行 ----
        st = s
        peeled = 0
        while st < head_until and peeled < 3:
            v = single_merged_row(st)
            if v is None:
                break
            vtxt = _cell_text(v)
            wide = ctx.is_merged_wide(st + 1, _TITLE_COVER_RATIO)
            if seg_title is None and wide:
                seg_title = vtxt
            elif seg_title is None:
                seg_title = vtxt          # 首行短合并也作块标题(G: 短合并标题)
            else:
                seg_subs.append(vtxt)     # 伴随行(单位说明等)与后续合并标题行
            st += 1
            peeled += 1
        if not has_value:
            # 全文本段: 段落行(单格 或 整行单合并块说明文)之外还有 ≥2 格的多值行 → 表格(无表头)
            def para_row(rr: int) -> bool:
                h = ctx.hspans.get(rr + 1, ())
                nz = [v for v in ctx.grid[rr] if v is not None]
                if not nz:
                    return True
                if len(h) == 1 and h[0] == (1, ctx.cols):  # 整行一个合并块=一段说明文
                    return True
                return len(nz) < 2                            # 单格(含纵向块展开)
            wide = sum(1 for rr in range(st, e + 1) if not para_row(rr))
            if (e - st + 1) >= 2 and wide * 2 >= (e - st + 1):
                if seg_title:
                    subtitle_rows.extend(seg_subs)
                    title_rows.append(seg_title)
                blocks.append({"type": "table", "title": seg_title,
                               "subtitle": seg_subs or None,
                               "headers": [],
                               "data_rows": list(range(st, e + 1)),
                               "warnings": []})
                continue
            # notes 文本块: 按行收集, 合并块展开的同值去重(含被剥离的 title/subtitle 说明)
            lines: list[str] = []
            for rr in range(st, e + 1):
                vals = [_cell_text(v) for v in ctx.grid[rr] if v is not None]
                if not vals:
                    continue
                uniq: list[str] = []
                for x in vals:
                    if not uniq or uniq[-1] != x:
                        uniq.append(x)
                lines.append(" ".join(uniq))
            joined = _text_rows(lines)
            if joined or seg_title:
                blocks.append({"type": "notes", "title": seg_title,
                               "subtitle": seg_subs or None, "text": joined,
                               "warnings": []})
            continue
        # ---- 表头候选: st 起的 text 行(上限 4,超限弃表头,文本行降为数据) ----
        text_start = st
        hr = st
        head_texts: list[str] = []
        while hr < head_until and len(head_texts) < _HEADER_AUTO_LIMIT:
            head_texts.append([_cell_text(v) for v in ctx.grid[hr]])
            hr += 1
        if hr < head_until:               # 还有未消费文本行 → 超限
            warn.append("前 %d 行以上均为文本,无法自动判定表头,该段按无表头输出"
                        "(可用 --header-rows N 手工指定)" % _HEADER_AUTO_LIMIT)
            head_texts = []
            hr = text_start
        if header_arg != "auto":
            n = int(header_arg)
            if n > 0:
                take = [r for r in range(text_start, min(text_start + n, e + 1))]
                head_texts = [[_cell_text(v) for v in ctx.grid[r]] for r in take]
                hr = max(hr, text_start + n)
            else:
                head_texts = []
        # ---- 纯 value 段紧接前 table 块 → 数据续接 ----
        if (not head_texts and not seg_title and not seg_subs
                and blocks and blocks[-1]["type"] == "table"
                and all(classes[r] == "value" for r in range(text_start, e + 1))):
            last = blocks[-1]
            last["data_rows"].extend(range(hr, e + 1))
            continue
        blocks.append({"type": "table", "title": seg_title,
                       "subtitle": seg_subs or None,
                       "headers": head_texts,
                       "data_rows": list(range(hr, e + 1)),
                       "warnings": []})
        if seg_title:
            title_rows.append(seg_title)
        subtitle_rows.extend(seg_subs)

    return blocks, title_rows, subtitle_rows


def _render_block(ctx: _Ctx, block: dict, env: _SheetEnv | None,
                  cached_provider, sheet_title: str) -> dict:
    """把结构块渲染成输出块: 公式求值、格式渲染、列标签合成、行号。

    block(data_rows/headers/title/subtitle/warnings) → out(columns/headers/rows/
    row_numbers/formula_unresolved/...) 列宽按块内实际使用裁剪。
    """
    headers = block.get("headers") or []
    data_rows = block["data_rows"]
    # 块级列裁剪: 取数据行/表头非空最大列(表头展开行含尾部空串不撑宽)
    used = 0
    for rr in data_rows:
        row = ctx.grid[rr]
        for c in range(len(row) - 1, -1, -1):
            if row[c] is not None:
                used = max(used, c + 1)
                break
    for h in headers:
        for c in range(len(h) - 1, -1, -1):
            if (h[c] or "").strip():
                used = max(used, c + 1)
                break
    cols = min(ctx.cols, used or ctx.cols)

    def fmt_at(r: int, c: int) -> str | None:
        return ctx.fmt_at.get((r, c))

    def cell_display(r: int, c: int, raw):
        """单格 → 输出值: 公式求值(失败: 缓存兜底 → 原文);统一格式渲染。"""
        nf = fmt_at(r, c)
        if not (isinstance(raw, str) and xlutil.is_formula_text(raw)):
            return _display(raw, nf)
        # ---- 公式: 求值器优先 ----
        if env is not None:
            try:
                v = env.cell(sheet_title, r, c)
                if isinstance(v, float) and not math.isfinite(v):
                    raise ValueError
                return _display(v, nf)
            except Exception:  # noqa: BLE001 FormulaError 及 env 抛出的引用错误
                pass
        # 回退: 缓存值(data_only 工作簿)
        cached_wb = cached_provider()
        if cached_wb is not None:
            cv = cached_wb[sheet_title].cell(row=r, column=c).value
            if cv is not None and not (isinstance(cv, str)
                                       and xlutil.is_formula_text(cv)):
                return _display(cv, nf)
        return None  # 调用方记 unresolved,输出原文

    out_rows: list[list] = []
    out_nums: list[int] = []
    unresolved: list[dict] = []
    for r in data_rows:          # r: 0-based grid 行索引(与 segments 一致)
        real = r + 1             # 原表 1-based 行号
        row = ctx.grid[r]
        out_row: list = []
        has_cell = False
        for c, v in enumerate(row[:cols], start=1):
            if v is None:
                out_row.append(None)
                continue
            dv = cell_display(real, c, v)
            if dv is None and isinstance(v, str) and xlutil.is_formula_text(v):
                out_row.append(_cell_text(v))  # 原文兜底
                if len(unresolved) < _MAX_UNRESOLVED:
                    unresolved.append({"row": real, "col": xlutil.idx_to_col(c),
                                       "formula": v})
            else:
                out_row.append(dv)
                has_cell = True
        if has_cell:
            out_rows.append(out_row)
            out_nums.append(real)
    if unresolved:
        block["warnings"].append(
            f"{len(unresolved)} 个公式无法求值且无缓存,已保留公式原文"
            f"(用 Excel/WPS 打开另存可生成缓存)")
    if len(data_rows) != len(out_rows):
        block["warnings"].append(f"跳过 {len(data_rows) - len(out_rows)} 个全空行")

    labels = _column_labels(headers, cols)
    if not headers and out_rows:
        block["warnings"].append("未识别表头行,列名为占位(列A/列B…)")
    return {
        "type": "table",
        "title": block.get("title"),
        "subtitle": block.get("subtitle"),
        "columns": labels,
        "headers": [[_cell_text(v) for v in h[:cols]] for h in headers],
        "rows": out_rows,
        "row_numbers": out_nums,
        "formula_unresolved": unresolved,
        "warnings": block.get("warnings") or [],
    }


def _column_labels(headers: list[list[str]], cols: int) -> list[str]:
    """按列合成表头标签: 各行文本去连续重复后 ' / ' 连接;全空列用 列X 占位。"""
    labels: list[str] = []
    for c in range(cols):
        parts: list[str] = []
        for hrow in headers:
            t = (hrow[c] if c < len(hrow) else "").strip()
            if not t:
                continue
            if parts and parts[-1] == t:    # 合并展开造成的重复
                continue
            parts.append(t)
        labels.append(" / ".join(parts) if parts else f"列{xlutil.idx_to_col(c + 1)}")
    return labels


def _notes_md_text(block: dict) -> str:
    return block.get("text") or ""


def _prep_xlsx(src: str) -> dict:
    plan = ioplan.excel_plan(src, write=False)
    path = plan.read_path
    wb = xlutil.open_workbook(path)
    cached_wb = None

    def cached_provider():
        nonlocal cached_wb
        if cached_wb is None:
            try:
                cached_wb = xlutil.open_workbook(path, data_only=True)
            except Exception:  # noqa: BLE001 缓存工作簿打不开不影响主流程
                cached_wb = False
        return cached_wb if cached_wb is not False else None

    doc_title = os.path.splitext(os.path.basename(src))[0]
    out_sheets: list[dict] = []
    file_warn: list[str] = []

    # 预计算各 sheet 逻辑行数(合并修正)供整列引用收敛
    mr_of: dict = {}
    for ws in wb.worksheets:
        mr = ws.max_row or 1
        for rng in ws.merged_cells.ranges:
            mr = max(mr, rng.max_row)
        mr_of[ws.title] = mr
    env = _SheetEnv(wb, mr_of, set())

    for ws in wb.worksheets:
        swarn: list[str] = []
        if ws.sheet_state != "visible":
            file_warn.append(f"跳过隐藏工作表「{ws.title}」")
            continue
        mr = mr_of[ws.title]
        mc = ws.max_column or 1
        for rng in ws.merged_cells.ranges:
            mc = max(mc, rng.max_col)
        if mr <= 1 and mc <= 1 and ws["A1"].value is None:
            swarn.append("空表,已跳过")
            out_sheets.append({"name": ws.title, "blocks": [],
                               "warnings": swarn})
            continue

        grid, fmt_at, two_d = _grid_of(ws, mr, mc)
        if two_d:
            joined = ",".join(
                f"{ws.cell(row=r1, column=c1).coordinate}:{ws.cell(row=r2, column=c2).coordinate}"
                for (r1, r2, c1, c2) in two_d[:5])
            swarn.append(f"{len(two_d)} 个二维合并块已整块铺开左上角值"
                         f"({joined}{'…' if len(two_d) > 5 else ''})")

        # 列裁剪: 取非空最大列(合并区可能探出)
        max_used = 0
        for row in grid:
            for c in range(mc - 1, -1, -1):
                if row[c] is not None:
                    max_used = max(max_used, c + 1)
                    break
        cols = max_used or 1
        grid = [row[:cols] for row in grid]
        ctx = _Ctx(ws, grid, fmt_at, two_d, cols)

        classes = _analyze_rows(ctx)
        blocks_raw, title_rows, sub_rows = _segment_sheet(ctx, classes,
                                                          _header_mode, swarn)
        if not blocks_raw:
            swarn.append("无可抽取数据")
            out_sheets.append({"name": ws.title, "blocks": [], "warnings": swarn})
            continue

        blocks_out = []
        for br in blocks_raw:
            if br["type"] == "notes":
                blocks_out.append(br)      # 原样(text/title/warnings)
            else:
                blocks_out.append(_render_block(ctx, br, env, cached_provider,
                                                ws.title))
        if title_rows:
            file_warn.append(f"sheet「{ws.title}」标题行已识别为块标题,未混入表头")
        if sub_rows:
            file_warn.append(f"sheet「{ws.title}」顶部短合并标题已抽为 subtitle")
        out_sheets.append({"name": ws.title, "blocks": blocks_out,
                           "warnings": swarn})

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
        blocks = s.get("blocks") or []
        if not blocks:
            out.append("")
            out.append("*(空表)*")
            continue
        for bi, b in enumerate(blocks):
            if b["type"] == "notes":
                out.append("")
                out.append(f"### 说明" if not b.get("title") else f"### {b['title']}")
                for s_ in (b.get("subtitle") or []):
                    out.append("> " + str(s_))
                for para in (b.get("text") or "").split("\n\n"):
                    out.append("")
                    out.append(para.strip())
                continue
            rows = b.get("rows") or []
            head = b.get("title")
            subs = b.get("subtitle") or []
            if head or subs:
                out.append("")
                label = head or f"表 {bi + 1}"
                out.append(f"### {label}")
                for s_ in subs:
                    out.append("")
                    out.append(f"> {s_}")     # subtitle 以引用行呈现
            if rows:
                out.append("")
                cols = b["columns"]
                out.append("| " + " | ".join(_cell_text(x) for x in cols) + " |")
                out.append("| " + " | ".join(["---"] * len(cols)) + " |")
                for row in rows:
                    padded = list(row) + [None] * (len(cols) - len(row))
                    out.append("| " + " | ".join(_cell_text(x) for x in padded) + " |")
            else:
                out.append("")
                out.append("*(空表)*")
    out.append("")
    return "\n".join(out)


# ---------------------------------------------------------------------------
# csv / txt
# ---------------------------------------------------------------------------

def _prep_csv(src: str) -> dict:
    from ..convert import _read_csv  # 复用编码探测 + 分隔符识别

    rows = _read_csv(src)
    doc_title = os.path.splitext(os.path.basename(src))[0]
    if not rows:
        raise CliError("cannot_open", f"{src} 没有可读内容")
    mc = max(len(r) for r in rows)
    # 数值化(宽松): 可转数字的文本转 float(丢失前导零与长号 — 与 Excel 打开行为一致)
    grid = []
    for r in rows:
        out_row = []
        for v in r:
            sv = (v or "").strip()
            try:
                if sv:
                    out_row.append(float(sv))
                else:
                    out_row.append(None)
            except ValueError:
                out_row.append(v if v.strip() else None)
        out_row += [None] * (mc - len(out_row))
        grid.append(out_row)
    mr = len(grid)
    fmt_at: dict = {}

    class _FakeWs:
        title = "csv"
        merged_cells = type("M", (), {"ranges": []})()

    ctx = _Ctx(_FakeWs(), grid, fmt_at, [], mc)
    classes = _analyze_rows(ctx)
    blocks_raw, _, _ = _segment_sheet(ctx, classes, _header_mode, [])
    blocks_out = []
    for br in blocks_raw:
        if br["type"] == "notes":
            blocks_out.append(br)
        else:
            blocks_out.append(_render_block(ctx, br, None, lambda: None, "csv"))
    sheet = {"name": "csv", "blocks": blocks_out,
             "warnings": [] if any(b.get("rows") for b in blocks_out) else ["无可抽取数据"]}
    md_text = _render_xlsx_md(doc_title, [sheet])
    return {"format": "csv", "doc_title": doc_title, "md": md_text,
            "sheets": [sheet], "warnings": []}


def _prep_txt(src: str) -> dict:
    encodings = ["utf-8-sig", "utf-8", "gb18030"]
    raw = None
    for enc in encodings:
        try:
            with open(src, encoding=enc) as fh:
                raw = fh.read()
            break
        except (UnicodeDecodeError, OSError):
            continue
    if raw is None:
        raise CliError("cannot_open", f"无法解码 {src}(尝试了 utf-8/gb18030)")
    doc_title = os.path.splitext(os.path.basename(src))[0]
    text = _text_rows(raw.splitlines())
    block = {"type": "notes", "title": None, "text": text, "warnings": []}
    md_text = f"# {doc_title}\n\n{text}\n"
    return {"format": "txt", "doc_title": doc_title, "md": md_text,
            "sheets": [{"name": "txt", "blocks": [block], "warnings": []}],
            "warnings": []}


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
        t0 = time.time()
        size = os.path.getsize(f)
        ext = os.path.splitext(f)[1].lower()
        stem = os.path.splitext(os.path.basename(f))[0]
        d = out_dir or os.path.dirname(os.path.abspath(f))
        md_path = os.path.join(d, stem + ".md")
        json_path = os.path.join(d, stem + ".json")
        try:
            if size > _MAX_BYTES:
                raise CliError("cannot_open", f"文件超过 {_MAX_BYTES // (1024 * 1024)}MB,"
                                              f"已拒绝处理(请拆分后重试)")
            if ext in (".xlsx", ".xlsm", ".xls"):
                r = _prep_xlsx(f)
            elif ext == ".csv":
                r = _prep_csv(f)
            elif ext == ".txt":
                r = _prep_txt(f)
            elif ext == ".pdf":
                r = _prep_pdf(f)
            else:  # .docx / .doc
                r = _prep_docx(f, md_path)
            if "md" in r:  # xlsx/csv/txt/pdf 的 md 由本命令落地; docx 由 docx2md 落地
                _atomic_write(md_path, r.pop("md"))
            r.pop("md_path", None)
            with open(json_path, "w", encoding="utf-8") as fo:
                json.dump({"source": f, "ok": True, **r}, fo,
                          ensure_ascii=False, indent=1)
            if "sheets" in r:
                rows_total = sum(len(b.get("rows", []))
                                 for s in r["sheets"] for b in s.get("blocks", []))
                tables_total = sum(1 for s in r["sheets"]
                                   for b in s.get("blocks", [])
                                   if b.get("type") == "table")
                unresolved = sum(len(b.get("formula_unresolved", []))
                                 for s in r["sheets"] for b in s.get("blocks", []))
                stat = {"sheets": len(r["sheets"]), "tables": tables_total,
                        "rows": rows_total, "formula_unresolved": unresolved,
                        "warnings": r["warnings"] + [w for s in r["sheets"]
                                                       for w in s.get("warnings", [])]}
            else:
                stat = {"sheets": 0, "tables": 0, "rows": 0,
                        "formula_unresolved": 0, "warnings": r.get("warnings", [])}
            results.append({"file": f, "ok": True, "md": md_path, "json": json_path,
                            "format": r["format"], "size": size,
                            "seconds": round(time.time() - t0, 3), **stat})
        except CliError as e:
            results.append({"file": f, "ok": False, "size": size,
                            "seconds": round(time.time() - t0, 3),
                            "error": {"code": e.code, "message": e.message}})
        except Exception as e:  # noqa: BLE001 单个文件失败不中断批量
            results.append({"file": f, "ok": False, "size": size,
                            "seconds": round(time.time() - t0, 3),
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

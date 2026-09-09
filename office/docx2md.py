"""docx2md — Word 文档(.docx)转 Markdown(近似保序转换)。

支持: 标题(Heading N)、列表(Bullet/Number)、表格、粗体/斜体/等宽、
      代码样式段落、图片占位注释;.doc 由上层 ioplan 先行升级为 .docx。
"""

from __future__ import annotations

import os
import re
import tempfile

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table, _Cell  # noqa: PLC2701 — python-docx 无公开 Cell 名
from docx.text.paragraph import Paragraph

from ._atomic import replace_with_retry
from .errors import CliError

_HEADING_RE = re.compile(r"^Heading (\d)$")
_CODE_STYLES = {"Code", "HTML Code", "Macro Text", "Source Code", "代码"}


def _iter_blocks(doc: Document):
    for child in doc.element.body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def _runs_of(p: Paragraph) -> list[tuple[str, bool, bool, bool]]:
    """遍历段内所有 w:r, 返回 [(text, bold, italic, mono)]。"""
    out: list[tuple[str, bool, bool, bool]] = []
    for r in p._p.iter(qn("w:r")):
        rPr = r.find(qn("w:rPr"))
        bold = italic = mono = False
        fonts = None
        if rPr is not None:
            bold = rPr.find(qn("w:b")) is not None
            italic = rPr.find(qn("w:i")) is not None
            rf = rPr.find(qn("w:rFonts"))
            if rf is not None:
                fonts = (rf.get(qn("w:ascii")) or "").lower()
            if rPr.find(qn("w:rStyle")) is not None:
                pass
        if fonts and any(f in fonts for f in ("consolas", "courier", "monaco")):
            mono = True
        texts = []
        for t in r.iter(qn("w:t")):
            if t.text:
                texts.append(t.text)
        for tab in r.iter(qn("w:tab")):
            texts.append("\t")
        for br in r.iter(qn("w:br")):
            texts.append("\n")
        if texts:
            out.append(("".join(texts), bold, italic, mono))
    return out


def _has_drawing(p: Paragraph) -> bool:
    return (p._p.find(qn("w:drawing")) is not None
            or p._p.find(qn("w:pict")) is not None)


def _para_text(p: Paragraph) -> str:
    """段 -> markdown 行内文本。"""
    parts: list[str] = []
    for text, bold, italic, mono in _runs_of(p):
        if not text:
            continue
        if mono:
            t = text.replace("`", "\\`")
            parts.append("`" + t + "`")
            continue
        t = text
        if bold:
            t = "**" + t.replace("*", "\\*") + "**"
        if italic:
            t = "*" + t + "*"
        parts.append(t)
    return "".join(parts)


def _cell_md(cell) -> str:
    """单元格 -> markdown 单行文本。"""
    parts = []
    for p in cell.paragraphs:
        parts.append(_para_text(p).replace("|", "\\|"))
    return "<br>".join(x for x in parts if x.strip())


def _table_md(t: Table) -> str:
    """表格 -> markdown。按 XML 单元格(tc)遍历,展开 gridSpan、跳过 vMerge
    continue 行(避免 python-docx row.cells 把合并文本重复到整列)。"""
    lines: list[str] = []
    ncols = len(t.columns)   # grid 总列数(=各 tc gridSpan 之和)
    first = True
    for row in t.rows:
        cells: list[str] = []
        for tc in row._tr.tc_lst:
            tcPr = tc.tcPr
            span, cont = 1, False
            if tcPr is not None:
                gs = tcPr.find(qn("w:gridSpan"))
                if gs is not None:
                    span = int(gs.get(qn("w:val")) or 1)
                vm = tcPr.find(qn("w:vMerge"))
                if vm is not None and (vm.get(qn("w:val")) or "restart") == "continue":
                    cont = True
            if cont:
                cells.extend([""] * span)   # 纵向合并的后续行: 空格,不重复
            else:
                cell = _Cell(tc, t)
                cells.extend([_cell_md(cell)] * span)
        if len(cells) < ncols:               # 行尾缺格补空,保证列对齐
            cells.extend([""] * (ncols - len(cells)))
        lines.append("| " + " | ".join(cells) + " |")
        if first:
            lines.append("| " + " | ".join("---" for _ in range(ncols)) + " |")
            first = False
    return "\n".join(lines)


def _list_depth(p: Paragraph) -> int:
    name = p.style.name
    m = re.search(r"(\d+)$", name)
    if m:
        return max(0, int(m.group(1)) - 1)
    return 0


def docx_to_md(docx_path: str, md_path: str) -> dict:
    """docx -> md。返回统计信息。"""
    if not os.path.exists(docx_path):
        raise CliError("no_file", f"文件不存在: {docx_path}")
    try:
        doc = Document(docx_path)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 {docx_path}: {e}") from e

    out: list[str] = []
    stat = {"paragraphs": 0, "tables": 0, "tables_skipped": 0,
            "headings": 0, "list_items": 0,
            "images_ignored": 0, "chars": 0}

    def add_blank_if_needed() -> None:
        if out and out[-1] != "":
            out.append("")

    for block in _iter_blocks(doc):
        if isinstance(block, Table):
            add_blank_if_needed()
            try:
                out.append(_table_md(block))
                stat["tables"] += 1
            except Exception as e:  # noqa: BLE001 脏文档缺 tblGrid 等: 单表降级跳过
                out.append(f"<!-- 损坏的表格已跳过: {e} -->")
                stat["tables_skipped"] += 1
            add_blank_if_needed()
            continue
        p: Paragraph = block
        style = p.style.name
        if _has_drawing(p) and not _para_text(p).strip():
            stat["images_ignored"] += 1
            continue
        text = p.text.strip()
        if not text and not _para_text(p).strip():
            continue
        m = _HEADING_RE.match(style)
        if m:
            level = int(m.group(1))
            add_blank_if_needed()
            out.append("#" * min(level, 6) + " " + _para_text(p).strip())
            add_blank_if_needed()
            stat["headings"] += 1
        elif style == "Title":
            add_blank_if_needed()
            out.append("# " + _para_text(p).strip())
            add_blank_if_needed()
            stat["headings"] += 1
        elif style.startswith("List Number"):
            depth = _list_depth(p)
            add_blank_if_needed()
            out.append("  " * depth + "1. " + _para_text(p).strip())
            stat["list_items"] += 1
        elif style.startswith("List Bullet"):
            depth = _list_depth(p)
            add_blank_if_needed()
            out.append("  " * depth + "- " + _para_text(p).strip())
            stat["list_items"] += 1
        elif style in _CODE_STYLES:
            add_blank_if_needed()
            out.append("```text")
            out.append(p.text.rstrip())
            out.append("```")
            add_blank_if_needed()
            stat["paragraphs"] += 1
        elif style == "Quote" or style.startswith("Intense Quote"):
            out.append("> " + _para_text(p).strip())
            stat["paragraphs"] += 1
        else:
            runs = _runs_of(p)
            # 空白(break/空)run 不参与判定;非空 run 须全为等宽才作代码块
            if runs and all(not t.strip() or m for t, _, _, m in runs):
                # 整段等宽(如代码块): 输出为 fenced 代码块
                add_blank_if_needed()
                out.append("```text")
                out.append("".join(t for t, *_ in runs).rstrip())
                out.append("```")
                add_blank_if_needed()
            else:
                t = _para_text(p).strip()
                out.append(t)
            stat["paragraphs"] += 1
        stat["chars"] += len(p.text)

    md_dir = os.path.dirname(os.path.abspath(md_path))
    os.makedirs(md_dir, exist_ok=True)
    body = "\n".join(out)
    if stat["images_ignored"]:
        body += (f"\n\n<!-- 注: 原文含 {stat['images_ignored']} 张图片,"
                 "docx2md 未导出图片文件 -->\n")
    _tmp = tempfile.mkstemp(suffix=".md", prefix=".office-", dir=md_dir)
    try:
        with os.fdopen(_tmp[0], "w", encoding="utf-8") as fh:
            fh.write(body)
        replace_with_retry(_tmp[1], md_path)
    except Exception:
        try:
            os.remove(_tmp[1])
        except OSError:
            pass
        raise
    stat["chars"] = len(body)
    return stat

"""pdf search — 在 PDF 文本层中搜索关键词,返回页码、次数与上下文片段。"""

from __future__ import annotations

import argparse
import re

from .. import pdfutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "search"
HELP = "全文搜索关键词(--find),返回页码/次数/上下文片段"
DESCRIPTION = """在 PDF 文本层中搜索文本(默认不区分大小写),输出每处命中位置与上下文。

用法示例:
  office pdf search -f 手册.pdf --find 保修
  office pdf search -f 合同.pdf --find '违约责任' --match-case --pages 2-8

说明:
- 返回每页命中次数与片段(命中位置前后各约 40 字符,每页最多 5 个片段)
- 纯扫描件没有文本层,不会有结果(可先 to-image + OCR)
- 加密 PDF 需先 decrypt
输出 JSON:
{"ok": true, "file": "...", "total": N,
 "matches": [{"page": 3, "count": 2, "snippets": ["..."]}], "pages_searched": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="PDF 文件路径")
    sp.add_argument("--find", required=True, metavar="TEXT", help="搜索关键词")
    sp.add_argument("--match-case", action="store_true", help="区分大小写")
    sp.add_argument("--pages", metavar="RANGE",
                    help="限定页范围,如 1-3,5(默认全部)")


def run(args: argparse.Namespace) -> dict:
    probe = pdfutil.open_reader(args.file)
    if probe.is_encrypted:
        raise CliError("encrypted",
                       f"PDF 已加密: {args.file}(请先 office pdf decrypt)")
    try:
        probe.stream.close()
    except Exception:
        pass

    try:
        import pdfplumber
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf search 需要 pdfplumber: pip install pdfplumber") from None
    import logging

    logging.getLogger("pdfminer").setLevel(logging.ERROR)

    flags = 0 if args.match_case else re.IGNORECASE
    pattern = re.compile(re.escape(args.find), flags)

    try:
        pdf = pdfplumber.open(args.file)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {args.file}: {e}") from e
    try:
        pages_sel = pdfutil.parse_pages(args.pages, len(pdf.pages))
    except CliError:
        pdf.close()
        raise

    matches: list[dict] = []
    total = 0
    CTX = 40
    try:
        for idx in pages_sel:
            text = (pdf.pages[idx].extract_text() or "").replace("\n", " ")
            hits = list(pattern.finditer(text))
            if not hits:
                continue
            total += len(hits)
            snippets = []
            for m in hits[:5]:
                s = max(0, m.start() - CTX)
                e = min(len(text), m.end() + CTX)
                snip = text[s:e]
                if s > 0:
                    snip = "…" + snip
                if e < len(text):
                    snip = snip + "…"
                snippets.append(snip)
            matches.append({"page": idx + 1, "count": len(hits),
                            "snippets": snippets})
    finally:
        pdf.close()

    return {"ok": True, "file": args.file, "total": total,
            "matches": matches, "pages_searched": len(pages_sel)}

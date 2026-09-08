"""pdf read — 提取 PDF 文本(与表格)。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "read"
HELP = "提取 PDF 文本(逐页);--tables 同时提取表格"
DESCRIPTION = """按页提取 PDF 文本层内容,输出 JSON。

输出 JSON:
{
  "ok": true, "file": "...",
  "pages": [
    {"index": 1, "chars": 320, "text": "第 1 页文本...",
     "tables": [{"rows": 3, "cols": 2, "rows_data": [[...]]}]}   // 仅 --tables
  ],
  "no_text_pages": [3, 7],          // 无文本层页(可能是扫描图)
  "warnings": []
}

说明:
- 默认提取全部页;大文件用 --pages 限定(如 --pages 1-3,5)
- 文本按阅读顺序尽量还原,复杂多栏/图文混排可能有偏差
- --tables 会逐页做版面分析,较慢;无表格的页 tables 为空
- 扫描件(纯图片)提取为空:no_text_pages 会提示,可用 pdf to-image 转图后走 OCR
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="PDF 文件路径")
    sp.add_argument("--pages", metavar="RANGE", default=None,
                    help="页范围,如 1-3,5;缺省全部")
    sp.add_argument("--tables", action="store_true",
                    help="同时提取表格(逐页版面分析,较慢)")
    sp.add_argument("--max-chars", type=int, default=8000, metavar="N",
                    help="每页文本长度上限,超出截断并警告(默认 8000;0 不限)")


def run(args: argparse.Namespace) -> dict:
    try:
        import pdfplumber
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf read 需要 pdfplumber: pip install pdfplumber") from None

    try:
        pdf = pdfplumber.open(args.file)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {args.file}: {e}") from e

    try:
        pages_sel = pdfutil.parse_pages(args.pages, len(pdf.pages))
    except CliError:
        pdf.close()
        raise

    warnings: list[str] = []
    limit = args.max_chars if args.max_chars and args.max_chars > 0 else None
    out_pages = []
    no_text = []
    try:
        for idx in pages_sel:
            page = pdf.pages[idx]
            text = page.extract_text() or ""
            n0 = len(text)
            if limit is not None and len(text) > limit:
                text = text[:limit]
                warnings.append(f"第 {idx + 1} 页文本超过 --max-chars {limit},已截断")
            if n0 == 0:
                no_text.append(idx + 1)
            item = {"index": idx + 1, "chars": n0, "text": text}
            if args.tables:
                try:
                    raw = page.extract_tables() or []
                except Exception:  # noqa: BLE001 单页版面异常不中断整体
                    raw = []
                item["tables"] = [{"rows": len(t), "cols": (len(t[0]) if t else 0),
                                   "rows_data": t} for t in raw]
            out_pages.append(item)
    finally:
        pdf.close()

    return {"ok": True, "file": args.file, "pages": out_pages,
            "no_text_pages": no_text,
            "warnings": warnings +
            (["该 PDF 无文本层(可能是扫描件),需要 OCR 可把页面转图处理"]
             if len(no_text) == len(pages_sel) and pages_sel else [])}

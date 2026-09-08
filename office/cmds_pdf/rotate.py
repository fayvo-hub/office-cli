"""pdf rotate — 旋转页面。"""

from __future__ import annotations

import argparse

from .. import pdfutil
from ..errors import CliError

NAME = "rotate"
HELP = "旋转 PDF 页面(90/180/270 度)"
DESCRIPTION = """旋转 PDF 页面。

用法:
  office pdf rotate -f a.pdf --angle 90                 # 旋转全部页(原地)
  office pdf rotate -f a.pdf --angle 90 --page 1,3      # 只旋转第 1、3 页
  office pdf rotate -f a.pdf --angle 90 --out b.pdf     # 另存为 b.pdf

说明:
- 默认原地覆盖(原子写入);--out 指定另存
- --page 格式同 --pages:1-3,5;缺省全部
输出 JSON: {"ok": true, "file": "...", "rotated_pages": 5}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--angle", type=int, choices=(90, 180, 270), required=True,
                    help="旋转角度(顺时针)")
    sp.add_argument("--page", "--pages", dest="pages", metavar="RANGE", default=None,
                    help="要旋转的页范围,如 1-3,5;缺省全部")
    sp.add_argument("--out", metavar="PATH", help="另存路径(缺省原地覆盖)")


def run(args: argparse.Namespace) -> dict:
    reader = pdfutil.open_reader(args.file)
    pdfutil.ensure_decrypted(reader)
    try:
        pages_sel = pdfutil.parse_pages(args.pages, len(reader.pages))
        try:
            from pypdf import PdfWriter
        except ImportError:  # pragma: no cover
            raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None
        writer = PdfWriter()
        for p in range(len(reader.pages)):
            page = reader.pages[p]
            if p in pages_sel:
                page.rotate(args.angle)
            writer.add_page(page)
        out = args.out or args.file
        pdfutil.save_writer(writer, out)
        return {"ok": True, "file": out, "rotated_pages": len(pages_sel)}
    finally:
        try:
            reader.stream.close()
        except Exception:
            pass

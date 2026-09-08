"""pdf to-image — PDF 页面渲染为 PNG 图片(整页转图)。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..errors import CliError

NAME = "to-image"
HELP = "PDF 页面渲染成 PNG 图片(--dpi 控制清晰度)"
DESCRIPTION = """把 PDF 页面渲染为 PNG 图片,每页一张。

用法:
  office pdf to-image -f a.pdf --out-dir pages/           # 全部页,150 dpi
  office pdf to-image -f a.pdf --pages 1-3 --dpi 300      # 只转前 3 页,高清

说明:
- dpi 越高越清晰(默认 150;打印/OCR 建议 300)
- 输出命名: <原名>_p001.png ...
- 扫描件 PDF 可用本命令转图后接 OCR;或人工查看页面内容
输出 JSON: {"ok": true, "file": "...", "out_dir": "...", "images": ["..."]}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--out-dir", required=True, metavar="DIR",
                    help="输出目录(自动创建)")
    sp.add_argument("--dpi", type=int, default=150, metavar="N", help="分辨率(默认 150)")
    sp.add_argument("--pages", metavar="RANGE", default=None,
                    help="页范围,如 1-3,5;缺省全部")


def run(args: argparse.Namespace) -> dict:
    if not 10 <= args.dpi <= 1200:
        raise CliError("bad_args", "--dpi 应在 10-1200 之间")
    try:
        import pymupdf
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf to-image 需要 PyMuPDF: pip install pymupdf") from None
    doc = None
    try:
        doc = pymupdf.open(args.file)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {args.file}: {e}") from e
    try:
        if doc.needs_pass:
            raise CliError("encrypted", "PDF 已加密,请先 office pdf decrypt 再转图")
        pages_sel = pdfutil.parse_pages(args.pages, doc.page_count)
        os.makedirs(args.out_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(args.file))[0]
        out_list = []
        for p in pages_sel:
            page = doc[p]
            pix = page.get_pixmap(dpi=args.dpi)
            path = os.path.join(args.out_dir, f"{stem}_p{p + 1:03d}.png")
            pix.save(path)
            out_list.append(path)
        return {"ok": True, "file": args.file, "out_dir": args.out_dir,
                "images": out_list, "dpi": args.dpi}
    finally:
        if doc is not None:
            doc.close()

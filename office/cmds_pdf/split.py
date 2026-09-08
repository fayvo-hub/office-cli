"""pdf split — 按页拆分/抽取 PDF。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..errors import CliError

NAME = "split"
HELP = "拆分 PDF:抽取指定页,或每页拆成独立文件"
DESCRIPTION = """拆分/抽取 PDF 页面。

两种用法(二选一):
1) 抽取页子集为新 PDF:
     office pdf split -f a.pdf --pages 1-3,5 --out part.pdf
2) 拆成每页一个文件(输出到目录):
     office pdf split -f a.pdf --pages 1-2 --out-dir pages/        # pages/a_p001.pdf ...
     office pdf split -f a.pdf --out-dir pages/                    # --pages 缺省=全部页

说明:
- --pages 格式: 1-3,5,8-   (开区间 8- 表示 8 到末页);缺省全部
- 加密文件需先 decrypt
输出 JSON: {"ok": true, "file": "...", "pages": [1,2,3], "parts": ["..."]}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--pages", metavar="RANGE", default=None,
                    help="页范围,如 1-3,5;缺省全部页")
    g = sp.add_mutually_exclusive_group(required=True)
    g.add_argument("--out", metavar="PATH", help="输出单个 PDF(抽取页合并)")
    g.add_argument("--out-dir", metavar="DIR",
                   help="输出目录:每页保存一个 PDF 文件")


def run(args: argparse.Namespace) -> dict:
    try:
        from pypdf import PdfWriter
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None

    reader = pdfutil.open_reader(args.file)
    pdfutil.ensure_decrypted(reader)
    try:
        pages_sel = pdfutil.parse_pages(args.pages, len(reader.pages))
        if args.out:
            if os.path.abspath(args.out) == os.path.abspath(args.file):
                raise CliError("same_file", "输出与输入是同一路径,请换一个 --out")
            writer = PdfWriter()
            for p in pages_sel:
                writer.add_page(reader.pages[p])
            pdfutil.save_writer(writer, args.out)
            return {"ok": True, "file": args.out, "pages": [p + 1 for p in pages_sel],
                    "parts": [args.out]}
        # 拆目录模式
        os.makedirs(args.out_dir, exist_ok=True)
        stem = os.path.splitext(os.path.basename(args.file))[0]
        parts = []
        for p in pages_sel:
            writer = PdfWriter()
            writer.add_page(reader.pages[p])
            part = os.path.join(args.out_dir, f"{stem}_p{p + 1:03d}.pdf")
            pdfutil.save_writer(writer, part)
            parts.append(part)
        return {"ok": True, "file": args.file, "pages": [p + 1 for p in pages_sel],
                "parts": parts}
    finally:
        try:
            reader.stream.close()
        except Exception:
            pass

"""pdf merge — 合并多个 PDF。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..errors import CliError

NAME = "merge"
HELP = "合并多个 PDF 为一个文件"
DESCRIPTION = """把多个 PDF 按顺序合并为一个文件。

用法: office pdf merge -f a.pdf b.pdf c.pdf --out merged.pdf

输出 JSON: {"ok": true, "file": "...", "merged_pages": 9, "sources": 3}

说明:
- 输入文件至少两个;输出不能与任一输入相同路径
- 书签/超链接不保留(仅页面内容),加密文件需先 decrypt
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", nargs="+", required=True, metavar="PDF",
                    help="输入 PDF(多个,按顺序合并)")
    sp.add_argument("-o", "--out", required=True, metavar="PATH", help="输出 PDF 路径")


def run(args: argparse.Namespace) -> dict:
    if len(args.file) < 2:
        raise CliError("bad_args", "至少需要两个输入 PDF 才能合并")
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None

    readers = []
    try:
        writer = PdfWriter()
        for f in args.file:
            if os.path.abspath(f) == os.path.abspath(args.out):
                raise CliError("same_file", f"输出路径与输入相同: {f}")
            reader = pdfutil.open_reader(f)
            readers.append(reader)
            pdfutil.ensure_decrypted(reader)
            for page in reader.pages:
                writer.add_page(page)
        total = len(writer.pages)
        pdfutil.save_writer(writer, args.out)
    finally:
        for r in readers:
            try:
                r.stream.close()
            except Exception:
                pass
    return {"ok": True, "file": args.out, "merged_pages": total,
            "sources": len(args.file)}

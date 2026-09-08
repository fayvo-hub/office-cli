"""pdf encrypt — 密码加密 PDF。"""

from __future__ import annotations

import argparse

from .. import pdfutil
from ..errors import CliError

NAME = "encrypt"
HELP = "PDF 加密(用户密码,可选所有者密码)"
DESCRIPTION = """给 PDF 加密码保护(AES-256)。

用法:
  office pdf encrypt -f a.pdf --password secret --out a_enc.pdf
  office pdf encrypt -f a.pdf --password secret         # 原地加密(覆盖原文件)

说明:
- --owner 可单独设置所有者密码(控制打印/复制等权限);缺省同用户密码
- 已加密文件会先报错,请先 decrypt
输出 JSON: {"ok": true, "file": "...", "algorithm": "AES-256"}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--password", required=True, metavar="PWD", help="用户打开密码")
    sp.add_argument("--owner", metavar="PWD", help="所有者密码(缺省同用户密码)")
    sp.add_argument("--out", metavar="PATH", help="另存路径(缺省原地覆盖)")


def run(args: argparse.Namespace) -> dict:
    reader = pdfutil.open_reader(args.file)
    if reader.is_encrypted:
        raise CliError("encrypted", "文件已加密,请先 decrypt 后再加密")
    try:
        from pypdf import PdfWriter
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(user_password=args.password,
                   owner_password=args.owner or args.password,
                   algorithm="AES-256")
    out = args.out or args.file
    pdfutil.save_writer(writer, out)
    return {"ok": True, "file": out, "algorithm": "AES-256"}

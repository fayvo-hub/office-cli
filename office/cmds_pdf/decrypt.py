"""pdf decrypt — 移除 PDF 密码。"""

from __future__ import annotations

import argparse

from .. import pdfutil
from ..errors import CliError

NAME = "decrypt"
HELP = "解密 PDF(用密码去除保护)"
DESCRIPTION = """用密码解除 PDF 加密,输出无密码副本。

用法:
  office pdf decrypt -f a.pdf --password secret --out a_open.pdf
  office pdf decrypt -f a.pdf --password secret          # 原地覆盖

说明:
- 密码错误会给出提示;文件未加密直接报错(无需本命令)
输出 JSON: {"ok": true, "file": "..."}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--password", required=True, metavar="PWD", help="文件密码")
    sp.add_argument("--out", metavar="PATH", help="另存路径(缺省原地覆盖)")


def run(args: argparse.Namespace) -> dict:
    reader = pdfutil.open_reader(args.file)
    if not reader.is_encrypted:
        try:
            reader.stream.close()
        except Exception:
            pass
        raise CliError("not_encrypted", "文件未加密,无需解密")
    try:
        from pypdf import PdfWriter
        from pypdf.errors import FileNotDecryptedError
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None
    try:
        reader.decrypt(args.password)
    except FileNotDecryptedError:
        raise CliError("bad_password", "密码错误,无法解密(请确认密码后重试)") from None
    except Exception as e:
        raise CliError("bad_password", f"解密失败: {e}") from e
    try:
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
        out = args.out or args.file
        pdfutil.save_writer(writer, out)
        return {"ok": True, "file": out}
    finally:
        try:
            reader.stream.close()
        except Exception:
            pass

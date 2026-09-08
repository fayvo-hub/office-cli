"""pdf info — PDF 元信息:页数、尺寸、加密状态、文档属性。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "info"
HELP = "PDF 基本信息:页数/页面尺寸/加密状态/元数据"
DESCRIPTION = """输出 PDF 基本信息(快速,不解析全部内容)。

输出 JSON:
{
  "ok": true, "file": "...",
  "pages": 12,
  "encrypted": false,
  "size_bytes": 204800,
  "metadata": {"title": null, "author": null, "subject": null,
               "keywords": null, "creator": null, "producer": null},
  "first_page": {"width_pt": 595.28, "height_pt": 841.89,
                 "width_mm": 210.0, "height_mm": 297.0}
}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="PDF 文件路径")


def run(args: argparse.Namespace) -> dict:
    reader = pdfutil.open_reader(args.file)
    meta = {}
    try:
        if reader.metadata is not None:
            for k in ("title", "author", "subject", "keywords", "creator", "producer"):
                v = getattr(reader.metadata, k, None)
                meta[k] = str(v) if v else None
    except Exception:
        pass  # 加密未解锁时元数据不可读
    pages = first = None
    try:
        pages = len(reader.pages)
        first = pdfutil.page_size_mm(reader.pages[0]) if pages else None
    except Exception:
        pass  # 加密未解锁时页数不可得, 仅报 encrypted
    try:
        reader.stream.close()
    except Exception:
        pass
    return {"ok": True, "file": args.file, "pages": pages,
            "encrypted": bool(reader.is_encrypted),
            "size_bytes": os.path.getsize(args.file),
            "metadata": meta, "first_page": first}

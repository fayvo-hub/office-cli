"""pdf from-images — 多张图片合成一个 PDF(每张一页,等比适配 A4)。"""

from __future__ import annotations

import argparse
import os
import re

from .. import pdfutil
from ..errors import CliError

NAME = "from-images"
HELP = "多张图片合成 PDF(--out out.pdf 图片...;每张一页适配 A4)"
DESCRIPTION = """把若干图片(png/jpg/webp/bmp/tiff…)按顺序合成一个 PDF。

用法示例:
  office pdf from-images -f 不存在 --out 相册.pdf a.jpg b.png
  office pdf from-images --out 合同扫描.pdf 扫描1.jpg 扫描2.jpg 扫描3.jpg

说明:
- 文件名按自然顺序排序(1.jpg, 2.jpg, …, 10.jpg),可自由乱序传参
- 每张图片占一页:图片等比缩放并居中放入 A4 纵向页(白边留白)
- 需要 PyMuPDF 库
输出 JSON: {"ok": true, "file": "...", "pages": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("images", nargs="+", metavar="IMG", help="图片文件(png/jpg/…)")
    sp.add_argument("--out", required=True, metavar="PDF", help="输出 PDF 路径")
    sp.add_argument("-f", "--file", help=argparse.SUPPRESS)  # 组内统一参数,忽略


def _natural_key(text: str):
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r"(\d+)", text)]


def run(args: argparse.Namespace) -> dict:
    images = sorted(args.images, key=_natural_key)
    for p in images:
        if not os.path.exists(p):
            raise CliError("no_file", f"图片不存在: {p}")
    try:
        fitz = pdfutil.import_fitz()
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 PyMuPDF 库: pip install PyMuPDF") from None

    A4W, A4H = 595.27, 841.89  # pt
    try:
        doc = fitz.open()
        for p in images:
            pm = fitz.Pixmap(p)
            w_pt, h_pt = pm.width * 72 / 96, pm.height * 72 / 96
            scale = min(A4W / w_pt, A4H / h_pt, 1.0)
            w, h = w_pt * scale, h_pt * scale
            page = doc.new_page(width=A4W, height=A4H)
            x0, y0 = (A4W - w) / 2, (A4H - h) / 2
            page.insert_image(fitz.Rect(x0, y0, x0 + w, y0 + h), filename=p)
    except CliError:
        raise
    except Exception as e:
        raise CliError("image_error", f"合成图片失败: {e}") from e

    out = args.out
    try:
        pdfutil.save_doc_atomic(out, doc, garbage=4, deflate=True)
    except CliError:
        raise
    except Exception as e:
        raise CliError("write_failed", f"写入 {out} 失败: {e}") from e
    finally:
        try:
            if not doc.is_closed:
                doc.close()
        except Exception:
            pass
    return {"ok": True, "file": out, "pages": len(images)}

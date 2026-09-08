"""pdf images — 导出 PDF 内嵌图片。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..errors import CliError

NAME = "images"
HELP = "导出 PDF 内嵌的全部图片到目录"
DESCRIPTION = """把 PDF 里内嵌的图片(照片/插图/扫描图块)导出为独立文件。

用法: office pdf images -f a.pdf --out-dir imgs/

输出 JSON:
{"ok": true, "file": "...", "out_dir": "...", "images": [
  {"path": "imgs/p1-0.png", "page": 1, "index": 0, "width": 800, "height": 600}],
 "count": 3}

说明:
- 同页多图按出现顺序编号;跨页重复引用的图片按首次出现页导出一次
- 注意:页面背景/矢量图形不是内嵌图片,不会出现在这里;
  想整页转图请用 pdf to-image
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--out-dir", required=True, metavar="DIR",
                    help="图片输出目录(自动创建)")


def run(args: argparse.Namespace) -> dict:
    try:
        import pymupdf
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf images 需要 PyMuPDF: pip install pymupdf") from None
    doc = None
    try:
        doc = pymupdf.open(args.file)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {args.file}: {e}") from e
    try:
        if doc.needs_pass:
            raise CliError("encrypted", "PDF 已加密,请先 office pdf decrypt 再提取图片")
        os.makedirs(args.out_dir, exist_ok=True)
        seen: set[int] = set()
        out_list = []
        stem = os.path.splitext(os.path.basename(args.file))[0]
        for pno, page in enumerate(doc, start=1):
            infos = page.get_images(full=True)
            for k, img in enumerate(infos):
                xref = img[0]
                if xref in seen:
                    continue
                seen.add(xref)
                try:
                    data = doc.extract_image(xref)
                except Exception:  # noqa: BLE001
                    continue
                ext = data.get("ext", "png")
                path = os.path.join(args.out_dir, f"{stem}_p{pno}-{k}.{ext}")
                with open(path, "wb") as fh:
                    fh.write(data["image"])
                out_list.append({"path": path, "page": pno, "index": k,
                                 "width": data.get("width"),
                                 "height": data.get("height")})
        if not out_list:
            raise CliError("no_images",
                           "PDF 内没有可导出的内嵌图片。整页转图请用 pdf to-image")
        return {"ok": True, "file": args.file, "out_dir": args.out_dir,
                "images": out_list, "count": len(out_list)}
    finally:
        if doc is not None:
            doc.close()

"""pdf watermark — 给 PDF 每页加半透明文字水印。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..errors import CliError

NAME = "watermark"
HELP = "给 PDF 添加半透明文字水印(如: 机密 / 内部资料 / 草稿)"
DESCRIPTION = """给 PDF 每页添加居中半透明文字水印(支持中文)。

用法:
  office pdf watermark -f a.pdf --text 内部资料
  office pdf watermark -f a.pdf --text 机密 --size 48 --out w.pdf
  office pdf watermark -f a.pdf --text 内部资料 --no-rotate

说明:
- 默认文字旋转 45° 铺在页面中央;--no-rotate 水平摆放
- 中文字体自动用系统字体(微软雅黑/黑体);--size 字号(默认按页高 12%)
- 默认原地覆盖;--out 另存
输出 JSON: {"ok": true, "file": "...", "watermarked_pages": 12, "font": "msyh"}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF", help="输入 PDF")
    sp.add_argument("--text", required=True, metavar="TEXT", help="水印文字(中文可用)")
    sp.add_argument("--size", type=float, default=0, metavar="PT",
                    help="字号(pt);缺省按页面高度 12% 自动")
    sp.add_argument("--opacity", type=float, default=0.18, metavar="0-1",
                    help="不透明度(默认 0.18)")
    sp.add_argument("--no-rotate", action="store_true", help="水印不旋转(水平摆放)")
    sp.add_argument("--out", metavar="PATH", help="另存路径(缺省原地覆盖)")


def _find_cjk_font() -> str | None:
    for p in (r"C:\Windows\Fonts\msyh.ttc", r"C:\Windows\Fonts\msyh.ttf",
              r"C:\Windows\Fonts\simhei.ttf", r"C:\Windows\Fonts\simsun.ttc"):
        if os.path.exists(p):
            return p
    return None


def run(args: argparse.Namespace) -> dict:
    if not 0 < args.opacity <= 1:
        raise CliError("bad_args", "--opacity 必须在 (0, 1] 之间")
    try:
        import pymupdf
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "pdf watermark 需要 PyMuPDF: pip install pymupdf") from None

    font_path = _find_cjk_font()
    doc = None
    try:
        doc = pymupdf.open(args.file)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {args.file}: {e}") from e
    try:
        if doc.needs_pass:
            raise CliError("encrypted", "PDF 已加密,请先 office pdf decrypt 再加水印")
        font = None
        if font_path:
            font = pymupdf.Font(fontfile=font_path)
        n = 0
        for page in doc:
            r = page.rect
            size = args.size or max(20.0, r.height * 0.12)
            tw = pymupdf.TextWriter(r)
            try:
                text_w = pymupdf.get_text_length(args.text, font=font, fontsize=size)
            except Exception:
                text_w = size * len(args.text)
            # 居中:文字近似占位
            pos = pymupdf.Point((r.width - text_w) / 2, (r.height + size * 0.35) / 2)
            tw.append(pos, args.text, font=font, fontsize=size)
            center = pymupdf.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
            if args.no_rotate:
                tw.write_text(page, color=(0.45, 0.45, 0.45), opacity=args.opacity)
            else:
                tw.write_text(page, color=(0.45, 0.45, 0.45), opacity=args.opacity,
                              morph=(center, pymupdf.Matrix(45, 45)))
            n += 1
        out = args.out or args.file
        if os.path.abspath(out) == os.path.abspath(args.file):
            pdfutil.write_atomic(out, lambda fh, d=doc: d.save(fh))  # 原子覆盖
        else:
            doc.save(out)
        return {"ok": True, "file": out, "watermarked_pages": n,
                "font": os.path.basename(font_path) if font_path else "base14",
                "text": args.text}
    finally:
        if doc is not None:
            doc.close()

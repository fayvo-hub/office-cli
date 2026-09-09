"""pdf footer — 给 PDF 每页底部添加页脚文本(页码等)。"""

from __future__ import annotations

import argparse

from .. import pdfutil
from ..errors import CliError

NAME = "footer"
HELP = "添加页脚文本(支持 {page}/{pages} 占位符)"
DESCRIPTION = """给 PDF 的每页(或指定页)底部居中添加一行页脚文字。

用法示例:
  office pdf footer -f 报告.pdf --text '第 {page} 页 / 共 {pages} 页'
  office pdf footer -f 报告.pdf --text '内部资料 · {page}' --pages 2-5
  office pdf footer -f 报告.pdf --text '机密' --size 10 --color 990000

说明:
- {page} = 当前页码,{pages} = 总页数;纯文本(无占位符)每页相同
- 加密 PDF 需先 decrypt;会覆盖原文件,建议先备份
- 需要 PyMuPDF 库
输出 JSON: {"ok": true, "file": "...", "pages": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PDF",
                    help="PDF 文件路径")
    sp.add_argument("--text", default="第 {page} 页 / 共 {pages} 页",
                    metavar="TEXT", help="页脚文本(默认页码)")
    sp.add_argument("--pages", metavar="RANGE",
                    help="只加在指定页,如 1 / 2-5 / 1,3,5-7(默认全部)")
    sp.add_argument("--size", type=float, default=9, metavar="PT",
                    help="字号(默认 9)")
    sp.add_argument("--margin", type=float, default=24, metavar="PT",
                    help="距页面底部距离(默认 24)")
    sp.add_argument("--color", default="666666", metavar="HEX",
                    help="颜色,如 666666 或 FF0000(默认 666666)")


def run(args: argparse.Namespace) -> dict:
    reader = pdfutil.open_reader(args.file)
    if reader.is_encrypted:
        raise CliError("encrypted",
                       f"PDF 已加密: {args.file}(请先 office pdf decrypt)")
    try:
        reader.stream.close()
    except Exception:
        pass
    try:
        fitz = pdfutil.import_fitz()
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 PyMuPDF 库: pip install PyMuPDF") from None

    try:
        import os

        CJK_FONTS = (r"C:/Windows/Fonts/msyh.ttc",
                     r"C:/Windows/Fonts/simhei.ttf",
                     r"C:/Windows/Fonts/msyh.ttf")
        doc = fitz.open(args.file)
        if doc.needs_pass:
            raise CliError("encrypted",
                           "PDF 已加密,需要密码。请先执行 office pdf decrypt")
        total = doc.page_count
        pages = pdfutil.parse_pages(args.pages, total) if args.pages else list(range(total))
        color = _parse_color(args.color)
        for i in pages:
            if not 0 <= i < total:
                raise CliError("bad_args",
                               f"页码越界: {i + 1}(共 {total} 页)")
            page = doc[i]
            text = args.text.replace("{page}", str(i + 1)) \
                            .replace("{pages}", str(total))
            if not text:
                continue
            if any(ord(ch) > 127 for ch in text):
                # 含非 ASCII 时用系统中文字体(helv 无 CJK 字形)。
                # 注意:insert_text+fontfile 对同一字体的多次插入会共享子集导致
                # 字符乱序,因此用每页独立的 fitz.Font + TextWriter
                ff = next((c for c in CJK_FONTS if os.path.exists(c)), None)
                if ff is None:
                    raise CliError("font_missing",
                                   "未找到中文字体(msyh.ttc/simhei.ttf),无法写入中文页脚")
                fnt = fitz.Font(fontname=f"cjf{os.getpid()}{i}", fontfile=ff)
            else:
                fnt = fitz.Font("helv")
            width = fnt.text_length(text, fontsize=args.size)
            x = (page.rect.width - width) / 2
            y = page.rect.height - args.margin
            tw = fitz.TextWriter(page.rect)
            tw.append(fitz.Point(x, y), text, font=fnt, fontsize=args.size)
            tw.write_text(page, color=color)
        pdfutil.save_doc_atomic(args.file, doc, garbage=4, deflate=True)
    except CliError:
        raise
    except Exception as e:
        raise CliError("pdf_error", f"添加页脚失败: {e}") from e
    finally:
        try:
            if not doc.is_closed:
                doc.close()
        except Exception:
            pass
    return {"ok": True, "file": args.file, "pages": len(pages)}


def _parse_color(hex_text: str) -> tuple[float, float, float]:
    h = hex_text.strip().lstrip("#")
    if len(h) != 6:
        raise CliError("bad_args", f"--color 应为 6 位十六进制,如 666666,收到 {hex_text!r}")
    try:
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        raise CliError("bad_args", f"--color 不是合法十六进制: {hex_text!r}") from None
    return (r, g, b)

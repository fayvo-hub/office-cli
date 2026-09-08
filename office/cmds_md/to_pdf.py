"""md to-pdf — Markdown 渲染为带样式 PDF。"""

from __future__ import annotations

import argparse
import os

from .. import mdutil
from ..errors import CliError
from ..pdfutil import open_reader

NAME = "to-pdf"
HELP = "Markdown 转 PDF(mermaid 图/代码高亮/表格,保留样式)"
DESCRIPTION = """把 Markdown 渲染成排版好的 A4 PDF。

用法:
  office md to-pdf -f README.md --out README.pdf
  office md to-pdf -f 报告.md                  # 输出同目录同名 .pdf
  office md to-pdf -f a.md --css extra.css     # 附加自定义样式

说明:
- 支持: 标题层级、表格、代码高亮(pygments)、```mermaid 流程图/时序图、
  引用、图片(本地相对路径)
- 每页底部自动带页码;用系统 Chrome/Edge 渲染(找不到时报错,可设 OFFICE_CHROME)
输出 JSON: {"ok": true, "file": "...", "pages": 3, "mermaid": true}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="MD", help="输入 Markdown")
    sp.add_argument("--out", metavar="PDF", help="输出 PDF(缺省同名 .pdf)")
    sp.add_argument("--css", metavar="CSS", help="附加的自定义样式表")


def run(args: argparse.Namespace) -> dict:
    out = args.out or os.path.splitext(args.file)[0] + ".pdf"
    if out.lower() == args.file.lower():
        raise CliError("bad_args", "--out 不能与输入相同")
    info = mdutil.md_to_pdf(args.file, out, css_path=args.css)
    try:
        reader = open_reader(out)
        pages = len(reader.pages)
        try:
            reader.stream.close()
        except Exception:
            pass
    except CliError:
        pages = 0
    return {"ok": True, "file": out, "pages": pages,
            "mermaid": info.get("mermaid_ok", True)}

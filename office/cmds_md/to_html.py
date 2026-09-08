"""md to-html — Markdown 渲染为单文件 HTML(可打印/预览)。"""

from __future__ import annotations

import argparse
import os

from .. import mdutil

NAME = "to-html"
HELP = "Markdown 转单文件 HTML(mermaid/代码高亮内联)"
DESCRIPTION = """把 Markdown 渲染成一个自包含 HTML 文件(样式/脚本全内联)。

用法:
  office md to-html -f README.md --out README.html
  office md to-html -f 报告.md                  # 输出同目录同名 .html

说明:
- 浏览器打开即可预览;mermaid 图直接可用;也可用本工具转 PDF
输出 JSON: {"ok": true, "file": "..."}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="MD", help="输入 Markdown")
    sp.add_argument("--out", metavar="HTML", help="输出 .html(缺省同名 .html)")


def run(args: argparse.Namespace) -> dict:
    out = args.out or os.path.splitext(args.file)[0] + ".html"
    if out.lower() == args.file.lower():
        from ..errors import CliError
        raise CliError("bad_args", "--out 不能与输入相同")
    mdutil.md_to_html(args.file, out)
    return {"ok": True, "file": out}

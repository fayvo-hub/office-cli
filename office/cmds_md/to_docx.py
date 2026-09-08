"""md to-docx — Markdown 转 Word 文档。"""

from __future__ import annotations

import argparse
import os

from .. import mdutil

NAME = "to-docx"
HELP = "Markdown 转 Word(.docx,标题/列表/表格/代码/图片)"
DESCRIPTION = """把 Markdown 转成 Word 文档(近似转换)。

用法:
  office md to-docx -f README.md --out README.docx
  office md to-docx -f 报告.md                  # 输出同目录同名 .docx

说明:
- 标题→Word 标题样式;列表/表格(Table Grid)/引用/代码(灰底等宽)均保留;
  本地图片按原尺寸插入(宽 >15cm 自动缩放);```mermaid 图转为代码文本
- 反向 word read / convert docx→md 可还原大部分结构
输出 JSON: {"ok": true, "file": "...", "paragraphs": 12, "tables": 2}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="MD", help="输入 Markdown")
    sp.add_argument("--out", metavar="DOCX", help="输出 .docx(缺省同名 .docx)")


def run(args: argparse.Namespace) -> dict:
    out = args.out or os.path.splitext(args.file)[0] + ".docx"
    if out.lower() == args.file.lower():
        from ..errors import CliError
        raise CliError("bad_args", "--out 不能与输入相同")
    mdutil.md_to_docx(args.file, out)
    return {"ok": True, "file": out}

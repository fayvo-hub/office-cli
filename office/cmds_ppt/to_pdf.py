"""ppt to-pdf — 演示文稿导出 PDF(经 WPS 演示组件,排版保真)。"""

from __future__ import annotations

import argparse
import os

from .. import pdfutil
from ..cli import add_file_arg
from ..errors import CliError

NAME = "to-pdf"
HELP = "PPT 导出 PDF(需本机 WPS Office)"
DESCRIPTION = """把 .pptx/.ppt 导出为 PDF(经 WPS 演示组件,排版 100% 保真)。

用法:
  office ppt to-pdf -f 汇报.pptx                 # 输出 汇报.pdf(同目录)
  office ppt to-pdf -f 汇报.pptx --out out/a.pdf

说明:
- 需要本机安装 WPS Office(可用 office info 查看 engines.wps)
- 旧版 .ppt 直接导出,无需先转 .pptx
- 转换在子进程执行,最长 180s;WPS 弹窗/卡死不会拖垮命令
输出 JSON: {"ok": true, "file": "...", "pages": N}
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="PPT 文件路径(.pptx/.ppt)")
    sp.add_argument("--out", metavar="PATH", help="输出 PDF 路径(缺省与源文件同名)")


def run(args: argparse.Namespace) -> dict:
    if not os.path.exists(args.file):
        raise CliError("no_file", f"文件不存在: {args.file}")
    if not args.file.lower().endswith((".pptx", ".ppt")):
        raise CliError("unsupported_format",
                       "仅支持 .pptx/.ppt 文件,收到: " + args.file)
    if args.out and not args.out.lower().endswith(".pdf"):
        raise CliError("bad_args", "--out 需以 .pdf 结尾")

    from .. import wps  # 惰性导入

    out = args.out or os.path.splitext(args.file)[0] + ".pdf"
    if os.path.abspath(out).lower() == os.path.abspath(args.file).lower():
        raise CliError("bad_args", "--out 不能与源文件相同")
    wps.convert(args.file, out)

    pages = None
    try:
        reader = pdfutil.open_reader(out)
        pages = len(reader.pages)
        reader.stream.close()
    except Exception:
        pass
    return {"ok": True, "file": out, **({"pages": pages} if pages else {})}

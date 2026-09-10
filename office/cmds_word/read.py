"""word read — 读取 Word 文档(.docx/.doc/.rtf)结构与内容。

- 旧版 .doc 与 .rtf 自动经办公引擎(WPS/LibreOffice)升级后读取(原文件不动)
- 输出段落(含样式/标题层级)、表格、图片清单、字数统计
- 公式域、批注等高级对象不支持(python-docx 能力边界)
"""

from __future__ import annotations

import argparse
import os
import zipfile

from .. import ioplan
from ..cli import add_file_arg
from ..errors import CliError

NAME = "read"
HELP = "读取 Word 文档:段落(含样式层级)/表格/图片/统计"
DESCRIPTION = """读取 Word 文档(.docx/.doc/.rtf),输出结构化 JSON。

输出 JSON:
{
  "ok": true, "file": "...",
  "meta": {"paragraphs": 25, "tables": 2, "images": 1, "chars": 320, "sections": 1},
  "paragraphs": [
    {"idx": 0, "style": "Heading 1", "level": 1, "text": "第一章 概述", "chars": 6},
    {"idx": 1, "style": "Normal", "level": null, "text": "正文段落...", "chars": 18}
  ],
  "tables": [{"index": 0, "rows": 3, "cols": 2, "style": "Table Grid",
              "rows_data": [["表头1", "表头2"], ["a", "1"]]}],
  "images": [{"filename": "image1.png", "size_bytes": 5120}],
  "warnings": []
}

说明:
- 段落按文档顺序输出;style 如 "Heading 1"/"Title"/"Normal" 等;level 仅标题有(1-9)
- 表格单元格取全部文本;合并单元格在 docx 中重复出现是正常现象
- 大文档可用 --limit 截断段落数(默认 500;0 = 不限)
- 图片默认只列清单;加 --save-images DIR 可导出全部图片
- 旧版 .doc 与 .rtf 需要本机 WPS Office 或 LibreOffice 自动升级(OFFICE_ENGINE 可指定)
"""


def register(sp: argparse.ArgumentParser) -> None:
    add_file_arg(sp, help_text="Word 文件路径(.docx;旧版 .doc/.rtf 自动升级)")
    sp.add_argument("--limit", type=int, default=500, metavar="N",
                    help="最多输出段落数(默认 500;0 不限制)")
    sp.add_argument("--tables", action="store_true", default=True,
                    help=argparse.SUPPRESS)  # 保持默认开;表格另行控制
    sp.add_argument("--no-tables", action="store_true", help="不输出表格数据")
    sp.add_argument("--save-images", metavar="DIR",
                    help="把文档内图片导出到目录(默认只列出图片清单)")


def run(args: argparse.Namespace) -> dict:
    plan = ioplan.word_plan(args.file, write=False)
    try:
        from docx import Document
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "word 命令需要 python-docx: pip install python-docx") from None

    try:
        doc = Document(plan.read_path)
    except CliError:
        raise
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 Word 文档 {args.file}: {e}"
                                      f"(文件损坏或不是有效 .docx)") from e

    warnings: list[str] = []
    if plan.upgraded_from:
        _ext = os.path.splitext(args.file)[1].lower()
        _label = "旧版 .doc" if _ext == ".doc" else f" {_ext} 格式"
        warnings.append(f"{args.file} 为{_label},已由办公引擎自动升级后读取(原文件未改动)")

    limit = args.limit if args.limit and args.limit > 0 else None
    paragraphs = []
    total_chars = 0
    for i, p in enumerate(doc.paragraphs):
        text = p.text
        style_name = p.style.name if p.style is not None else "Normal"
        level = None
        if style_name.lower().startswith("heading") and style_name[-1:].isdigit():
            level = int(style_name.split()[-1])
        elif style_name.lower() == "title":
            level = 0
        paragraphs.append({"idx": i, "style": style_name, "level": level,
                           "text": text, "chars": len(text)})
        total_chars += len(text)

    if limit is not None and len(paragraphs) > limit:
        paragraphs = paragraphs[:limit]
        warnings.append(f"段落数超过 --limit {limit},已截断(需要全部请加 --limit 0)")

    tables = []
    if not args.no_tables:
        for ti, tb in enumerate(doc.tables):
            data = [[cell.text for cell in row.cells] for row in tb.rows]
            tables.append({"index": ti, "rows": len(data),
                           "cols": (len(data[0]) if data else 0),
                           "style": (tb.style.name if tb.style is not None else None),
                           "rows_data": data})

    images = []
    saved_images = None
    try:
        with zipfile.ZipFile(plan.read_path) as zf:
            for name in zf.namelist():
                if name.startswith("word/media/") and not name.endswith("/"):
                    info = zf.getinfo(name)
                    images.append({"filename": os.path.basename(name),
                                   "size_bytes": info.file_size})
            if args.save_images:
                saved_images = []
                os.makedirs(args.save_images, exist_ok=True)
                for name in zf.namelist():
                    if name.startswith("word/media/") and not name.endswith("/"):
                        out = os.path.join(args.save_images, os.path.basename(name))
                        with zf.open(name) as src, open(out, "wb") as dst:
                            dst.write(src.read())
                        saved_images.append(out)
    except zipfile.BadZipFile:
        pass  # 理论不可达(docx 必为 zip)

    return {"ok": True, "file": args.file, "meta": {
                "paragraphs": len(paragraphs), "tables": len(tables),
                "images": len(images), "chars": total_chars},
            "paragraphs": paragraphs, "tables": tables, "images": images,
            "saved_images": saved_images, "warnings": warnings}

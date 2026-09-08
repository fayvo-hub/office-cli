"""image — 向表中插入图片(需要 Pillow)。"""

from __future__ import annotations

import argparse
import os

from .. import ioplan, xlutil
from ..cli import add_file_arg, add_sheet_arg
from ..errors import CliError

NAME = "image"
HELP = "插入图片到工作表(需要 Pillow:pip install Pillow)"
DESCRIPTION = """把图片文件插入到工作表的指定位置。

用法示例:
  xcli image add -f demo.xlsx --image logo.png --at B2
  xcli image add -f demo.xlsx --image photo.jpg --at C5 --width 300   # 按像素等比缩放
  xcli image add -f demo.xlsx --image pic.png --at A1 --scale 0.5     # 按比例缩放

说明:
- 需要 Pillow 支持(pip install Pillow);png/jpg/gif/webp 等常见格式均可
- --width / --scale 二选一;不缩放时保持原图尺寸
输出 JSON: {"ok": true, "file": "...", "sheet": "...", "image": "logo.png",
            "anchor": "B2", "width_px": 120, "height_px": 60}
"""


def register(sp: argparse.ArgumentParser) -> None:
    sub = sp.add_subparsers(dest="image_action", metavar="动作", required=True)
    s = sub.add_parser("add", help="插入图片")
    add_file_arg(s)
    add_sheet_arg(s)
    s.add_argument("--image", required=True, metavar="PATH", help="图片文件路径")
    s.add_argument("--at", default="A1", metavar="REF", help="锚点单元格,如 B2(默认 A1)")
    size = s.add_mutually_exclusive_group()
    size.add_argument("--width", type=int, metavar="PX", help="目标宽度(像素,高度等比)")
    size.add_argument("--scale", type=float, metavar="RATIO", help="缩放比例,如 0.5")
    s.set_defaults(image_func="add")


def run(args: argparse.Namespace) -> dict:
    if args.image_func != "add":  # pragma: no cover
        raise CliError("internal", f"未知动作 {args.image_func}")

    if not os.path.exists(args.image):
        raise CliError("no_file", f"图片不存在: {args.image}")

    try:
        from PIL import Image as PILImage
    except ImportError:
        raise CliError(
            "need_pillow",
            "插入图片需要 Pillow 库。请先安装: pip install Pillow",
        ) from None

    try:
        from openpyxl.drawing.image import Image as XLImage
    except ImportError:  # pragma: no cover
        raise CliError("internal", "当前 openpyxl 不含 drawing.image 支持,请升级 openpyxl") from None

    plan = ioplan.excel_plan(args.file)
    wb = xlutil.open_workbook(plan.read_path)
    ws = xlutil.choose_sheet(wb, args.sheet)
    xlutil.parse_range(args.at, ws)  # 校验锚点

    try:
        with PILImage.open(args.image) as im:
            orig_w, orig_h = im.size
    except Exception as e:
        raise CliError("bad_image", f"无法读取图片 {args.image}: {type(e).__name__}: {e}") from e

    if args.scale is not None:
        if args.scale <= 0:
            raise CliError("bad_args", "--scale 必须大于 0")
        width = int(orig_w * args.scale)
        height = int(orig_h * args.scale)
    elif args.width is not None:
        if args.width <= 0:
            raise CliError("bad_args", "--width 必须大于 0")
        width = args.width
        height = max(1, int(orig_h * (args.width / orig_w)))
    else:
        width, height = orig_w, orig_h

    xlimg = XLImage(args.image)
    xlimg.width, xlimg.height = width, height
    ws.add_image(xlimg, args.at)

    xlutil.save_workbook_atomic(wb, plan.write_path)
    _up = {"upgraded_from": plan.upgraded_from} if plan.upgraded_from else {}
    return {"ok": True, "file": plan.write_path, **_up, "sheet": ws.title,
            "image": os.path.basename(args.image), "anchor": args.at,
            "width_px": width, "height_px": height}

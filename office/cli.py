"""CLI 入口:两级路由(excel/word/pdf/md 组 + convert/info 顶层)、统一 JSON 输出、统一错误处理。

命令形态:
  office excel list/read/write/sheet/style/merge/chart/image/pivot ...
  office word read/write ...
  office pdf info/read/merge/split/rotate/encrypt/decrypt/watermark/images/to-image ...
  office md to-pdf/to-docx/to-html ...
  office convert -f a.md --out a.pdf           (跨格式转换,按扩展名自动路由)
  office info                                   (环境自检,供 AI 排错)
"""

from __future__ import annotations

import argparse
import importlib
import json
import os
import sys
import traceback

from . import __version__
from .errors import CliError

# ---------------------------------------------------------------------------
# 输出与错误协议:
# - 成功: JSON -> stdout,退出码 0
# - 失败: {"error": {"code": ..., "message": ...}} -> stderr,退出码非 0
#   错误信息为中文、可操作,AI 应能根据 message 自行修正后重试
# ---------------------------------------------------------------------------

def _emit(result: dict) -> None:
    print(json.dumps(result, ensure_ascii=False, indent=2))


def _emit_error(code: str, message: str) -> None:
    print(json.dumps({"error": {"code": code, "message": message}},
                     ensure_ascii=False, indent=2), file=sys.stderr)


def add_file_arg(sp: argparse.ArgumentParser, help_text: str | None = None) -> None:
    sp.add_argument("-f", "--file", required=True, metavar="PATH",
                    help=help_text or "文件路径(.xlsx/.xlsm/.xls;旧版 .xls 自动升级后操作)")


def add_sheet_arg(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--sheet", metavar="NAME",
                    help="工作表名;缺省用文件打开时激活的表(可用 list 命令查看所有表名)")


# 各组内子命令对应的模块名(与文件名一致)
_GROUP_MODS = {
    "excel": ["list", "read", "write", "sheet", "style", "merge", "chart",
              "image", "pivot"],
    "word": ["read", "write"],
    "pdf": ["info", "read", "merge", "split", "rotate", "encrypt", "decrypt",
            "watermark", "images", "to-image"],
    "md": ["to-pdf", "to-docx", "to-html"],
}


def _add_group(sub, group: str) -> None:
    """注册一个命令组(如 office excel <子命令>)。组内模块未就绪时自动跳过。"""
    gsp = sub.add_parser(group, help=f"{group} 相关命令(office {group} <子命令> --help)")
    gsub = gsp.add_subparsers(dest="cmd", metavar="子命令", required=True)
    for name in _GROUP_MODS[group]:
        try:
            mod = importlib.import_module(f".cmds_{group}.{name.replace('-', '_')}",
                                          package=__package__)
        except ModuleNotFoundError:
            continue  # 该子命令模块尚未就绪(开发中),跳过
        sp = gsub.add_parser(mod.NAME, help=mod.HELP, description=mod.DESCRIPTION,
                             formatter_class=argparse.RawDescriptionHelpFormatter)
        mod.register(sp)
        sp.set_defaults(func=mod.run)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="office",
        description="面向 AI 调用的办公文档 CLI:Excel/Word/PDF/Markdown 读写与互转。"
                    "所有命令输出 JSON;失败时向 stderr 输出 {\"error\": {...}} 并以非零码退出。",
    )
    parser.add_argument("--version", action="version", version=f"office {__version__}")
    sub = parser.add_subparsers(dest="command", metavar="命令", required=True)

    # 顶层扁平命令:convert(跨格式转换)、info(环境自检)
    for name in ("convert", "info"):
        mod = importlib.import_module(f".{name}", package=__package__)
        sp = sub.add_parser(mod.NAME, help=mod.HELP, description=mod.DESCRIPTION,
                            formatter_class=argparse.RawDescriptionHelpFormatter)
        mod.register(sp)
        sp.set_defaults(func=mod.run)

    for group in _GROUP_MODS:
        _add_group(sub, group)
    return parser


def main(argv: list[str] | None = None) -> int:
    # PyInstaller 单文件模式:office.exe 被主进程以 "-m office.wps_worker" 参数
    # 再次拉起时,充当 WPS COM 子进程(与 pip 版 subprocess 行为一致)
    if getattr(sys, "frozen", False) and list(sys.argv[1:3]) == ["-m", "office.wps_worker"]:
        from . import wps_worker

        return wps_worker.worker_main(sys.argv[3] if len(sys.argv) > 3 else "{}")

    # Windows 控制台可能是 GBK,强制 UTF-8 输出,避免中文/特殊字符编码崩溃
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        result = args.func(args)
    except CliError as e:
        _emit_error(e.code, e.message)
        return e.exit_code
    except KeyboardInterrupt:
        _emit_error("interrupted", "用户中断")
        return 130
    except Exception as e:  # 兜底:第三方库内部错误等
        if os.environ.get("OFFICE_DEBUG"):
            traceback.print_exc()
        _emit_error("internal", f"{type(e).__name__}: {e}")
        return 4

    if result is not None:
        _emit(result)
    return 0

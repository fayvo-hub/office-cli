"""输入/输出路径规划:旧版 .xls/.doc 自动升级为 .xlsx/.docx 后再操作。

规则(对用户透明):
- .xlsx/.docx 等新格式:原地读写
- .xls/.doc 老格式:先用本机 WPS Office 无损升级到临时新格式文件供读取;
  写操作的结果保存为同目录同名的新格式文件(如 报表.xls -> 报表.xlsx),
  原老格式文件一律不动(可随时用原文件重新生成)。
"""

from __future__ import annotations

import os
import tempfile

from .errors import CliError

_cache: dict[tuple[str, str], str] = {}


class IoPlan:
    """一次文件操作的计划:
    read_path     实际打开读取的路径(老格式时为升级后的临时文件)
    write_path    保存目标(老格式时为同目录同名 .xlsx/.docx;只读为 None)
    upgraded_from 若输入为老格式,其原始路径;否则 None
    """

    __slots__ = ("read_path", "write_path", "upgraded_from")

    def __init__(self, read_path: str, write_path: str | None = None,
                 upgraded_from: str | None = None) -> None:
        self.read_path = read_path
        self.write_path = write_path
        self.upgraded_from = upgraded_from


def excel_plan(path: str, *, write: bool = True) -> IoPlan:
    """Excel 文件计划:.xls 老格式 -> 升级 .xlsx。"""
    ext = os.path.splitext(path)[1].lower()
    if ext != ".xls":
        return IoPlan(path, path if write else None)
    tmp = _upgrade(path, ".xls", ".xlsx", kind="xls")
    final = os.path.splitext(path)[0] + ".xlsx" if write else None
    return IoPlan(tmp, final, path)


def word_plan(path: str, *, write: bool = True) -> IoPlan:
    """Word 文件计划:.doc 老格式 -> 升级 .docx。"""
    ext = os.path.splitext(path)[1].lower()
    if ext != ".doc":
        return IoPlan(path, path if write else None)
    tmp = _upgrade(path, ".doc", ".docx", kind="doc")
    final = os.path.splitext(path)[0] + ".docx" if write else None
    return IoPlan(tmp, final, path)


def ppt_plan(path: str, *, write: bool = True) -> IoPlan:
    """PPT 文件计划:.ppt 老格式 -> 升级 .pptx。"""
    ext = os.path.splitext(path)[1].lower()
    if ext != ".ppt":
        return IoPlan(path, path if write else None)
    tmp = _upgrade(path, ".ppt", ".pptx", kind="ppt")
    final = os.path.splitext(path)[0] + ".pptx" if write else None
    return IoPlan(tmp, final, path)


def _upgrade(src: str, src_ext: str, dst_ext: str, kind: str) -> str:
    """把老格式文件升级为同内容的新格式临时文件(进程内缓存)。"""
    abs_src = os.path.abspath(src)
    key = (kind, abs_src)
    if key in _cache and os.path.exists(_cache[key]):
        return _cache[key]
    if not os.path.exists(src):
        raise CliError("no_file", f"文件不存在: {src}")
    from . import wps  # 惰性导入,避免无 WPS 环境下拖慢启动

    if not wps.available():
        raise CliError(
            "wps_unavailable",
            f"'{src}' 是旧版 {src_ext} 格式,需要本机 WPS Office 自动升级后才能操作,"
            f"但 WPS COM 不可用。请安装/打开 WPS Office 后重试,"
            f"或先用 WPS 手动另存为 {dst_ext} 文件。",
        )
    fd, tmp = tempfile.mkstemp(prefix=f".office-{kind}-", suffix=dst_ext)
    os.close(fd)
    os.remove(tmp)
    try:
        wps.convert(src, tmp)
    except CliError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    _cache[key] = tmp
    return tmp

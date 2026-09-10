# -*- coding: utf-8 -*-
"""办公转换引擎门面: WPS Office(COM) 与 LibreOffice(无头) 之间自动选择。

选择规则(环境变量 `OFFICE_ENGINE`,默认 auto):
- `auto`   : Windows 优先 WPS(常驻 COM 更快,PDF 与 Office 更像),不可用则回退 LibreOffice;
             macOS/Linux 优先 LibreOffice。两者都没有时报错并给出安装指引。
- `wps`    : 只用 WPS;
- `lo`     : 只用 LibreOffice(跨平台一致);
- `none`   : 关闭外部引擎(老格式升级/PDF 导出会报 engine_unavailable)。

为什么两个后端: WPS/Office 的 COM 在 Windows 上快且保真,但不跨平台;
LibreOffice 跨平台(含容器/服务器),语义等价(实测 .doc→.docx、.xls→.xlsx、
.doc/.docx/.ppt/.pptx→.pdf、公式重算均一致),代价是每次冷启动数秒、
复杂版式 PDF 有细微差异、Linux 容器需自备中文字体。

调用方只依赖本模块(不再直接 import wps/lo),引擎切换对上层透明。
"""

from __future__ import annotations

import os

from . import lo, wps
from .errors import CliError

_DEFAULT_TIMEOUT = 180
ENV_VAR = "OFFICE_ENGINE"

_DISPLAY = {"wps": "WPS", "lo": "LibreOffice", "none": "none"}

_HINT = (
    "本机没有可用的办公转换引擎(老格式升级与 Office→PDF 需要其一):\n"
    "  Windows: 装 WPS Office(最快)或 LibreOffice —— winget install "
    "TheDocumentFoundation.LibreOffice\n"
    "  macOS:   brew install --cask libreoffice\n"
    "  Linux:   sudo apt install libreoffice(容器里另装 fonts-noto-cjk 中文字体)\n"
    "也可用环境变量 OFFICE_ENGINE=wps|lo 指定引擎"
)


def _explicit() -> str | None:
    """读取 OFFICE_ENGINE 显式设置(auto/None 表示未指定)。"""
    raw = (os.environ.get(ENV_VAR) or "").strip().lower()
    if raw in ("wps", "lo", "libreoffice", "none"):
        return "lo" if raw == "libreoffice" else raw
    return None


def _has(name: str) -> bool:
    return wps.available() if name == "wps" else lo.available()


def order() -> list[str]:
    """auto 模式的候选顺序。"""
    return ["wps", "lo"] if os.name == "nt" else ["lo", "wps"]


def active() -> str | None:
    """当前生效的后端标识('wps' | 'lo' | None)。"""
    exp = _explicit()
    if exp == "none":
        return None
    if exp in ("wps", "lo"):
        return exp if _has(exp) else None
    for name in order():
        if _has(name):
            return name
    return None


def name() -> str:
    """当前生效的引擎名(用于输出文案): WPS / LibreOffice / none。"""
    return _DISPLAY[active() or "none"]


def available() -> bool:
    """是否有可用引擎(与历史 wps.available() 语义一致,便于平滑替换)。"""
    return active() is not None


def convert(src: str, dst: str, timeout: int = _DEFAULT_TIMEOUT) -> None:
    """格式转换: 老格式升级(.xls/.doc/.rtf/.ppt → 新格式)或导出 PDF。"""
    n = active()
    if n is None:
        raise CliError("engine_unavailable", _HINT)
    (wps if n == "wps" else lo).convert(src, dst, timeout)


def recalc(src: str, dst: str, timeout: int = _DEFAULT_TIMEOUT) -> None:
    """用办公引擎重算 Excel 全部公式后另存 dst(取公式真值,源文件不动)。"""
    n = active()
    if n is None:
        raise CliError("engine_unavailable", _HINT)
    (wps if n == "wps" else lo).recalc(src, dst, timeout)


def is_legacy(path: str, ext: str | None = None) -> bool:
    """是否为旧版办公格式(.xls/.doc/.ppt,需要引擎升级)。"""
    e = (ext or os.path.splitext(path)[1]).lower()
    return e in (".xls", ".doc", ".ppt")


def info() -> dict:
    """引擎探测结果(供 `office info` 展示)。"""
    return {
        "active": name(),
        "preference": _explicit() or "auto",
        "wps": wps.available(),
        "libreoffice": lo.available(),
        "libreoffice_path": lo.path(),
    }

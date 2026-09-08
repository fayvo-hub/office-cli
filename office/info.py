"""info — 环境自检:依赖库与引擎可用性,供 AI/用户排错。

用法: office info [--fast]   (--fast 跳过 WPS COM 探测,立即返回)
"""

from __future__ import annotations

import argparse
import importlib.util
import os
import platform
import sys

from . import __version__

NAME = "info"
HELP = "环境自检:Python/依赖库/转换引擎可用性(供排错)"
DESCRIPTION = """环境自检:输出本机 Python、各依赖库、转换引擎(WPS/Chrome/Edge)的可用状态。

输出 JSON:
{
  "ok": true, "tool": "office", "version": "0.2.0",
  "python": "D:/Python313/python.exe 3.13.2",
  "platform": "Windows-...",
  "modules": {"openpyxl": true, "docx": true, ...},
  "engines": {
    "wps": true,                      // 本机 WPS Office COM 可用(旧格式/.docx->pdf 依赖)
    "chrome": "C:/.../chrome.exe",    // md->pdf 渲染浏览器
    "edge": "C:/.../msedge.exe",
    "mermaid_assets": true            // md->pdf 的 mermaid 渲染支持文件
  }
}

当某个命令报缺依赖时,先跑 office info 确认环境,再对症处理。
"""

_MODULES = (
    "openpyxl", "docx", "fitz", "pdfplumber", "pypdf", "xlrd",
    "pdf2docx", "markdown", "pygments", "playwright", "win32com",
    "PIL", "pandas",
)

_CHROME_PATHS = (
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
)
_EDGE_PATHS = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
)


def register(sp: argparse.ArgumentParser) -> None:
    sp.add_argument("--fast", action="store_true",
                    help="跳过 WPS COM 探测(约省数秒;引擎项输出 null)")


def run(args: argparse.Namespace) -> dict:
    modules = {m: importlib.util.find_spec(m) is not None for m in _MODULES}

    engines: dict = {"chrome": None, "edge": None, "mermaid_assets": False}
    for p in _CHROME_PATHS:
        if os.path.exists(p):
            engines["chrome"] = p
            break
    for p in _EDGE_PATHS:
        if os.path.exists(p):
            engines["edge"] = p
            break
    assets = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                          "assets", "mermaid.min.js")
    engines["mermaid_assets"] = os.path.exists(assets)
    if args.fast:
        engines["wps"] = None
    else:
        try:
            from . import wps

            engines["wps"] = wps.available()
        except Exception as e:  # noqa: BLE001
            engines["wps"] = False
            engines["wps_error"] = str(e)

    missing = [k for k, v in modules.items() if not v]
    return {
        "ok": True,
        "tool": "office",
        "version": __version__,
        "python": f"{sys.executable} {platform.python_version()}",
        "platform": platform.platform(),
        "modules": modules,
        "engines": engines,
        "missing": missing,
        "tip": ("全部就绪" if not missing and engines.get("wps") is not False
                else f"缺失库: {missing or '无'};"
                     f"WPS 状态: {engines.get('wps')} —— 按需 pip install 或安装 WPS Office"),
    }

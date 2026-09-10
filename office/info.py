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
    "active": "WPS",                  // 当前生效的转换引擎(WPS/LibreOffice/none)
    "wps": true,                      // 本机 WPS Office COM 可用
    "libreoffice": "C:/Program Files/LibreOffice/program/soffice.exe",  // 跨平台引擎
    "chrome": "C:/.../chrome.exe",    // md->pdf 渲染浏览器
    "edge": "C:/.../msedge.exe",
    "mermaid_assets": true,           // md->pdf 的 mermaid 渲染支持文件
    "printer": {                      // 默认打印机隔离状态(Windows;引擎启动时防连网络打印机)
      "isolate_enabled": true,
      "default": "HP LaserJet ...",
      "default_port": "WSD-...",      // WSD-/IP_/UNC 等网络端口
      "default_is_network": true,
      "virtual_target": "Microsoft Print to PDF"
    }
  }
}

引擎说明: 旧格式升级(.xls/.doc/.ppt)与 Office→PDF 依赖 WPS 或 LibreOffice 其一,
可用环境变量 OFFICE_ENGINE=auto|wps|lo|none 指定;详见 README "引擎"一节。

当某个命令报缺依赖时,先跑 office info 确认环境,再对症处理。
"""

_MODULES = (
    "openpyxl", "docx", "fitz", "pdfplumber", "pypdf",
    "pdf2docx", "markdown", "pygments", "playwright", "win32com",
    "PIL",
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
    from . import engine, lo

    engines["libreoffice"] = lo.path()
    engines["printer"] = engine.info().get("printer")
    if args.fast:
        engines["wps"] = None
        engines["active"] = "LibreOffice" if lo.available() else None
    else:
        try:
            from . import wps

            engines["wps"] = wps.available()
        except Exception as e:  # noqa: BLE001
            engines["wps"] = False
            engines["wps_error"] = str(e)
        engines["active"] = (engine.name() if engine.available() else "none")
        engines["preference"] = engine.info().get("preference")

    missing = [k for k, v in modules.items() if not v]
    mode = "exe" if getattr(sys, "frozen", False) else "pip"
    has_engine = bool(engines.get("wps") or engines.get("libreoffice"))
    return {
        "ok": True,
        "tool": "office",
        "version": __version__,
        "mode": mode,
        "python": f"{sys.executable} {platform.python_version()}",
        "platform": platform.platform(),
        "modules": modules,
        "engines": engines,
        "missing": missing,
        "tip": ("全部就绪" if not missing and has_engine
                else f"缺失库: {missing or '无'};引擎: "
                     f"WPS={engines.get('wps')} LibreOffice={engines.get('libreoffice')}"
                     f" —— 按需 pip install,或安装 WPS Office / LibreOffice"
                     f"(winget install TheDocumentFoundation.LibreOffice)"),
    }

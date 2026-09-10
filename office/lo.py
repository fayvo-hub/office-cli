# -*- coding: utf-8 -*-
"""LibreOffice 无头后端(跨平台: Windows / macOS / Linux)。

与 WPS COM 后端(wps.py)等价的能力:
  .xls  -> .xlsx      .doc/.rtf -> .docx      .ppt -> .pptx
  .doc/.docx/.ppt/.pptx -> .pdf(打印级)
  xlsx/xlsm 公式重算(profile 预置 OOXMLRecalcMode=0 → 打开即重算,另存写出新值)

实现要点:
- 每次调用独立子进程 `soffice --headless --convert-to ...`,无 COM、无窗口;
- 复用固定 user profile(预置重算策略),冷启动 1.5~6s,之后约 1.4s/次;
- 产物先落到临时目录再移动到目标(soffice 只能指定输出目录,不能指定文件名);
- 定位顺序: 环境变量 LIBREOFFICE_SOFFICE / SOFFICE_PATH → PATH → 各平台常见安装路径。

与前端的差异(相对 WPS): 复杂版式 PDF 与 Office 存在细微差异,Linux 容器里需自备中文字体
(如 fonts-noto-cjk),否则中文可能显示为方框。
"""

from __future__ import annotations

import glob
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from .errors import CliError

_DEFAULT_TIMEOUT = 180
_PATH_CACHE: str | None | bool = None      # None=未探测; False=不可用; str=可执行文件

_WIN_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
)
_MAC_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
)
_POSIX_CANDIDATES = (
    "/usr/bin/soffice", "/usr/local/bin/soffice", "/snap/bin/libreoffice",
    "/usr/lib/libreoffice/program/soffice", "/opt/libreoffice/program/soffice",
)

# 目标扩展名 → soffice 转换过滤器名(简写由 LO 自动选默认过滤器)
_FILTER = {".xlsx": "xlsx", ".docx": "docx", ".pptx": "pptx", ".pdf": "pdf"}
_SUPPORTED = {(".xls", ".xlsx"), (".doc", ".docx"), (".rtf", ".docx"),
              (".ppt", ".pptx"), (".doc", ".pdf"), (".docx", ".pdf"),
              (".ppt", ".pdf"), (".pptx", ".pdf")}

# 强制"打开即重算"(0=总是重算, 1=从不, 2=询问用户),否则带缓存的表不会被刷新
_PROFILE_XCU = """<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" \
xmlns:xs="http://www.w3.org/2001/XMLSchema" \
xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load">\
<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load">\
<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>
</oor:items>
"""


def _looks_like_error(line: str) -> bool:
    """soffice 输出中是否为错误行(用于挑出更有用的报错信息)。"""
    low = line.lower()
    return "error" in low or "could not" in low or "failed" in low


def path() -> str | None:
    """返回 soffice 可执行文件路径(进程内缓存);未安装返回 None。"""
    global _PATH_CACHE
    if _PATH_CACHE is None:
        cands: list[str] = []
        for var in ("LIBREOFFICE_SOFFICE", "SOFFICE_PATH"):
            v = os.environ.get(var)
            if v:
                cands.append(v)
        found = shutil.which("soffice") or shutil.which("libreoffice")
        if found:
            cands.append(found)
        if sys.platform == "darwin":
            cands += list(_MAC_CANDIDATES)
        elif os.name == "nt":
            cands += list(_WIN_CANDIDATES)
        else:
            cands += list(_POSIX_CANDIDATES)
        cands += sorted(glob.glob("/opt/libreoffice*/program/soffice"))
        _PATH_CACHE = next((c for c in cands if c and os.path.isfile(c)), False)
    return _PATH_CACHE or None


def available() -> bool:
    """LibreOffice 是否可用(只探测可执行文件,不启动进程)。"""
    return path() is not None


def _profile_dir() -> str:
    """固定 user profile(首次创建并写入"总是重算"策略)。"""
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    else:
        base = os.environ.get("XDG_CACHE_HOME") or os.path.join(os.path.expanduser("~"), ".cache")
    prof = os.path.join(base, "office-cli", "lo-profile")
    xcu = os.path.join(prof, "user", "registrymodifications.xcu")
    if not os.path.exists(xcu):
        os.makedirs(os.path.dirname(xcu), exist_ok=True)
        with open(xcu, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(_PROFILE_XCU)
    return prof


def _invoke(args: list[str], timeout: int) -> subprocess.CompletedProcess:
    exe = path()
    if exe is None:
        raise CliError("engine_unavailable",
                       "未找到 LibreOffice(soffice)。请安装 LibreOffice,"
                       "或用 LIBREOFFICE_SOFFICE 指定 soffice 路径。")
    env = dict(os.environ)
    for var in ("PYTHONHOME", "PYTHONPATH"):   # 避免 LO 自带 python 被宿主解释器干扰
        env.pop(var, None)
    cmd = [exe,
           f"-env:UserInstallation={Path(_profile_dir()).as_uri()}",
           "--headless", "--norestore", "--nolockcheck", "--nodefault",
           "--nologo", "--nofirststartwizard", *args]
    try:
        return subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        raise CliError(
            "lo_timeout",
            f"LibreOffice 转换超过 {timeout}s 未完成(文件过大或引擎异常)。"
            f"可加大超时重试,或先用 LibreOffice 手动转换该文件。",
        ) from None
    except OSError as exc:
        raise CliError("lo_failed", f"无法启动 LibreOffice: {exc}") from None


def _convert(src: str, dst: str, dst_ext: str, timeout: int) -> None:
    """调 soffice 转成 dst_ext,产物移动到 dst。"""
    out_dir = tempfile.mkdtemp(prefix="office-lo-out-")
    try:
        proc = _invoke(["--convert-to", _FILTER[dst_ext], "--outdir", out_dir,
                        os.path.abspath(src)], timeout)
        produced = os.path.join(
            out_dir, os.path.splitext(os.path.basename(src))[0] + dst_ext)
        deadline = time.time() + 20      # soffice 偶有"进程先退、写入稍后"的窗口
        while not os.path.exists(produced) and time.time() < deadline:
            time.sleep(0.3)
        if not os.path.exists(produced):
            lines = [ln.strip() for ln in (proc.stdout or "").splitlines()
                     + (proc.stderr or "").splitlines() if ln.strip()]
            detail = next((ln for ln in lines if _looks_like_error(ln)), "")
            detail = detail or (lines[-1] if lines else "无输出")
            raise CliError("lo_failed",
                           f"LibreOffice 转换失败: {detail[:300]}(源文件: {src})")
        try:
            if os.path.exists(dst):
                os.remove(dst)
            shutil.move(produced, dst)
        except OSError as exc:
            raise CliError("write_failed", f"转换产物写入 {dst} 失败: {exc}") from None
    finally:
        shutil.rmtree(out_dir, ignore_errors=True)


def convert(src: str, dst: str, timeout: int = _DEFAULT_TIMEOUT) -> None:
    """按扩展名转换(与 wps.convert 同签名同语义)。"""
    s_ext = os.path.splitext(src)[1].lower()
    d_ext = os.path.splitext(dst)[1].lower()
    if (s_ext, d_ext) not in _SUPPORTED:
        raise CliError(
            "unsupported",
            f"LibreOffice 转换不支持 {s_ext} -> {d_ext};支持: .xls->.xlsx、"
            f".doc/.rtf->.docx、.ppt->.pptx、.doc/.docx/.ppt/.pptx->.pdf",
        )
    if not os.path.exists(src):
        raise CliError("no_file", f"文件不存在: {src}")
    _convert(src, dst, d_ext, timeout)


def recalc(src: str, dst: str, timeout: int = _DEFAULT_TIMEOUT) -> None:
    """重算全部公式并另存为 dst(源文件不动)。

    profile 里已设 OOXMLRecalcMode=0: LibreOffice 打开工作簿时重算,保存时写出新值。
    """
    ext = os.path.splitext(src)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls"):
        raise CliError("unsupported",
                       f"LibreOffice 重算只支持 Excel 工作簿,收到 {ext}")
    if not os.path.exists(src):
        raise CliError("no_file", f"文件不存在: {src}")
    _convert(src, dst, ".xlsx", timeout)

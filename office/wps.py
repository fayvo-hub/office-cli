"""WPS Office COM 桥(主进程侧):探测可用性、后台子进程转换,防 COM 卡死主 CLI。

所有转换都在独立子进程(office.wps_worker)内执行,主进程 subprocess 等待并设超时;
WPS 崩溃/卡死不会拖垮命令行。转换过程不弹窗(Visible=False)。

支持的转换(依赖本机装有 WPS Office):
  .xls  -> .xlsx      (KET 表格)
  .doc  -> .docx      (KWPS 文字)
  .docx -> .pdf       (KWPS 导出 PDF,排版 100% 保真)
  .doc  -> .pdf       (内部先转 docx)
"""

from __future__ import annotations

import json
import os
import subprocess
import sys

from .errors import CliError

_DEFAULT_TIMEOUT = 180

_probe_cache: bool | None = None
_probe_cache_checked: bool | None = None

# 转换完成后原样保留(供本进程内后续步骤复用)的临时文件前缀
_KEEP_PREFIXES = (".office-xls-", ".office-doc-")


def _run_worker(job: dict, timeout: int = _DEFAULT_TIMEOUT) -> dict:
    """在子进程执行一次转换,返回 worker 的 JSON 结果;超时/失败抛 CliError。"""
    payload = json.dumps(job, ensure_ascii=False)
    try:
        proc = subprocess.run(
            [sys.executable, "-m", "office.wps_worker", payload],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        raise CliError(
            "wps_timeout",
            f"WPS 转换超过 {timeout}s 未完成(文件过大或 WPS 异常)。"
            f"请检查 WPS 是否弹出对话框,或先用 WPS 手动转换该文件。",
        ) from None
    try:
        result = json.loads(proc.stdout.strip() or "{}")
    except json.JSONDecodeError:
        result = {}
    if proc.returncode != 0 or not result.get("ok"):
        msg = result.get("error") or proc.stderr.strip()[:300] or "未知错误"
        raise CliError(
            "wps_failed",
            f"WPS 转换失败: {msg}(请确认文件未损坏、WPS 未弹窗;"
            f"可用 office info 查看 WPS 是否可用)",
        )
    return result


def available() -> bool:
    """WPS 是否可用(进程内缓存结果,首次探测约需数秒)。"""
    global _probe_cache
    if _probe_cache is None:
        try:
            _run_worker({"op": "probe"}, timeout=40)
            _probe_cache = True
        except CliError:
            _probe_cache = False
    return _probe_cache


def convert(src: str, dst: str, timeout: int = _DEFAULT_TIMEOUT) -> None:
    """按扩展名自动选择转换 op;目标已存在则覆盖。"""
    s_ext = os.path.splitext(src)[1].lower()
    d_ext = os.path.splitext(dst)[1].lower()
    op = {
        (".xls", ".xlsx"): "xls_to_xlsx",
        (".doc", ".docx"): "doc_to_docx",
        (".docx", ".pdf"): "docx_to_pdf",
        (".doc", ".pdf"): "doc_to_pdf",
    }.get((s_ext, d_ext))
    if op is None:
        raise CliError(
            "unsupported",
            f"WPS 转换不支持 {s_ext} -> {d_ext};"
            f"支持: .xls->.xlsx、.doc->.docx、.doc/.docx->.pdf",
        )
    if not os.path.exists(src):
        raise CliError("no_file", f"文件不存在: {src}")
    _run_worker({"op": op, "src": src, "dst": dst}, timeout=timeout)


def is_legacy(path: str, ext: str | None = None) -> bool:
    """是否为旧版办公格式(.xls/.doc)"""
    e = (ext or os.path.splitext(path)[1]).lower()
    return e in (".xls", ".doc")

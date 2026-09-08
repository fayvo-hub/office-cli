"""PDF 命令公共工具:打开/原子写出/页码范围解析/水印。"""

from __future__ import annotations

import os
import re
import tempfile

from .errors import CliError

_PAGES_RE = re.compile(r"^\d+(-\d+)?$")


def open_reader(path: str):
    """打开 PDF 供读取(pypdf PdfReader);统一错误。"""
    if not os.path.exists(path):
        raise CliError("no_file", f"文件不存在: {path}")
    try:
        from pypdf import PdfReader
    except ImportError:  # pragma: no cover
        raise CliError("need_dep", "需要 pypdf 库: pip install pypdf") from None
    try:
        return PdfReader(path, strict=False)
    except Exception as e:
        raise CliError("cannot_open", f"无法打开 PDF {path}: {e}"
                                      f"(文件损坏或不是有效 PDF)") from e


def ensure_decrypted(reader) -> None:
    """加密 PDF 且未解锁时给可操作错误。"""
    if reader.is_encrypted:
        raise CliError("encrypted",
                       "PDF 已加密,需要密码。请先执行 office pdf decrypt -f 文件 --password XXX")


def write_atomic(path: str, do_write) -> dict:
    """原子写出:先写同目录临时文件再替换,原文件中途损坏风险为零。

    do_write(fileobj) 负责把内容写入打开的临时文件。
    """
    directory = os.path.dirname(os.path.abspath(path)) or "."
    name = os.path.basename(path)
    fd, tmp = tempfile.mkstemp(prefix=".office-pdf-", suffix=name, dir=directory)
    try:
        with os.fdopen(fd, "wb") as fh:
            do_write(fh)
        os.replace(tmp, path)
    except PermissionError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise CliError("file_busy",
                       f"无法写入 {path}: 文件正被 PDF 阅读器/杀软占用,请关闭后重试") from None
    except Exception:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    return {"ok": True, "file": path}


def save_writer(writer, path: str) -> None:
    """把 pypdf PdfWriter 原子写出。"""
    write_atomic(path, writer.write)


def parse_pages(text: str | None, page_count: int) -> list[int]:
    """解析页范围:None/'all'/'1-3,5,8-' → 0-based 页号列表(升序去重)。"""
    if text is None or str(text).strip().lower() in ("", "all"):
        return list(range(page_count))
    out: list[int] = []
    for part in str(text).split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            lo = int(a) if a else 1
            hi = int(b) if b else page_count
        else:
            lo = hi = int(part)
        if lo < 1 or hi > page_count:
            raise CliError("bad_args",
                           f"页范围 {text} 超出文档页数 1-{page_count}")
        out.extend(range(lo - 1, hi))
    if not out:
        raise CliError("bad_args", f"页范围为空: {text}")
    seen, result = set(), []
    for p in out:
        if p not in seen:
            seen.add(p)
            result.append(p)
    return result


def page_size_mm(page) -> dict:
    """页尺寸:pt 与 mm(1pt = 25.4/72mm)。"""
    w, h = page.mediabox.width, page.mediabox.height
    return {"width_pt": round(float(w), 2), "height_pt": round(float(h), 2),
            "width_mm": round(float(w) * 25.4 / 72, 1),
            "height_mm": round(float(h) * 25.4 / 72, 1)}

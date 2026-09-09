# -*- coding: utf-8 -*-
"""原子写公共工具: os.replace 带退避重试。

Windows 上新建文件后立即 os.replace 覆盖旧文件,偶发 WinError 5 拒绝访问
(杀软/索引器对新文件的瞬时锁,通常 <1s)。office 各组(rag/prep、docx2md、
xlutil、mdutil 等)统一走 replace_with_retry,避免用户侧偶发写盘失败。
"""

import os
import time


def replace_with_retry(tmp: str, path: str, tries: int = 8) -> None:
    """临时文件 → 目标路径,失败退避重试(指数 0.2s 起)。"""
    last: OSError | None = None
    for i in range(tries):
        try:
            os.replace(tmp, path)
            return
        except OSError as e:  # noqa: BLE001 — 文件锁/目录句柄都算瞬时故障
            last = e
            if i == tries - 1:
                break
            time.sleep(0.2 * (i + 1))
    raise last  # type: ignore[misc]

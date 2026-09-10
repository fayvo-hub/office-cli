"""构建无 Python 环境可用的单文件 office.exe(PyInstaller)。

用法: python scripts/build_exe.py
产物: dist/office.exe(约 100MB;每次代码更新后重新构建)
覆盖: Excel/CSV/JSON/Word/PDF/md→docx|html 等纯 Python 能力
      + 老格式升级与 Office→PDF(WPS COM 或跨平台 LibreOffice, 经 office.engine 门面)
      + xlsx 公式重算(WPS/LibreOffice 无头引擎)
不覆盖: md→pdf 需要 playwright 库(单文件版不打包,运行时提示装完整版)
"""
from __future__ import annotations

import glob
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _module_names() -> list[str]:
    """收集全部命令模块(cli.py 用 importlib 动态导入,PyInstaller 需显式列出)。"""
    # 函数内延迟导入的内部模块 + 动态注册的命令模块
    mods = ["office.wps_worker", "office.convert", "office.info",
            "office.mdutil", "office.docx2md", "office.pdfutil",
            "office.xlutil", "office.ioplan", "office.engine", "office.lo", "office.printer"]
    for f in glob.glob(os.path.join(ROOT, "office", "cmds_*", "*.py")):
        if os.path.basename(f).startswith("_"):
            continue
        rel = os.path.relpath(f, ROOT).replace(os.sep, "/")[:-3]
        mods.append(rel.replace("/", "."))
    return sorted(set(mods))


def main() -> None:
    entry = os.path.join(ROOT, "scripts", "exe_launcher.py")
    if not os.path.exists(entry):
        sys.exit(f"找不到入口: {entry}")
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean",
        "--onefile", "--console",
        "--name", "office",
        "--paths", ROOT,
        "--exclude-module", "playwright",
        "--hidden-import", "pythoncom",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pdf2docx",
        "--collect-submodules", "win32com",
    ]
    # 本机可能装有 ML 大库(torch/scipy...),与 office-cli 无关,一律排除
    # (cv2/opencv 不能排除:pdf2docx 运行时依赖)
    for mod in ("torch", "llvmlite", "pyarrow", "av", "scipy",
                "transformers", "numba", "matplotlib", "pandas",
                "tensorflow", "keras", "sklearn", "nltk", "jieba"):
        cmd += ["--exclude-module", mod]
    for mod in _module_names():
        cmd += ["--hidden-import", mod]
    cmd.append(entry)
    subprocess.check_call(cmd, cwd=ROOT)
    exe = os.path.join(ROOT, "dist", "office.exe")
    print(f"构建完成: {exe}({os.path.getsize(exe) / 1024 / 1024:.1f} MB)")


if __name__ == "__main__":
    main()

"""PyInstaller 单文件入口(等价 python -m office)。"""
from __future__ import annotations

import sys

from office.cli import main

if __name__ == "__main__":
    sys.exit(main())

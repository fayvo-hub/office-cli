# -*- coding: utf-8 -*-
"""可选第三方公式引擎适配(PyPI `formulas` 包)。

安装: `pip install "office-cli[formula]"`(会带入 numpy/scipy 等, 约 70MB)。
启用: `office rag prep -f 表.xlsx --engine formulas`

为什么需要: 内置求值器不实现数组表达式(如 `=SUMPRODUCT((C5:C9>3)*B5:B9)`),
`formulas` 包能算(实测同批 23 条公式 23/23)。代价是重依赖与每文件 1~10s 加载。

接口: `load(path) -> {(sheet_lower, row, col): value}`,供 `_SheetEnv` 在公式格优先取值;
错误值(#DIV/0! 等)不返回 —— 交由缓存兜底/记 unresolved,与内置求值器语义一致。
"""

from __future__ import annotations

import datetime as _dt
import importlib.util
import re

from ..errors import CliError

# formulas 的解 key 形如: '[(文件名)]Sheet 名'!F5(不含 $ 也可出现)
_KEY_RE = re.compile(r"^'\[(?P<file>[^\]]*)\](?P<sheet>[^']*)'!"
                     r"(?P<col>\$?[A-Za-z]{1,3})(?P<row>\$?\d+)$")
_ERR_RE = re.compile(r"^#(?:NULL|DIV/0|VALUE|REF|NAME|NUM|N/A|GETTING_DATA)[!?]?$", re.I)


def available() -> bool:
    """当前环境是否装了 formulas 引擎(仅探测,不导入)。"""
    try:
        return importlib.util.find_spec("formulas") is not None
    except (ImportError, ValueError):
        return False


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters.lstrip("$").upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _plain(value):
    """numpy/日期包装 → 纯 Python 值;错误值/复杂结构返回 None。"""
    if value is None:
        return None
    try:                                  # numpy 标量/数组(不强制依赖 numpy)
        item = getattr(value, "item", None)
        if item is not None and getattr(value, "shape", None) == ():
            value = item()
        elif getattr(value, "size", None) == 1 and hasattr(value, "reshape"):
            value = value.reshape(-1)[0]
            if hasattr(value, "item"):
                value = value.item()
    except Exception:  # noqa: BLE001
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return None if _ERR_RE.match(value.strip()) else value
    if isinstance(value, (int, float)):
        return value
    if isinstance(value, _dt.datetime):
        if (value.hour, value.minute, value.second, value.microsecond) == (0, 0, 0, 0):
            return value.date()           # 与内置求值器一致: 零点日期按 date 输出
        return value
    if isinstance(value, _dt.date):
        return value
    return None


def load(path: str) -> dict:
    """用 formulas 包计算整个工作簿, 返回 {(sheet_lower, row, col): value}。

    未安装依赖时抛 CliError(missing_dependency);文件无法解析时抛 CliError(cannot_open)。
    """
    if not available():
        raise CliError("missing_dependency",
                       "未安装 formulas 公式引擎,请先安装: "
                       "pip install \"office-cli[formula]\"(或改用 --engine wps)")
    import formulas                       # 延迟导入: 默认路径零成本

    try:
        model = formulas.ExcelModel().loads(path).finish()
        solution = model.calculate()
    except Exception as exc:  # noqa: BLE001 第三方引擎异常统一转业务错误
        raise CliError("cannot_open", f"formulas 引擎无法解析该工作簿: {exc}") from exc

    out: dict = {}
    for key, cell in solution.items():
        m = _KEY_RE.match(key)
        if m is None:
            continue
        val = _plain(getattr(cell, "value", None))
        if val is None:
            continue
        out[(m.group("sheet").lower(), int(m.group("row").lstrip("$")),
             _col_num(m.group("col")))] = val
    return out

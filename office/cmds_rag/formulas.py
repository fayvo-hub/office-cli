# -*- coding: utf-8 -*-
"""rag prep 内置轻量公式求值器(v2)。

为什么存在: xlsx 公式无缓存值时,v1 保留公式原文并记入 formula_unresolved。
v2 对常用公式直接求值 —— 求值器覆盖真实台账高频场景:

- 算术: + - * / ^ %、一元负、括号、文本连接 &
- 比较: = <> < > <= >=(结果供 IF/条件用)
- 引用: A1 / $A$1 / A1:B3 / A:A 整列 / Sheet1!A1 / '工作 表'!A1:B2(跨 sheet)
- 函数: 聚合 SUM AVERAGE COUNT COUNTA COUNTBLANK MAX MIN PRODUCT SUMPRODUCT
        MEDIAN LARGE SMALL RANK STDEV(.S/.P) VAR(.S/.P) PERCENTILE
        条件 SUMIF(S) COUNTIF(S) AVERAGEIF(S) MAXIFS MINIFS
        数学 ROUND ROUNDUP ROUNDDOWN INT ABS MOD SQRT EXP LN LOG LOG10 POWER
        SIGN TRUNC CEILING FLOOR MROUND FACT EVEN ODD SUMSQ
        财务 PMT FV PV NPV IRR
        逻辑 IF IFS IFERROR IFNA AND OR NOT CHOOSE SWITCH
        查找 VLOOKUP HLOOKUP XLOOKUP LOOKUP INDEX MATCH
        引用 ROW COLUMN ROWS COLUMNS OFFSET
        文本 LEFT RIGHT MID LEN TRIM UPPER LOWER SUBSTITUTE CONCATENATE CONCAT
        TEXTJOIN TEXT FIND SEARCH REPLACE REPT EXACT VALUE PROPER CLEAN CHAR CODE
        日期 TODAY NOW DATE DATEVALUE YEAR MONTH DAY HOUR MINUTE SECOND WEEKDAY
        DAYS EOMONTH EDATE DATEDIF TIME NETWORKDAYS WORKDAY
        信息 ISNUMBER ISBLANK ISTEXT ISNONTEXT ISLOGICAL ISERROR ISERR ISNA NA

原则:
- 能算的算出真值;算不了(未知函数/循环引用/外部链接/数组公式)由调用方保留原文并记录原因
- 已知缺口与兜底: 数组表达式(如 `=SUMPRODUCT((C5:C9>3)*B5:B9)`)、数组常量 `{…}`、
  RAND/INDIRECT 类易变函数不实现 —— 需真值时 `office rag prep --recalc` 用本机 WPS/Excel
  引擎重算后读缓存(实测同批 23 条公式: WPS 23/23、本求值器 21/23(差数组表达式与
  EOMONTH 日期表示)、PyPI formulas 包 23/23 但需 numpy+scipy 且单文件加载 1~10s)
- IF/IFS/IFERROR/CHOOSE/SWITCH 惰性求值: 只算被选中的分支(与 Excel 一致, 避免 false 分支抛错)
- 不实现易变函数(RAND/RANDBETWEEN/OFFSET 的易变部分除外);数组常量 {…} 不支持
- 日期按 Excel 1900 序列号参与运算,输出还原为 ISO 日期
- 不引入任何第三方依赖,纯标准库

调用方需提供 FormulaEnv:
    cell(sheet, row, col)      -> 该坐标的值:number|str|bool|date|None(空)或抛 FormulaError
    sheet_max_row(sheet)       -> 整列区域上限(取实际最大行)
    sheet_exists(sheet)        -> 工作表是否存在
未捕获 FormulaError 即"算不出",由调用方决定兜底。
"""

from __future__ import annotations

import calendar
import datetime as _dt
import fnmatch
import math
import re

# 值域统一为: number(float) / text(str) / boolean(bool) / date(datetime.date)
_EMPTY = object()          # 空单元格
_DATE_EPOCH = _dt.date(1899, 12, 30)   # Excel 1900 序列号基准


class FormulaError(Exception):
    """求值失败(引用错误/除零/不支持函数/循环等)。message 面向调用方展示。

    code 为可选的 Excel 错误类别('na'/'value'/'ref'/'div0'/'num'/'name'),
    供 IFNA/ISNA/ISERR 等区分; 缺省 None 表示普通算不出。
    """

    def __init__(self, message: str, code: str | None = None):
        super().__init__(message)
        self.code = code


def _to_number(v) -> float:
    if v is None:                # 空白单元格引用 → 0(Excel 语义)
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, _dt.datetime):
        return float(_serial_of(v))
    if isinstance(v, _dt.date):
        return float(_serial_of(v))
    if isinstance(v, str):
        s = v.strip()
        if not s:                # 空文本按 0(宽容,避免空白格拖垮整列)
            return 0.0
        if s == "TRUE":
            return 1.0
        if s == "FALSE":
            return 0.0
        try:
            return float(s.replace("%", "") or 0) / (100 if s.endswith("%") else 1)
        except ValueError:
            raise FormulaError(f"文本不能参与数值运算: '{v}'") from None
    raise FormulaError(f"无法数值化: {type(v).__name__}")


def _serial_of(d) -> float:
    """date/datetime → Excel 序列号(1900 系统)。"""
    if isinstance(d, _dt.datetime):
        frac = (d.hour * 3600 + d.minute * 60 + d.second + d.microsecond / 1e6) / 86400
        return float((d.date() - _DATE_EPOCH).days) + frac
    return float((d - _DATE_EPOCH).days)


def _from_serial(n: float):
    """Excel 序列号 → date;带小数部分转 datetime。"""
    if not math.isfinite(n):
        raise FormulaError(f"非法日期序列号: {n}")
    if abs(n - round(n)) < 1e-9:
        return _DATE_EPOCH + _dt.timedelta(days=int(round(n)))
    return _DATE_EPOCH + _dt.timedelta(days=n)


def _num_str(v) -> str:
    """数值 → 干净文本(修浮点噪声,去尾零)。"""
    x = round(float(v), 12)
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.10f}".rstrip("0").rstrip(".")
    return s if s not in ("-0", "") else "0"


# ---------------------------------------------------------------------------
# 词法
# ---------------------------------------------------------------------------

_TOKEN_RE = re.compile(r"""
    (?P<ws>\s+)
  | (?P<str>"[^"]*")
  | (?P<sq>'[^']*')
  | (?P<num>\d+(?:\.\d+)?(?:[eE][+-]?\d+)?)
  | (?P<ident>[A-Za-z_\u4e00-\u9fff][A-Za-z_\u4e00-\u9fff]*)
  | (?P<op><>|<=|>=|=|<|>|&|\+|-|\*|/|\^|%|\(|\)|,|:|!|\$|\.)
""", re.X)

# 引用语法里会出现: $ A1 : ! '名称' 等,由 parser 在 token 流上直接拼
_NUM_FUNCS = {"ABS", "INT", "MOD", "ROUND", "ROUNDUP", "ROUNDDOWN"}
_TEXT_FUNCS = {"LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER",
               "SUBSTITUTE", "CONCATENATE"}
_LOG_FUNCS = {"AND", "OR", "NOT"}
_AGG_FUNCS = {"SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MAX", "MIN",
              "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS",
              "AVERAGEIF", "AVERAGEIFS", "MAXIFS", "MINIFS"}
_DATE_FUNCS = {"TODAY", "NOW", "DATE"}
_INFO_FUNCS = {"ISNUMBER", "ISBLANK", "ISTEXT", "ISNONTEXT", "ISLOGICAL"}
_EXTRA_INFO_FUNCS = {"ISERROR", "ISERR", "ISNA"}
# 惰性参数族(只求值用到的分支): 条件/错误捕获/多路选择
_LAZY_FUNCS = {"IF", "IFS", "IFERROR", "IFNA", "CHOOSE", "SWITCH",
               "ISERROR", "ISERR", "ISNA"}
# 统计族(区域展开, 引用中的文本/布尔忽略)
_STAT_FUNCS = {"MEDIAN", "LARGE", "SMALL", "RANK", "RANK.EQ", "STDEV",
               "STDEV.S", "STDEV.P", "VAR", "VAR.S", "VAR.P", "PRODUCT",
               "PERCENTILE", "SUMPRODUCT"}
# 数学/财务族
_MATH_FUNCS = {"SQRT", "EXP", "LN", "LOG", "LOG10", "SIGN", "TRUNC",
               "POWER", "CEILING", "FLOOR", "MROUND", "FACT", "EVEN",
               "ODD", "SUMSQ", "PMT", "FV", "PV", "NPV", "IRR"}
# 查找族
_LOOKUP_FUNCS = {"VLOOKUP", "HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "LOOKUP"}
# 引用族
_REF_FUNCS = {"ROW", "COLUMN", "ROWS", "COLUMNS", "OFFSET"}
# 日期族扩展
_EXTRA_DATE_FUNCS = {"YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND",
                     "WEEKDAY", "DAYS", "EOMONTH", "EDATE", "DATEDIF",
                     "DATEVALUE", "TIME", "NETWORKDAYS", "WORKDAY"}
# 文本族扩展
_EXTRA_TEXT_FUNCS = {"FIND", "SEARCH", "REPLACE", "REPT", "EXACT", "VALUE",
                     "PROPER", "CLEAN", "CHAR", "CODE", "TEXTJOIN",
                     "CONCAT", "TEXT"}
_NA_FUNCS = {"NA"}


def _is_ref(a):
    """args() 产出的元素可能是 ('cell'|'region', payload) 引用或直接标量值。"""
    return isinstance(a, tuple) and len(a) == 2 and a[0] in ("cell", "region")


class _Tok:
    __slots__ = ("kind", "text")

    def __init__(self, kind: str, text: str):
        self.kind = kind
        self.text = text

    def __repr__(self):  # pragma: no cover
        return f"<{self.kind}:{self.text}>"


def _tokenize(s: str) -> list[_Tok]:
    toks: list[_Tok] = []
    pos = 0
    while pos < len(s):
        m = _TOKEN_RE.match(s, pos)
        if not m:
            raise FormulaError(f"无法解析公式片段: '…{s[pos:pos+12]}…'")
        pos = m.end()
        kind = m.lastgroup
        if kind == "ws":
            continue
        if kind == "num":
            toks.append(_Tok("num", m.group()))
        elif kind == "str":
            toks.append(_Tok("str", m.group()[1:-1]))
        elif kind == "sq":
            toks.append(_Tok("sq", m.group()[1:-1]))
        elif kind == "ident":
            toks.append(_Tok("ident", m.group()))
        else:
            toks.append(_Tok(m.group(), m.group()))
    return toks


# ---------------------------------------------------------------------------
# 文本/日期格式化辅助(模块级, 供 TEXT 与日期族复用)
# ---------------------------------------------------------------------------

_DATE_TEXT_RES = (
    re.compile(r"^(\d{4})[-/年.](\d{1,2})[-/月.](\d{1,2})日?$"),
    re.compile(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$"),
)


def _parse_date_text(s: str):
    """常见日期文本 → date; 失败 None(2024-01-12 / 2024年1月12日 / 1/12/2024 / 20240112)。"""
    t = s.strip()
    if not t:
        return None
    for i, rx in enumerate(_DATE_TEXT_RES):
        m = rx.match(t)
        if m:
            g = [int(x) for x in m.groups()]
            y, mo, d = (g[2], g[0], g[1]) if i == 1 else (g[0], g[1], g[2])
            try:
                return _dt.date(y, mo, d)
            except ValueError:
                return None
    if re.fullmatch(r"\d{8}", t):
        try:
            return _dt.date(int(t[:4]), int(t[4:6]), int(t[6:]))
        except ValueError:
            return None
    return None


def _parse_number_text(s: str):
    """数值文本 → float(去千分位/货币/空格, 支持 % 与括号负数); 失败 None。"""
    t = s.strip()
    if not t:
        return None
    neg = t.startswith("(") and t.endswith(")")
    if neg:
        t = t[1:-1].strip()
    pct = t.endswith("%")
    if pct:
        t = t[:-1].strip()
    for ch in ("\u00a5", "$", "\u20ac", "\u00a3", ",", " ", "\u00a0"):
        t = t.replace(ch, "")
    try:
        n = float(t)
    except ValueError:
        return None
    if pct:
        n /= 100.0
    return -n if neg else n


def _proper(s: str) -> str:
    """每词首字母大写(仅 ASCII 字母, 其余原样)。"""
    out = []
    prev_alpha = False
    for ch in s:
        if ch.isalpha() and ch.isascii():
            out.append(ch.upper() if not prev_alpha else ch.lower())
            prev_alpha = True
        else:
            out.append(ch)
            prev_alpha = False
    return "".join(out)


def _shift_month(d: _dt.date, months: int, keep_day: bool) -> _dt.date:
    """按月偏移; keep_day=True 保持日(不存在则取月末), False 取目标月月末。"""
    y = d.year + (d.month - 1 + months) // 12
    m = (d.month - 1 + months) % 12 + 1
    last = calendar.monthrange(y, m)[1]
    return _dt.date(y, m, min(d.day, last) if keep_day else last)


def _fmt_date_tokens(d, fmt: str) -> str:
    """迷你日期格式: yyyy/yy/mm/dd/hh/ss; mm 在 h 后或 : 前按分钟。"""
    low = fmt.lower()
    out = []
    i = 0
    while i < len(fmt):
        if low.startswith("yyyy", i):
            out.append(f"{d.year:04d}")
            i += 4
        elif low.startswith("yy", i):
            out.append(f"{d.year % 100:02d}")
            i += 2
        elif low.startswith("mm", i):
            before = low[:i]
            after = low[i + 2:]
            minute = ("h" in before) or after.startswith(":") or after.startswith("ss")
            if minute:
                out.append(f"{getattr(d, 'minute', 0):02d}")
            else:
                out.append(f"{d.month:02d}")
            i += 2
        elif low.startswith("dd", i):
            out.append(f"{d.day:02d}")
            i += 2
        elif low.startswith("hh", i):
            out.append(f"{getattr(d, 'hour', 0):02d}")
            i += 2
        elif low.startswith("ss", i):
            out.append(f"{getattr(d, 'second', 0):02d}")
            i += 2
        elif low[i] == "m":
            out.append(str(d.month))
            i += 1
        elif low[i] == "d":
            out.append(str(d.day))
            i += 1
        elif low[i] == "h":
            out.append(str(getattr(d, "hour", 0)))
            i += 1
        elif low[i] == "s":
            out.append(str(getattr(d, "second", 0)))
            i += 1
        else:
            out.append(fmt[i])
            i += 1
    return "".join(out)


def _fmt_num_tokens(v, fmt: str) -> str:
    """迷你数字格式: 支持 % 缩放、小数位数(0.00)、千分位(#,##0)。"""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        n = float(v)
    else:
        n = _parse_number_text(_text_of(v))
    if n is None:
        return _text_of(v)
    pct = "%" in fmt
    if pct:
        n *= 100.0
    m = re.search(r"\.([0#]+)", fmt)
    dec = len(m.group(1)) if m else 0
    thousands = "," in fmt
    s = f"{n:,.{dec}f}" if thousands else f"{n:.{dec}f}"
    return s + "%" if pct else s


# ---------------------------------------------------------------------------
# 解析 + 求值(递归下降,边解析边求值;区域以未求值形式捕获)
# ---------------------------------------------------------------------------

class Evaluator:
    def __init__(self, env, sheet: str, row: int, col: int):
        self.env = env
        self.sheet = sheet
        self.row = row          # 公式所在行(1 基)
        self.col = col          # 公式所在列(1 基)
        self.toks: list[_Tok] = []
        self.i = 0

    # ---- 工具 ----
    def peek(self) -> _Tok | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> _Tok:
        t = self.peek()
        if t is None:
            raise FormulaError("公式不完整")
        self.i += 1
        return t

    def accept(self, text: str) -> _Tok | None:
        """仅当下一 token 是纯符号(op)且文本相等时消费;文本类 token 不参与。"""
        t = self.peek()
        if t is not None and t.text == text and t.kind == text:
            self.i += 1
            return t
        return None

    def expect(self, text: str) -> _Tok:
        t = self.accept(text)
        if t is None:
            raise FormulaError(f"公式语法错误: 期望 '{text}'")
        return t

    # ---- 入口 ----
    def run(self, text: str):
        self.toks = _tokenize(text)
        self.i = 0
        v = self.expr()
        if self.peek() is not None:
            raise FormulaError(f"公式结尾有多余内容: '{self.peek().text}'")
        return v

    # ---- 优先级: 比较 < & < 加减 < 乘除 < ^ < 一元/% < 原子 ----
    def expr(self):
        v = self.concat()
        while self.peek() is not None and self.peek().text in ("=", "<>", "<", ">", "<=", ">="):
            op = self.next().text
            r = self.concat()
            v = self.compare(op, v, r)
        return v

    def compare(self, op: str, l, r) -> bool:
        # 数值比较优先;文本 lex;文本与数值: 文本恒大于数值(Excel 语义)
        def norm(x):
            x = self._deref(x)               # 引用先取值
            if x is None:                    # 空单元格按 0(Excel 语义)
                return 0.0, "num"
            if isinstance(x, bool):
                return (1.0 if x else 0.0), "num"
            if isinstance(x, _dt.date):
                return _serial_of(x), "num"
            if isinstance(x, (int, float)):
                return float(x), "num"
            if isinstance(x, str):
                return x, "txt"
            raise FormulaError(f"无法比较的值: {type(x).__name__}")

        ln, lt = norm(l)
        rn, rt = norm(r)
        if lt == rt == "num":
            return _cmp_num(op, ln, rn)
        if lt == rt:
            return _cmp_str(op, ln, rn)
        # 数值 vs 文本: 数值更小
        order = -1 if lt == "num" else 1
        if op == "=":
            return False
        if op == "<>":
            return True
        if op == ">":
            return order == 1
        if op == "<":
            return order == -1
        if op == ">=":
            return order >= 0
        return order <= 0

    def concat(self):
        v = self.additive()
        while self.accept("&"):
            r = self.additive()
            v = _text_of(self._deref(v)) + _text_of(self._deref(r))
        return v

    def additive(self):
        v = self.multiplicative()
        while True:
            if self.accept("+"):
                v = _to_number(self._deref(v)) + _to_number(self._deref(self.multiplicative()))
            elif self.accept("-"):
                v = _to_number(self._deref(v)) - _to_number(self._deref(self.multiplicative()))
            else:
                return v

    def multiplicative(self):
        v = self.power()
        while True:
            if self.accept("*"):
                v = _to_number(self._deref(v)) * _to_number(self._deref(self.power()))
            elif self.accept("/"):
                d = _to_number(self._deref(self.power()))
                if d == 0:
                    raise FormulaError("除数为零")
                v = _to_number(self._deref(v)) / d
            else:
                return v

    def power(self):
        v = self.unary()
        while self.accept("^"):
            v = _to_number(self._deref(v)) ** _to_number(self._deref(self.unary()))
        return v

    def unary(self):
        if self.accept("-"):
            return -_to_number(self._deref(self.unary()))
        if self.accept("+"):
            return _to_number(self._deref(self.unary()))
        v = self.postfix()
        while self.accept("%"):
            v = _to_number(self._deref(v)) / 100.0
        return v

    def postfix(self):
        return self.primary()

    # ---- 引用 ----
    def _try_reference(self):
        """尝试消费一个引用;失败返回 None 且不前进。成功返回 ('cell'|'region', payload)。

        payload: cell -> (sheet, row, col);region -> (sheet, r1, c1, r2, c2)
        """
        save = self.i
        sheet = None
        t0 = self.peek()
        if t0 is not None and t0.kind == "sq":
            # 引号包裹的表名(tokenizer 已把 '…' 合成 sq token): '表 名'!
            nxt = self.toks[self.i + 1] if self.i + 1 < len(self.toks) else None
            if nxt is not None and nxt.text == "!":
                sheet = t0.text
                self.i += 2
            else:
                self.i = save
                return None
        else:
            # 裸表名(如 汇总!、Sheet1!): 试探读 ident/num 段 + '!'
            seg: list[str] = []
            while True:
                t = self.peek()
                if t is not None and t.kind in ("ident", "num"):
                    seg.append(t.text)
                    self.i += 1
                else:
                    break
            if seg and self.accept("!"):
                cand = "".join(seg)
                if self.env.sheet_exists(cand):
                    sheet = cand
                else:
                    self.i = save
                    return None
            elif seg:
                self.i = save          # 不是表名引用,回退单元格解析

        def col_of(tok) -> int:
            n = 0
            for ch in tok.text.lower():
                n = n * 26 + (ord(ch) - 96)
            if n > 16384:
                raise FormulaError(f"列超出范围 '{tok.text}'")
            return n

        def read_cell():
            """读 [$]列[$]行;列必选,行可缺省(整列引用)。返回 (col, row|None)。"""
            self.accept("$")
            c_tok = self.peek()
            if c_tok is None or c_tok.kind != "ident":
                raise FormulaError("引用缺少列字母")
            col = col_of(self.next())
            self.accept("$")
            r_tok = self.peek()
            row = None
            if r_tok is not None and r_tok.kind == "num":
                try:
                    row = int(float(self.next().text))
                except ValueError:
                    raise FormulaError(f"非法行号 '{r_tok.text}'") from None
                if row < 1 or row > 1048576:
                    raise FormulaError(f"行号超出范围 {row}")
            return col, row

        try:
            c1, r1 = read_cell()
        except FormulaError:
            self.i = save
            return None
        if r1 is None:                      # 整列: A 或 A:B
            if self.accept(":"):
                c2, _r2 = read_cell()
                c2 = c2 if c2 is not None else c1
                return ("region", (sheet or self.sheet, 1, c1,
                                   self.env.sheet_max_row(sheet or self.sheet), c2))
            return ("region", (sheet or self.sheet, 1, c1,
                               self.env.sheet_max_row(sheet or self.sheet), c1))
        if self.accept(":"):
            c2, r2 = read_cell()
            if r2 is None:                  # A1:B(到该列底)
                c2 = c2 if c2 is not None else c1
                return ("region", (sheet or self.sheet, r1, c1,
                                   self.env.sheet_max_row(sheet or self.sheet), c2))
            return ("region", (sheet or self.sheet, r1, c1, r2, c2))
        return ("cell", (sheet or self.sheet, r1, c1))

    # ---- 原子 ----
    def primary(self):
        t = self.peek()
        if t is None:
            raise FormulaError("公式不完整")
        if t.text == "(":
            self.next()
            v = self.expr()
            self.expect(")")
            return v
        if t.kind == "num":
            self.next()
            return float(t.text)
        if t.kind == "str":
            self.next()
            return t.text
        if t.kind == "sq":
            # 引号表名 '表 名'!A1 由 _try_reference 消费;否则按文本字面量
            ref = self._try_reference()
            if ref is not None:
                return ref
            self.next()
            return t.text
        if t.kind == "ident":
            name = t.text.upper()
            # 函数名可带数字/点(LOG10 / VAR.P / STDEV.S / RANK.EQ); 仅当后面紧跟 '(' 才合并
            j = self.i + 1
            while j < len(self.toks):
                nx = self.toks[j]
                if nx.kind == "num" and re.fullmatch(r"\d+", nx.text):
                    name += nx.text
                    j += 1
                elif (nx.text == "." and j + 1 < len(self.toks)
                      and self.toks[j + 1].kind == "ident"):
                    name += "." + self.toks[j + 1].text.upper()
                    j += 2
                else:
                    break
            nxt = self.toks[j] if j < len(self.toks) else None
            if nxt is not None and nxt.text == "(":
                self.i = j + 1                  # 消费 函数名 与 '('
                if name in _LAZY_FUNCS:         # 惰性: 只扫描参数区间, 被选中的分支才求值
                    spans = self.args_lazy()
                    self.expect(")")
                    return self.call_lazy(name, spans)
                args = self.args()
                self.expect(")")
                return self.call(name, args)
            # 引用(列/单元格/区域/跨表)
            ref = self._try_reference()
            if ref is not None:
                return ref
            if name in ("TRUE", "FALSE"):      # 裸 TRUE/FALSE 常量
                self.next()
                return name == "TRUE"
            raise FormulaError(f"未知名称 '{t.text}'")
        ref = self._try_reference()          # $A$1 等以 $ 开头的引用
        if ref is not None:
            return ref
        raise FormulaError(f"意外的符号 '{t.text}'")

    def args(self):
        """参数列表;每个参数走完整表达式(primary 自会解析 cell/region/嵌套函数)。"""
        if self.peek() is not None and self.peek().text == ")":
            return []
        out = []
        while True:
            out.append(self.expr())
            if not self.accept(","):
                return out

    def args_lazy(self):
        """惰性参数: 仅按括号/逗号配平扫描出每个参数的 token 区间(不求值)。"""
        spans = []
        if self.peek() is not None and self.peek().text == ")":
            return spans
        while True:
            start = self.i
            depth = 0
            while True:
                t = self.peek()
                if t is None:
                    raise FormulaError("公式不完整")
                if t.text == "(":
                    depth += 1
                elif t.text == ")":
                    if depth == 0:
                        break
                    depth -= 1
                elif t.text == "," and depth == 0:
                    break
                self.i += 1
            spans.append((start, self.i))
            if self.accept(","):
                continue
            return spans

    def eval_span(self, span):
        """按 token 区间求值子表达式(共享 env 与当前坐标, 与主解析同语义)。"""
        start, end = span
        if start >= end:
            return None              # 空参数(=IF(A1,,1))按空白处理
        sub = Evaluator(self.env, self.sheet, self.row, self.col)
        sub.toks = self.toks[start:end]
        sub.i = 0
        v = sub.expr()
        if sub.peek() is not None:
            raise FormulaError(f"参数结尾有多余内容: '{sub.peek().text}'")
        return v

    # ---- 值化 ----
    def value_of(self, arg):
        """参数 → 普通值;区域参数仅当是单格时值化。"""
        if not _is_ref(arg):
            return arg
        if arg[0] in ("cell",):
            return self.cell_value(*arg[1])
        if arg[0] == "region":
            _, (sh, r1, c1, r2, c2) = arg
            if r1 == r2 and c1 == c2:
                return self.cell_value(sh, r1, c1)
            raise FormulaError("函数参数不接受多格区域")
        return arg

    def _deref(self, x):
        """运算前对引用取值;多格区域禁止参与运算。"""
        if not _is_ref(x):
            return x
        if x[0] == "cell":
            return self.cell_value(*x[1])
        _s, r1, c1, r2, c2 = x[1]
        if r1 == r2 and c1 == c2:
            return self.cell_value(_s, r1, c1)
        raise FormulaError("区域不能参与算术运算(仅能作函数参数)")

    def cell_value(self, sheet: str, row: int, col: int):
        v = self.env.cell(sheet, row, col)      # 可能触发嵌套公式求值
        return v

    # ---- 函数分发 ----
    def call(self, name: str, args):
        if name in ("TRUE", "FALSE"):
            if args:
                raise FormulaError(f"{name}() 不接受参数")
            return name == "TRUE"
        if name in _NA_FUNCS:
            raise FormulaError("NA()", "na")
        if name in _LOOKUP_FUNCS:
            return self._lookup_fn(name, args)
        if name in _STAT_FUNCS:
            return self._stat_fn(name, args)
        if name in _MATH_FUNCS:
            return self._math_fn(name, args)
        if name in _REF_FUNCS:
            return self._ref_fn(name, args)
        if name in _EXTRA_DATE_FUNCS:
            return self._date_fn2(name, args)
        if name in _EXTRA_TEXT_FUNCS:
            return self._text_fn2(name, args)
        if name in _AGG_FUNCS:
            return self._agg(name, args)
        if name in _NUM_FUNCS:
            return self._num_fn(name, args)
        if name in _TEXT_FUNCS:
            return self._text_fn(name, args)
        if name in _LOG_FUNCS:
            return self._log_fn(name, args)
        if name in _INFO_FUNCS:
            return self._info_fn(name, args)
        if name in _DATE_FUNCS:
            if name == "DATE":
                if len(args) != 3:
                    raise FormulaError("DATE() 需要 3 个参数")
                nums = [self._num_loose(self.value_of(a)) for a in args]
                if any(n is None for n in nums):
                    raise FormulaError("DATE() 参数必须能转为数值")
                y, m, d = (int(n) for n in nums)
                try:
                    return _dt.date(y, m, d)
                except ValueError:
                    raise FormulaError(f"非法日期 {y}-{m}-{d}") from None
            return _dt.date.today() if name == "TODAY" else _dt.datetime.now()
        raise FormulaError(f"不支持的函数 {name}()")

    # ---- 聚合族 ----
    def _iter_region(self, arg):
        """区域展开: 每格值;区域越界自动收敛(引用到表外行/列取不到为空)。"""
        if not _is_ref(arg):
            yield arg
            return
        if arg[0] == "cell":
            yield self.env.cell(*arg[1])
            return
        sh, r1, c1, r2, c2 = arg[1]
        if not self.env.sheet_exists(sh):
            raise FormulaError(f"引用的工作表不存在: {sh}")
        r2 = min(r2, self.env.sheet_max_row(sh))
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                yield self.env.cell(sh, r, c)

    def _scalar_args(self, args):
        """参数 → (ref_vals, direct_vals): 引用/区域来的格值与直接表达式值分开。

        Excel 语义: 引用中的文本/布尔被忽略(仅数值/日期参与统计);
        直接参数(如 SUM("5"))才做宽松转换。
        """
        refs: list = []
        dirs: list = []
        for a in args:
            if not _is_ref(a):
                dirs.append(a)
            elif a[0] == "cell":
                refs.append(self.env.cell(*a[1]))
            elif a[0] == "region":
                refs.extend(self._iter_region(a))
            else:
                dirs.append(a)
        return refs, dirs

    @staticmethod
    def _num_strict(v):
        """引用格值 → 数值;文本/布尔/None 一律忽略(Excel 引用语义)。"""
        if v is None or isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, _dt.date):
            return _serial_of(v)
        return None

    @staticmethod
    def _num_loose(v):
        """直接参数 → 数值;纯数字文本可转(如 SUM("5"))。"""
        if isinstance(v, bool):
            return 1.0 if v else 0.0
        if isinstance(v, (int, float)):
            return float(v)
        if isinstance(v, _dt.date):
            return _serial_of(v)
        if isinstance(v, str):
            s = v.strip()
            if not s:
                return None
            try:
                return float(s)
            except ValueError:
                return None
        return None

    def _agg(self, name, args):
        if name in ("SUM", "AVERAGE", "MAX", "MIN", "COUNT", "COUNTA", "COUNTBLANK"):
            refs, dirs = self._scalar_args(args)
            nums = [x for x in (self._num_strict(v) for v in refs) if x is not None]
            for v in dirs:
                n = self._num_loose(v)
                if n is not None:
                    nums.append(n)
            if name == "SUM":
                return float(sum(nums))
            if name == "AVERAGE":
                if not nums:
                    raise FormulaError("AVERAGE() 参数不含数值")
                return float(sum(nums) / len(nums))
            if name == "MAX":
                return float(max(nums)) if nums else 0.0
            if name == "MIN":
                return float(min(nums)) if nums else 0.0
            if name == "COUNT":
                return float(len(nums))
            if name == "COUNTA":
                def _nonblank(v):
                    return not (v is None or (isinstance(v, str) and v == ""))
                return float(sum(1 for v in refs if _nonblank(v))
                             + sum(1 for v in dirs if _nonblank(v)))
            # COUNTBLANK: 空白 = None 或空串(真实文件常见写入脏空串)
            def _blank(v):
                return v is None or (isinstance(v, str) and v == "")
            return float(sum(1 for v in refs if _blank(v)))
        # SUMIF/SUMIFS/COUNTIF/COUNTIFS
        return self._cond_agg(name, args)

    def _cond_agg(self, name, args):
        def need(n):
            if len(args) < n:
                raise FormulaError(f"{name}() 参数不足")

        def geometry(arg):
            """区域几何 (sh, r1, c1, h, w);单格视为 1x1;整列按实际行高。"""
            if not _is_ref(arg):
                raise FormulaError("条件/求和区域必须是单元格引用")
            if arg[0] == "cell":
                sh, r, c = arg[1]
                return (sh, r, c, 1, 1)
            sh, r1, c1, r2, c2 = arg[1]
            r2 = min(r2, self.env.sheet_max_row(sh))
            return (sh, r1, c1, r2 - r1 + 1, c2 - c1 + 1)

        def at(arg, g, idx):
            """按主区域几何取第 idx 格;若次区域只是单格,按 Excel 语义从该格扩展同尺寸。"""
            if not _is_ref(arg):
                raise FormulaError("条件/求和区域必须是单元格引用")
            if arg[0] == "cell":
                sh, r, c = arg[1]
                r0, c0 = g[1] + idx // g[4], g[2] + idx % g[4]
                return self.env.cell(sh, r + (r0 - g[1]), c + (c0 - g[2]))
            sh, r1, c1, _r2, _c2 = arg[1]
            return self.env.cell(sh, r1 + idx // g[4], c1 + idx % g[4])

        def iterate_geom(arg):
            """(geometry, iter_values)"""
            if not _is_ref(arg):
                raise FormulaError("条件/求和区域必须是单元格引用")
            if arg[0] == "cell":
                v = self.env.cell(*arg[1])
                return (geometry(arg), iter([v]))
            g = geometry(arg)
            sh, r1, c1, h, w = g
            vals = (self.env.cell(sh, r, c)
                    for r in range(r1, r1 + h) for c in range(c1, c1 + w))
            return (g, vals)

        if name == "SUMIF":
            need(2)
            rng = args[0]
            crit = _Criteria.parse(self.value_of(args[1]))
            sum_rng = args[2] if len(args) > 2 else rng
            g, it = iterate_geom(rng)
            total = 0.0
            i = 0
            for v in it:
                if crit.match(v):
                    n = self._num_strict(at(sum_rng, g, i))
                    if n is not None:
                        total += n
                i += 1
            return total
        if name == "COUNTIF":
            need(2)
            crit = _Criteria.parse(self.value_of(args[1]))
            return float(sum(1 for v in self._iter_region(args[0]) if crit.match(v)))
        if name == "SUMIFS":     # SUMIFS(sum_range, crit_range1, crit1, …)
            need(3)
            g, it = iterate_geom(args[0])
            pairs = [(args[i], args[i + 1]) for i in range(1, len(args), 2)]
            total = 0.0
            i = 0
            for v in it:
                ok = all(_Criteria.parse(self.value_of(cc)).match(at(cr, g, i))
                         for cr, cc in pairs)
                if ok:
                    n = self._num_strict(v)
                    if n is not None:
                        total += n
                i += 1
            return total
        if name == "COUNTIFS":
            need(2)
            pairs = [(args[i], args[i + 1]) for i in range(0, len(args), 2)]
            g, it = iterate_geom(pairs[0][0])
            n = 0
            i = 0
            for v in it:
                if all(_Criteria.parse(self.value_of(cc)).match(at(cr, g, i))
                       for cr, cc in pairs):
                    n += 1
                i += 1
            return float(n)
        if name == "AVERAGEIF":
            need(2)
            rng = args[0]
            crit = _Criteria.parse(self.value_of(args[1]))
            avg_rng = args[2] if len(args) > 2 else rng
            g, it = iterate_geom(rng)
            vals = []
            i = 0
            for v in it:
                if crit.match(v):
                    n = self._num_strict(at(avg_rng, g, i))
                    if n is not None:
                        vals.append(n)
                i += 1
            if not vals:
                raise FormulaError("AVERAGEIF() 没有匹配的数值", "div0")
            return sum(vals) / len(vals)
        if name in ("AVERAGEIFS", "MAXIFS", "MINIFS"):
            need(3)
            tgt = args[0]
            pairs = [(args[i], args[i + 1]) for i in range(1, len(args), 2)]
            g, it = iterate_geom(pairs[0][0])
            vals = []
            i = 0
            for v in it:
                if all(_Criteria.parse(self.value_of(cc)).match(at(cr, g, i))
                       for cr, cc in pairs):
                    n = self._num_strict(at(tgt, g, i))
                    if n is not None:
                        vals.append(n)
                i += 1
            if not vals:
                if name == "AVERAGEIFS":
                    raise FormulaError("AVERAGEIFS() 没有匹配的数值", "div0")
                return 0.0
            if name == "AVERAGEIFS":
                return sum(vals) / len(vals)
            return max(vals) if name == "MAXIFS" else min(vals)
        raise FormulaError(f"不支持的聚合 {name}()")

    def _cell_at(self, arg, idx: int):
        """区域按游标取第 idx 格值(与另一区域配对);单格返回自身。"""
        if arg[0] == "cell":
            return self.env.cell(*arg[1])
        sh, r1, c1, r2, c2 = arg[1]
        r2 = min(r2, self.env.sheet_max_row(sh))
        width = c2 - c1 + 1
        r, c = r1 + idx // width, c1 + idx % width
        if r > r2:
            return None
        return self.env.cell(sh, r, c)

    # ---- 区域/取值辅助(查找/统计/引用族共用) ----
    def _region_geo(self, arg):
        """区域几何 (sh, r1, c1, h, w); 非引用返回 None。整列按实际使用范围收敛。"""
        if not _is_ref(arg):
            return None
        if arg[0] == "cell":
            sh, r, c = arg[1]
            return (sh, r, c, 1, 1)
        sh, r1, c1, r2, c2 = arg[1]
        if not self.env.sheet_exists(sh):
            raise FormulaError(f"引用的工作表不存在: {sh}", "ref")
        r2 = min(r2, self.env.sheet_max_row(sh))
        return (sh, r1, c1, max(r2 - r1 + 1, 0), c2 - c1 + 1)

    def _flat_values(self, args):
        """参数展开为平面值列表(区域按行列顺序展开; 单格取单值)。"""
        out = []
        for a in args:
            if not _is_ref(a):
                out.append(a)
            elif a[0] == "cell":
                out.append(self.env.cell(*a[1]))
            else:
                out.extend(self._iter_region(a))
        return out

    def _nums(self, args):
        """聚合数值: 引用/区域中的文本与布尔忽略; 直接参数宽松转换。"""
        refs, dirs = self._scalar_args(args)
        nums = [x for x in (self._num_strict(v) for v in refs) if x is not None]
        for v in dirs:
            n = self._num_loose(v)
            if n is not None:
                nums.append(n)
        return nums

    def _lookup_scan(self, vec, needle, mode):
        """一维值列表定位 needle → 0 基下标(未命中 None)。

        mode 0=精确(文本不区分大小写, 支持 * ? 通配), 1=最后一个 <=needle(升序),
        -1=最后一个 >=needle(降序)。
        """
        if mode == 0:
            if isinstance(needle, str) and ("*" in needle or "?" in needle):
                crit = _Criteria.parse(needle)
                for i, v in enumerate(vec):
                    if crit.match(v):
                        return i
                return None
            for i, v in enumerate(vec):
                if self._eq(v, needle):
                    return i
            return None
        hit = None
        for i, v in enumerate(vec):
            if v is None:
                continue
            if isinstance(v, str):
                if not isinstance(needle, str):
                    continue
                ok = (self.compare("<=", v.casefold(), needle.casefold()) if mode > 0
                      else self.compare(">=", v.casefold(), needle.casefold()))
            else:
                n = self._num_strict(v)
                if n is None:
                    continue
                ok = (self.compare("<=", n, needle) if mode > 0
                      else self.compare(">=", n, needle))
            if ok:
                hit = i
        return hit

    # ---- 查找族 ----
    def _lookup_fn(self, name, args):
        if name in ("VLOOKUP", "HLOOKUP"):
            if len(args) < 3:
                raise FormulaError(f"{name}() 参数不足")
            needle = self.value_of(args[0])
            geo = self._region_geo(args[1])
            if geo is None:
                raise FormulaError(f"{name}() 第二参数必须是区域", "value")
            sh, r1, c1, h, w = geo
            idx = self._num_loose(self.value_of(args[2]))
            if idx is None:
                raise FormulaError(f"{name}() 序号不是数值", "value")
            i = int(idx)
            approx = self._truth(self.value_of(args[3])) if len(args) > 3 else True
            mode = 1 if approx else 0
            if name == "VLOOKUP":
                if i < 1 or i > w:
                    raise FormulaError(f"VLOOKUP 第 {i} 列超出区域宽度 {w}", "ref")
                vec = [self.env.cell(sh, r1 + k, c1) for k in range(h)]
                pos = self._lookup_scan(vec, needle, mode)
                if pos is None:
                    raise FormulaError(f"VLOOKUP 未找到 {_text_of(needle)}", "na")
                v = self.env.cell(sh, r1 + pos, c1 + i - 1)
            else:
                if i < 1 or i > h:
                    raise FormulaError(f"HLOOKUP 第 {i} 行超出区域高度 {h}", "ref")
                vec = [self.env.cell(sh, r1, c1 + k) for k in range(w)]
                pos = self._lookup_scan(vec, needle, mode)
                if pos is None:
                    raise FormulaError(f"HLOOKUP 未找到 {_text_of(needle)}", "na")
                v = self.env.cell(sh, r1 + i - 1, c1 + pos)
            return 0.0 if v is None else v
        if name == "MATCH":
            if len(args) < 2:
                raise FormulaError("MATCH() 参数不足")
            needle = self.value_of(args[0])
            geo = self._region_geo(args[1])
            if geo is None:
                raise FormulaError("MATCH() 第二参数必须是区域", "value")
            sh, r1, c1, h, w = geo
            if h > 1 and w > 1:
                raise FormulaError("MATCH() 只支持单行或单列区域", "value")
            if w == 1:
                vec = [self.env.cell(sh, r1 + k, c1) for k in range(h)]
            else:
                vec = [self.env.cell(sh, r1, c1 + k) for k in range(w)]
            mode = 1
            if len(args) > 2:
                m = self._num_loose(self.value_of(args[2]))
                mode = int(m) if m is not None else 1
            pos = self._lookup_scan(vec, needle, mode)
            if pos is None:
                raise FormulaError(f"MATCH 未找到 {_text_of(needle)}", "na")
            return float(pos + 1)
        if name == "INDEX":
            if len(args) < 2:
                raise FormulaError("INDEX() 参数不足")
            a0 = args[0]
            if not _is_ref(a0):
                raise FormulaError("INDEX() 第一参数必须是区域", "value")
            sh, r1, c1, h, w = self._region_geo(a0)
            rn = self._num_loose(self.value_of(args[1]))
            rn = 0 if rn is None else int(rn)
            cn = None
            if len(args) > 2:
                cnv = self._num_loose(self.value_of(args[2]))
                cn = 0 if cnv is None else int(cnv)
            if len(args) == 2:
                if h == 1:                    # 单行区域: 参数是列序号
                    rn, cn = 1, rn
                else:
                    cn = 1
            elif cn is None:
                cn = 1
            if rn == 0 and cn == 0:
                return a0
            if rn == 0:                       # 整列切片 → 列区域
                if cn < 1 or cn > w:
                    raise FormulaError("INDEX() 下标越界", "ref")
                return ("region", (sh, r1, c1 + cn - 1, r1 + h - 1, c1 + cn - 1))
            if cn == 0:                       # 整行切片 → 行区域
                if rn < 1 or rn > h:
                    raise FormulaError("INDEX() 下标越界", "ref")
                return ("region", (sh, r1 + rn - 1, c1, r1 + rn - 1, c1 + w - 1))
            if rn < 1 or rn > h or cn < 1 or cn > w:
                raise FormulaError("INDEX() 下标越界", "ref")
            return ("cell", (sh, r1 + rn - 1, c1 + cn - 1))
        if name == "XLOOKUP":
            if len(args) < 3:
                raise FormulaError("XLOOKUP() 参数不足")
            needle = self.value_of(args[0])
            if self._region_geo(args[1]) is None or self._region_geo(args[2]) is None:
                raise FormulaError("XLOOKUP() 的查找/返回必须是区域", "value")
            lv = self._flat_values([args[1]])
            rv = self._flat_values([args[2]])
            if len(lv) != len(rv):
                raise FormulaError("XLOOKUP() 查找区域与返回区域尺寸不一致", "value")
            mode = 0
            if len(args) > 4:
                m = self._num_loose(self.value_of(args[4]))
                mode = int(m) if m is not None else 0
            rev = False
            if len(args) > 5:
                s = self._num_loose(self.value_of(args[5]))
                rev = s is not None and int(s) < 0
            order = range(len(lv) - 1, -1, -1) if rev else range(len(lv))
            pos = None
            if mode == 0:
                if isinstance(needle, str) and ("*" in needle or "?" in needle):
                    crit = _Criteria.parse(needle)
                    for i in order:
                        if crit.match(lv[i]):
                            pos = i
                            break
                else:
                    for i in order:
                        if self._eq(lv[i], needle):
                            pos = i
                            break
            else:
                n_needle = self._num_loose(needle)
                best = None
                best_val = None
                for i in order:
                    if self._eq(lv[i], needle):
                        pos = i
                        break
                    n = self._num_strict(lv[i])
                    if n is None or n_needle is None:
                        continue
                    if mode == -1 and n <= n_needle:
                        if best is None or n > best_val:
                            best, best_val = i, n
                    elif mode == 1 and n >= n_needle:
                        if best is None or n < best_val:
                            best, best_val = i, n
                if pos is None:
                    pos = best
            if pos is None:
                if len(args) > 3:
                    return self.value_of(args[3])
                raise FormulaError(f"XLOOKUP 未找到 {_text_of(needle)}", "na")
            v = rv[pos]
            return 0.0 if v is None else v
        if name == "LOOKUP":
            if len(args) < 2:
                raise FormulaError("LOOKUP() 参数不足")
            needle = self.value_of(args[0])
            lv = self._flat_values([args[1]])
            rv = self._flat_values([args[2]]) if len(args) > 2 else lv
            if len(rv) != len(lv):
                raise FormulaError("LOOKUP() 向量长度不一致", "value")
            pos = self._lookup_scan(lv, needle, 1)
            if pos is None:
                raise FormulaError(f"LOOKUP 未找到 {_text_of(needle)}", "na")
            v = rv[pos]
            return 0.0 if v is None else v
        raise FormulaError(f"不支持的查找函数 {name}()")

    # ---- 统计族 ----
    def _stat_fn(self, name, args):
        if name == "SUMPRODUCT":
            if not args:
                raise FormulaError("SUMPRODUCT() 参数不足")

            def sp_num(v):
                n = self._num_loose(v)
                return n if n is not None else 0.0

            cols = [[sp_num(v) for v in self._flat_values([a])] for a in args]
            lens = {len(c) for c in cols if len(c) > 1}
            if len(lens) > 1:
                raise FormulaError("SUMPRODUCT() 各参数尺寸不一致", "value")
            n = max(lens) if lens else 1
            total = 0.0
            for i in range(n):
                p = 1.0
                for c in cols:
                    p *= c[0] if len(c) == 1 else c[i]
                total += p
            return total
        if name in ("LARGE", "SMALL"):
            if len(args) < 2:
                raise FormulaError(f"{name}() 参数不足")
            arr = self._nums([args[0]])
            k = self._num_loose(self.value_of(args[1]))
            if k is None:
                raise FormulaError(f"{name}() 的 k 不是数值", "value")
            k = int(k)
            if k < 1 or k > len(arr):
                raise FormulaError(f"{name}() 的 k 越界: {k}", "num")
            s = sorted(arr, reverse=(name == "LARGE"))
            return float(s[k - 1])
        if name in ("RANK", "RANK.EQ"):
            if len(args) < 2:
                raise FormulaError(f"{name}() 参数不足")
            x = self._num_loose(self.value_of(args[0]))
            if x is None:
                raise FormulaError("RANK() 的值不是数值", "value")
            arr = self._nums([args[1]])
            order = self._num_loose(self.value_of(args[2])) if len(args) > 2 else 0
            asc = order is not None and order != 0
            for i, n in enumerate(sorted(arr, reverse=not asc)):
                if n == x:
                    return float(i + 1)
            raise FormulaError("RANK() 的值不在区域内", "na")
        if name == "PERCENTILE":
            if len(args) < 2:
                raise FormulaError("PERCENTILE() 参数不足")
            arr = sorted(self._nums([args[0]]))
            if not arr:
                raise FormulaError("PERCENTILE() 参数不含数值")
            p = self._num_loose(self.value_of(args[1]))
            if p is None or p < 0 or p > 1:
                raise FormulaError("PERCENTILE() 的 k 必须在 0~1", "num")
            pos = p * (len(arr) - 1)
            lo = int(math.floor(pos))
            hi = min(lo + 1, len(arr) - 1)
            return arr[lo] + (arr[hi] - arr[lo]) * (pos - lo)
        nums = self._nums(args)
        if name == "MEDIAN":
            if not nums:
                raise FormulaError("MEDIAN() 参数不含数值")
            s = sorted(nums)
            n = len(s)
            return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2
        if name == "PRODUCT":
            p = 1.0
            for n in nums:
                p *= n
            return p if nums else 0.0
        if name in ("STDEV", "STDEV.S", "STDEV.P", "VAR", "VAR.S", "VAR.P"):
            if not nums:
                raise FormulaError(f"{name}() 参数不含数值")
            sample = name in ("STDEV", "STDEV.S", "VAR", "VAR.S")
            if sample and len(nums) < 2:
                raise FormulaError(f"{name}() 样本至少需要 2 个数值", "div0")
            mean = sum(nums) / len(nums)
            denom = (len(nums) - 1) if sample else len(nums)
            var = sum((x - mean) ** 2 for x in nums) / denom
            return math.sqrt(var) if name.startswith("STDEV") else var
        raise FormulaError(f"不支持的统计函数 {name}()")

    # ---- 数学族 ----
    def _math_fn(self, name, args):
        def num(i=0):
            if len(args) <= i:
                raise FormulaError(f"{name}() 缺少参数")
            n = self._num_loose(self.value_of(args[i]))
            if n is None:
                raise FormulaError(f"{name}() 参数不是数值", "value")
            return n

        if name == "SQRT":
            x = num()
            if x < 0:
                raise FormulaError("SQRT() 参数为负", "num")
            return math.sqrt(x)
        if name == "EXP":
            return math.exp(num())
        if name == "LN":
            x = num()
            if x <= 0:
                raise FormulaError("LN() 参数必须为正", "num")
            return math.log(x)
        if name in ("LOG", "LOG10"):
            x = num()
            if x <= 0:
                raise FormulaError(f"{name}() 参数必须为正", "num")
            if name == "LOG" and len(args) > 1:
                b = num(1)
                if b <= 0 or b == 1:
                    raise FormulaError("LOG() 底数非法", "num")
                return math.log(x, b)
            return math.log10(x)
        if name == "SIGN":
            x = num()
            return float((x > 0) - (x < 0))
        if name == "TRUNC":
            x = num()
            d = int(num(1)) if len(args) > 1 else 0
            m = 10 ** d
            return math.floor(x * m) / m if x >= 0 else math.ceil(x * m) / m
        if name == "POWER":
            base, e = num(0), num(1)
            try:
                return float(base ** e)
            except (OverflowError, ZeroDivisionError, ValueError):
                raise FormulaError("POWER() 结果非法", "num") from None
        if name == "CEILING":
            x = num()
            sig = num(1) if len(args) > 1 else 1.0
            if sig == 0:
                return 0.0
            if x < 0 and sig > 0:
                raise FormulaError("CEILING() 参数符号不一致", "num")
            return math.ceil(x / sig) * sig
        if name == "FLOOR":
            x = num()
            sig = num(1) if len(args) > 1 else 1.0
            if sig == 0:
                raise FormulaError("FLOOR() 精度为 0", "div0")
            return math.floor(x / sig) * sig
        if name == "MROUND":
            x, mult = num(0), num(1)
            if mult == 0:
                return 0.0
            q = x / mult
            return (math.floor(q + 0.5) if q >= 0 else math.ceil(q - 0.5)) * mult
        if name == "FACT":
            x = int(num())
            if x < 0:
                raise FormulaError("FACT() 参数为负", "num")
            return float(math.factorial(min(x, 170)))
        if name in ("EVEN", "ODD"):
            x = num()
            a = math.ceil(abs(x))
            n = a + (a % 2) if name == "EVEN" else (a if a % 2 == 1 else a + 1)
            return float(n if x >= 0 else -n)
        if name == "SUMSQ":
            return float(sum(n * n for n in self._nums(args)))
        if name in ("PMT", "FV", "PV", "NPV", "IRR"):
            return self._finance(name, args)
        raise FormulaError(f"不支持的数学函数 {name}()")

    # ---- 财务族 ----
    def _finance(self, name, args):
        def val(i, default=None):
            if len(args) <= i:
                return default
            return self._num_loose(self.value_of(args[i]))

        if name == "NPV":
            if len(args) < 2:
                raise FormulaError("NPV() 参数不足")
            rate = val(0)
            if rate is None:
                raise FormulaError("NPV() 折现率不是数值", "value")
            vals = self._nums(args[1:])
            return float(sum(v / (1 + rate) ** (i + 1) for i, v in enumerate(vals)))
        if name == "IRR":
            if not args:
                raise FormulaError("IRR() 参数不足")
            flows = self._nums(args)
            if len(flows) < 2:
                raise FormulaError("IRR() 至少需要 2 期现金流")
            if all(f >= 0 for f in flows) or all(f <= 0 for f in flows):
                raise FormulaError("IRR() 现金流无正负变化", "num")

            def npv(r):
                return sum(f / (1 + r) ** i for i, f in enumerate(flows))

            lo, hi = -0.9999, 10.0
            if npv(lo) * npv(hi) > 0:
                raise FormulaError("IRR() 无法收敛", "num")
            for _ in range(200):
                mid = (lo + hi) / 2
                if npv(lo) * npv(mid) <= 0:
                    hi = mid
                else:
                    lo = mid
            return (lo + hi) / 2
        rate = val(0)
        nper = val(1)
        if rate is None:
            raise FormulaError(f"{name}() 利率不是数值", "value")
        if nper is None:
            raise FormulaError(f"{name}() 期数不是数值", "value")
        nper = int(nper)
        if nper == 0:
            raise FormulaError(f"{name}() 期数为 0", "div0")
        if name == "PMT":
            pv = val(2)
            if pv is None:
                raise FormulaError("PMT() 现值不是数值", "value")
            fv = val(3, 0.0) or 0.0
            typ = val(4, 0.0) or 0.0
            if rate == 0:
                return float(-(pv + fv) / nper)
            g = (1 + rate) ** nper
            r = -(pv * g + fv) * rate / (g - 1)
            return float(r / (1 + rate) if typ else r)
        if name == "FV":
            pmt = val(2)
            if pmt is None:
                raise FormulaError("FV() 每期金额不是数值", "value")
            pv = val(3, 0.0) or 0.0
            typ = val(4, 0.0) or 0.0
            if rate == 0:
                return float(-(pv + pmt * nper))
            g = (1 + rate) ** nper
            return float(-(pv * g + pmt * (1 + rate * typ) * (g - 1) / rate))
        pmt = val(2)                       # PV
        if pmt is None:
            raise FormulaError("PV() 每期金额不是数值", "value")
        fv = val(3, 0.0) or 0.0
        typ = val(4, 0.0) or 0.0
        if rate == 0:
            return float(-(fv + pmt * nper))
        g = (1 + rate) ** nper
        return float(-(fv + pmt * (1 + rate * typ) * (g - 1) / rate) / g)

    # ---- 引用族 ----
    def _ref_fn(self, name, args):
        if name in ("ROW", "COLUMN"):
            if not args:
                return float(self.row if name == "ROW" else self.col)
            geo = self._region_geo(args[0])
            if geo is None:
                raise FormulaError(f"{name}() 参数必须是引用", "value")
            return float(geo[1] if name == "ROW" else geo[2])
        if name in ("ROWS", "COLUMNS"):
            if len(args) != 1:
                raise FormulaError(f"{name}() 需要 1 个参数")
            geo = self._region_geo(args[0])
            if geo is None:
                raise FormulaError(f"{name}() 参数必须是区域", "value")
            return float(geo[3] if name == "ROWS" else geo[4])
        if len(args) < 3:                  # OFFSET
            raise FormulaError("OFFSET() 参数不足")
        geo = self._region_geo(args[0])
        if geo is None:
            raise FormulaError("OFFSET() 第一参数必须是引用", "value")
        sh, r1, c1, h, w = geo

        def onum(i, default):
            if len(args) <= i:
                return default
            n = self._num_loose(self.value_of(args[i]))
            return default if n is None else int(n)

        dr = onum(1, 0)
        dc = onum(2, 0)
        nh = onum(3, h)
        nw = onum(4, w)
        nr, nc = r1 + dr, c1 + dc
        if nr < 1 or nc < 1 or nh < 1 or nw < 1:
            raise FormulaError("OFFSET() 结果越界", "ref")
        if nh == 1 and nw == 1:
            return ("cell", (sh, nr, nc))
        return ("region", (sh, nr, nc, nr + nh - 1, nc + nw - 1))

    # ---- 日期族 ----
    def _date_of(self, v) -> _dt.date:
        """值 → 日期(数值按序列号, 文本按常见格式)。"""
        if isinstance(v, _dt.datetime):
            return v.date()
        if isinstance(v, _dt.date):
            return v
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return _from_serial(float(v))
        if isinstance(v, str):
            d = _parse_date_text(v)
            if d is not None:
                return d
        raise FormulaError(f"不是日期: {_text_of(v)}", "value")

    def _time_of(self, v):
        """值 → (h, m, s); 数值取小数部分(Excel 时间序列)。"""
        if isinstance(v, _dt.datetime):
            return (v.hour, v.minute, v.second)
        if isinstance(v, _dt.time):
            return (v.hour, v.minute, v.second)
        if isinstance(v, _dt.date):
            return (0, 0, 0)
        n = self._num_loose(v)
        if n is None:
            raise FormulaError(f"不是时间: {_text_of(v)}", "value")
        secs = int(round((n - math.floor(n)) * 86400))
        return (secs // 3600 % 24, secs // 60 % 60, secs % 60)

    def _date_fn2(self, name, args):
        if not args:
            raise FormulaError(f"{name}() 缺少参数")
        if name == "DATEVALUE":
            return self._date_of(self.value_of(args[0]))
        if name in ("YEAR", "MONTH", "DAY"):
            d = self._date_of(self.value_of(args[0]))
            return float({"YEAR": d.year, "MONTH": d.month, "DAY": d.day}[name])
        if name in ("HOUR", "MINUTE", "SECOND"):
            h, m, s = self._time_of(self.value_of(args[0]))
            return float({"HOUR": h, "MINUTE": m, "SECOND": s}[name])
        if name == "WEEKDAY":
            d = self._date_of(self.value_of(args[0]))
            t = 1
            if len(args) > 1:
                tv = self._num_loose(self.value_of(args[1]))
                t = int(tv) if tv is not None else 1
            wd = d.weekday()                  # 周一 = 0
            if t == 2:
                return float(wd + 1)
            if t == 3:
                return float(wd)
            return float((wd + 1) % 7 + 1)    # 默认/type=1: 周日 = 1
        if name == "DAYS":
            if len(args) < 2:
                raise FormulaError("DAYS() 需要 2 个参数")
            d2 = self._date_of(self.value_of(args[0]))
            d1 = self._date_of(self.value_of(args[1]))
            return float((d2 - d1).days)
        if name in ("EOMONTH", "EDATE"):
            d = self._date_of(self.value_of(args[0]))
            mv = self._num_loose(self.value_of(args[1])) if len(args) > 1 else 0.0
            months = int(mv) if mv is not None else 0
            return _shift_month(d, months, keep_day=(name == "EDATE"))
        if name == "DATEDIF":
            if len(args) < 3:
                raise FormulaError("DATEDIF() 参数不足")
            d1 = self._date_of(self.value_of(args[0]))
            d2 = self._date_of(self.value_of(args[1]))
            unit = _text_of(self.value_of(args[2])).strip().upper()
            if d2 < d1:
                raise FormulaError("DATEDIF() 结束日期早于起始日期", "num")
            if unit == "D":
                return float((d2 - d1).days)
            if unit == "Y":
                y = d2.year - d1.year
                if (d2.month, d2.day) < (d1.month, d1.day):
                    y -= 1
                return float(y)
            if unit == "M":
                m = (d2.year - d1.year) * 12 + d2.month - d1.month
                if d2.day < d1.day:
                    m -= 1
                return float(m)
            if unit == "YM":
                m = (d2.month - d1.month) % 12
                if d2.day < d1.day:
                    m = (m - 1) % 12
                return float(m)
            if unit == "MD":
                dd = d2.day - d1.day
                if dd < 0:
                    py, pm = ((d2.year, d2.month - 1) if d2.month > 1
                              else (d2.year - 1, 12))
                    dd += calendar.monthrange(py, pm)[1]
                return float(dd)
            if unit == "YD":
                y = d2.year - d1.year
                if y > 0 and (d2.month, d2.day) < (d1.month, d1.day):
                    y -= 1
                try:
                    n1 = d1.replace(year=d2.year - y)
                except ValueError:            # 2/29 → 2/28
                    n1 = d1.replace(year=d2.year - y, day=28)
                return float((d2 - n1).days)
            raise FormulaError(f"DATEDIF() 单位非法: {unit}", "value")
        if name == "TIME":
            if len(args) < 3:
                raise FormulaError("TIME() 参数不足")
            nums = []
            for i in range(3):
                n = self._num_loose(self.value_of(args[i]))
                nums.append(0.0 if n is None else n)
            if any(n < 0 for n in nums):
                raise FormulaError("TIME() 参数为负", "num")
            secs = int(nums[0]) * 3600 + int(nums[1]) * 60 + int(nums[2])
            secs %= 86400
            return _dt.time(secs // 3600, secs // 60 % 60, secs % 60)
        if name in ("NETWORKDAYS", "WORKDAY"):
            if len(args) < 2:
                raise FormulaError(f"{name}() 参数不足")
            start = self._date_of(self.value_of(args[0]))
            holi = set()
            if len(args) > 2:
                for v in self._flat_values([args[2]]):
                    try:
                        holi.add(self._date_of(v))
                    except FormulaError:
                        continue
            if name == "NETWORKDAYS":
                end = self._date_of(self.value_of(args[1]))
                if end < start:
                    start, end = end, start
                n = 0
                d = start
                while d <= end:
                    if d.weekday() < 5 and d not in holi:
                        n += 1
                    d += _dt.timedelta(days=1)
                return float(n)
            dv = self._num_loose(self.value_of(args[1]))
            if dv is None:
                raise FormulaError("WORKDAY() 天数不是数值", "value")
            left = abs(int(dv))
            step = 1 if int(dv) >= 0 else -1
            d = start
            while left > 0:
                d += _dt.timedelta(days=step)
                if d.weekday() < 5 and d not in holi:
                    left -= 1
            return d
        raise FormulaError(f"不支持的日期函数 {name}()")

    # ---- 文本族扩展 ----
    def _format_text(self, v, fmt: str) -> str:
        """TEXT(): 日期格式(含 y/m/d/h/s)走日期渲染, 其余按数值格式。"""
        low = fmt.lower()
        if re.search(r"yy|dd|hh|ss", low) or (re.search(r"[ymdhs]", low)
                                             and not re.search(r"[#0%]", fmt)):
            base = None
            if isinstance(v, _dt.datetime):
                base = v
            elif isinstance(v, _dt.date):
                base = _dt.datetime(v.year, v.month, v.day)
            elif isinstance(v, _dt.time):
                base = _dt.datetime.combine(_DATE_EPOCH, v)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                n = float(v)
                days = int(math.floor(n))
                secs = int(round((n - days) * 86400))
                base = (_dt.datetime.combine(_from_serial(days), _dt.time())
                        + _dt.timedelta(seconds=secs))
            else:
                d = _parse_date_text(_text_of(v))
                if d is not None:
                    base = _dt.datetime(d.year, d.month, d.day)
            if base is not None:
                return _fmt_date_tokens(base, fmt)
            return _text_of(v)
        return _fmt_num_tokens(v, fmt)

    def _text_fn2(self, name, args):
        if name == "CONCAT":
            return "".join(_text_of(v) for v in self._flat_values(args))
        if name == "TEXTJOIN":
            if len(args) < 3:
                raise FormulaError("TEXTJOIN() 参数不足")
            delim = _text_of(self.value_of(args[0]))
            skip_empty = self._truth(self.value_of(args[1]))
            parts = []
            for v in self._flat_values(args[2:]):
                t = _text_of(v)
                if skip_empty and t == "":
                    continue
                parts.append(t)
            return delim.join(parts)
        if name == "TEXT":
            if len(args) < 2:
                raise FormulaError("TEXT() 参数不足")
            return self._format_text(self.value_of(args[0]),
                                     _text_of(self.value_of(args[1])))
        if not args:
            raise FormulaError(f"{name}() 缺少参数")
        if name in ("FIND", "SEARCH"):
            if len(args) < 2:
                raise FormulaError(f"{name}() 参数不足")
            needle = _text_of(self.value_of(args[0]))
            hay = _text_of(self.value_of(args[1]))
            start = 1
            if len(args) > 2:
                sv = self._num_loose(self.value_of(args[2]))
                start = int(sv) if sv is not None else 1
            if start < 1:
                raise FormulaError(f"{name}() 起始位置非法", "value")
            pos = (hay.find(needle, start - 1) if name == "FIND"
                   else hay.casefold().find(needle.casefold(), start - 1))
            if pos < 0:
                raise FormulaError(f"{name}() 未找到 '{needle}'", "value")
            return float(pos + 1)
        s = _text_of(self.value_of(args[0]))
        if name == "REPLACE":
            if len(args) < 4:
                raise FormulaError("REPLACE() 参数不足")
            sv = self._num_loose(self.value_of(args[1]))
            lv = self._num_loose(self.value_of(args[2]))
            start = int(sv) if sv is not None else 0
            ln = int(lv) if lv is not None else 0
            new = _text_of(self.value_of(args[3]))
            if start < 1 or ln < 0:
                raise FormulaError("REPLACE() 起始/长度非法", "value")
            return s[:start - 1] + new + s[start - 1 + ln:]
        if name == "REPT":
            nv = self._num_loose(self.value_of(args[1]))
            if nv is None or nv < 0:
                raise FormulaError("REPT() 次数非法", "value")
            n = int(nv)
            if len(s) * n > 200000:
                raise FormulaError("REPT() 结果过长", "value")
            return s * n
        if name == "EXACT":
            return _text_of(self.value_of(args[0])) == _text_of(self.value_of(args[1]))
        if name == "VALUE":
            n = _parse_number_text(s)
            if n is None:
                raise FormulaError(f"VALUE() 无法解析: '{s}'", "value")
            return float(n)
        if name == "PROPER":
            return _proper(s)
        if name == "CLEAN":
            return "".join(ch for ch in s if ord(ch) >= 32)
        if name == "CHAR":
            n = self._num_loose(self.value_of(args[0]))
            if n is None or int(n) < 1 or int(n) > 255:
                raise FormulaError("CHAR() 参数非法", "value")
            return chr(int(n))
        if name == "CODE":
            if not s:
                raise FormulaError("CODE() 空文本", "value")
            return float(ord(s[0]))
        raise FormulaError(f"不支持的函数 {name}()")

    # ---- 数值函数 ----
    def _num_fn(self, name, args):
        if len(args) < 1:
            raise FormulaError(f"{name}() 缺少参数")
        vals = [self._num_loose(self.value_of(a)) for a in args]
        x = vals[0]
        if x is None:
            raise FormulaError(f"{name}() 参数不是数值")
        if name == "ABS":
            return abs(x)
        if name == "INT":
            return math.floor(x)
        if name == "MOD":
            d = vals[1]
            if not d:
                raise FormulaError("MOD 除数为零")
            return math.fmod(x, d)
        nd = int(vals[1]) if len(vals) > 1 and vals[1] is not None else 0
        if name == "ROUND":
            return float(round(x, nd))
        if name == "ROUNDUP":
            m = 10 ** nd
            return math.ceil(x * m) / m if x >= 0 else math.floor(x * m) / m
        if name == "ROUNDDOWN":
            m = 10 ** nd
            return math.floor(x * m) / m if x >= 0 else math.ceil(x * m) / m
        raise FormulaError(f"不支持的函数 {name}()")

    # ---- 文本函数 ----
    def _text_fn(self, name, args):
        if name == "CONCATENATE":
            return "".join(_text_of(self.value_of(a)) for a in args)
        if not args:
            raise FormulaError(f"{name}() 缺少参数")
        s = _text_of(self.value_of(args[0]))
        if name == "LEN":
            return float(len(s))
        if name == "TRIM":
            return " ".join(s.split())
        if name == "UPPER":
            return s.upper()
        if name == "LOWER":
            return s.lower()
        n = int(self._num_loose(self.value_of(args[1])) or 0)
        if name == "LEFT":
            return s[:max(n, 0)]
        if name == "RIGHT":
            return s[-max(n, 0):] if n else ""
        if name == "MID":
            start = int(self._num_loose(self.value_of(args[1])) or 1)
            ln = int(self._num_loose(self.value_of(args[2])) or 0)
            return s[max(start - 1, 0): max(start - 1, 0) + max(ln, 0)]
        if name == "SUBSTITUTE":
            old = _text_of(self.value_of(args[1]))
            new = _text_of(self.value_of(args[2]))
            return s.replace(old, new)
        raise FormulaError(f"不支持的函数 {name}()")

    # ---- 逻辑 ----
    def _log_fn(self, name, args):
        if name == "NOT":
            return not self._truth(self.value_of(args[0]))
        vals = [self._truth(self.value_of(a)) for a in args]
        return all(vals) if name == "AND" else any(vals)

    # ---- 通用取值辅助 ----
    @staticmethod
    def _truth(v) -> bool:
        """Excel 真值语义: 数值!=0、文本非空且非 'false'、空=False、日期=True。"""
        if isinstance(v, bool):
            return v
        if v is None:
            return False
        if isinstance(v, (int, float)):
            return v != 0
        if isinstance(v, _dt.date):
            return True
        t = str(v).strip().lower()
        if t == "true":
            return True
        if t == "false":
            return False
        return len(t) > 0

    @staticmethod
    def _eq_key(v):
        """等值比较键: (类别, 归一值); 0=数值/日期/空(空按0), 1=文本, 2=布尔。"""
        if v is None:
            return (0, 0.0)
        if isinstance(v, bool):
            return (2, v)
        if isinstance(v, _dt.date):
            return (0, _serial_of(v))
        if isinstance(v, (int, float)):
            return (0, float(v))
        return (1, str(v).casefold())

    def _eq(self, l, r) -> bool:
        """Excel 等值: 文本不区分大小写; 数值/日期按序列号; 数值与文本不相等。"""
        kl, vl = self._eq_key(l)
        kr, vr = self._eq_key(r)
        if kl != kr:
            return False
        return vl == vr

    # ---- 信息类(IS 系) ----
    def _info_fn(self, name, args):
        if len(args) != 1:
            raise FormulaError(f"{name}() 需要 1 个参数")
        v = self.value_of(args[0])      # 单格引用收敛; 多格区域抛错
        if name == "ISNUMBER":
            # bool 是 int 子类必须先排除; Excel 日期即数值序列号 → TRUE
            if isinstance(v, bool):
                return False
            return isinstance(v, (int, float, _dt.date))
        if name == "ISBLANK":
            return v is None
        if name == "ISTEXT":
            return isinstance(v, str)
        if name == "ISNONTEXT":
            return not isinstance(v, str)
        if name == "ISLOGICAL":
            return isinstance(v, bool)
        raise FormulaError(f"不支持的函数 {name}()")

    # ---- 惰性函数(条件 / 错误捕获 / 多路选择) ----
    def call_lazy(self, name, spans):
        if name == "IF":
            if len(spans) not in (2, 3):
                raise FormulaError("IF() 需要 2~3 个参数")
            cond = self._truth(self.value_of(self.eval_span(spans[0])))
            if cond:
                return self.value_of(self.eval_span(spans[1]))
            return (self.value_of(self.eval_span(spans[2]))
                    if len(spans) == 3 else False)
        if name in ("ISERROR", "ISERR", "ISNA"):
            if len(spans) != 1:
                raise FormulaError(f"{name}() 需要 1 个参数")
            try:
                self.value_of(self.eval_span(spans[0]))
            except FormulaError as e:
                if name == "ISNA":
                    return e.code == "na"
                if name == "ISERR":
                    return e.code != "na"
                return True
            return False
        if name == "IFERROR":
            if len(spans) != 2:
                raise FormulaError("IFERROR() 需要 2 个参数")
            try:
                return self.value_of(self.eval_span(spans[0]))
            except FormulaError:
                return self.value_of(self.eval_span(spans[1]))
        if name == "IFNA":
            if len(spans) != 2:
                raise FormulaError("IFNA() 需要 2 个参数")
            try:
                return self.value_of(self.eval_span(spans[0]))
            except FormulaError as e:
                if e.code != "na":
                    raise
                return self.value_of(self.eval_span(spans[1]))
        if name == "IFS":
            if len(spans) < 2 or len(spans) % 2:
                raise FormulaError("IFS() 需要成对的 条件,值 参数")
            for i in range(0, len(spans), 2):
                if self._truth(self.value_of(self.eval_span(spans[i]))):
                    return self.value_of(self.eval_span(spans[i + 1]))
            raise FormulaError("IFS() 没有为真的条件", "na")
        if name == "SWITCH":
            if len(spans) < 3:
                raise FormulaError("SWITCH() 参数不足")
            target = self.value_of(self.eval_span(spans[0]))
            rest = spans[1:]
            for i in range(len(rest) // 2):
                if self._eq(target, self.eval_span(rest[2 * i])):
                    return self.value_of(self.eval_span(rest[2 * i + 1]))
            if len(rest) % 2:                      # 末尾落空值
                return self.value_of(self.eval_span(rest[-1]))
            raise FormulaError("SWITCH() 没有匹配项", "na")
        if name == "CHOOSE":
            if len(spans) < 2:
                raise FormulaError("CHOOSE() 参数不足")
            idx = self._num_loose(self.value_of(self.eval_span(spans[0])))
            if idx is None:
                raise FormulaError("CHOOSE() 序号不是数值", "value")
            i = int(idx)
            if i < 1 or i > len(spans) - 1:
                raise FormulaError(f"CHOOSE() 序号越界: {i}", "value")
            return self.value_of(self.eval_span(spans[i]))
        raise FormulaError(f"不支持的函数 {name}()")


def _cmp_num(op: str, l: float, r: float) -> bool:
    if op == "=":
        return l == r
    if op == "<>":
        return l != r
    if op == ">":
        return l > r
    if op == "<":
        return l < r
    if op == ">=":
        return l >= r
    return l <= r


def _cmp_str(op: str, l: str, r: str) -> bool:
    if op == "=":
        return l == r
    if op == "<>":
        return l != r
    if op == ">":
        return l > r
    if op == "<":
        return l < r
    if op == ">=":
        return l >= r
    return l <= r


def _text_of(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return _num_str(v)
    if isinstance(v, _dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, _dt.date):
        return v.isoformat()
    return str(v)


class _Criteria:
    """SUMIF/COUNTIF 条件: =5、>10、<>0、苹果、苹果*、'=苹果'(等值文本带引号)。"""

    __slots__ = ("op", "val", "num", "pattern")

    def __init__(self, op: str, val, num: bool, pattern: str | None):
        self.op = op          # = <> > >= < <=
        self.val = val        # 数值或文本
        self.num = num        # True=数值条件
        self.pattern = pattern  # fnmatch 模式(None=精确)

    @classmethod
    def parse(cls, raw) -> "_Criteria":
        if isinstance(raw, bool):
            return cls("=", raw, False, None)
        if isinstance(raw, (int, float)):
            return cls("=", float(raw), True, None)
        if isinstance(raw, _dt.date):
            return cls("=", _serial_of(raw), True, None)
        s = _text_of(raw).strip()
        op = "="
        for cand in ("<>", "<=", ">=", "<", ">", "="):
            if s.startswith(cand):
                op = cand
                s = s[len(cand):].strip()
                break
        if not s:
            raise FormulaError(f"条件为空: '{raw}'")
        # 引号包裹 → 文本精确
        if len(s) >= 2 and s[0] in ("\"", "'") and s[-1] == s[0]:
            return cls(op, s[1:-1], False, None)
        try:
            return cls(op, float(s), True, None)
        except ValueError:
            pass
        return cls(op, s, False, re.compile(fnmatch.translate(s)))

    def match(self, v) -> bool:
        if v is None:
            return False
        if isinstance(v, bool):
            return self._match_text("TRUE" if v else "FALSE")
        if isinstance(v, _dt.date):
            return self.num and _cmp_num(self.op, _serial_of(v), self.val)
        if isinstance(v, (int, float)):
            if self.num:
                return _cmp_num(self.op, float(v), self.val)
            return False          # 文本条件 vs 数值格: Excel 视为不匹配
        return self._match_text(v)

    def _match_text(self, s: str) -> bool:
        if not self.num:
            if self.pattern is not None and self.op in ("=", "<>"):
                hit = self.pattern.match(s) is not None
                return hit if self.op == "=" else not hit
            if self.op == "=":
                return s == self.val
            if self.op == "<>":
                return s != self.val
            return _cmp_str(self.op, s, self.val)
        try:                                  # 数值条件 vs 数字文本
            n = float(s.strip())
        except ValueError:
            return False
        return _cmp_num(self.op, n, self.val)


def evaluate(env, sheet: str, row: int, col: int, formula: str):
    """对外入口: 求值公式文本(容忍前导 =),返回 Python 值或抛 FormulaError。"""
    ev = Evaluator(env, sheet, row, col)
    text = formula.strip()
    if text.startswith("="):
        text = text[1:]
    v = ev.run(text)
    # 顶层裸引用(=B4)取值;区域等参与运算会由 _deref 报错,裸多格区域在此兜底
    if _is_ref(v):
        v = ev._deref(v)
    # 数字结果统一浮点 → 交由输出端格式化;date 由 _from_serial 还原
    if isinstance(v, float) and not math.isfinite(v):
        raise FormulaError(f"结果非法: {v}")
    return v

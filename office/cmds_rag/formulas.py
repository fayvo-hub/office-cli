# -*- coding: utf-8 -*-
"""rag prep 内置轻量公式求值器(v2)。

为什么存在: xlsx 公式无缓存值时,v1 保留公式原文并记入 formula_unresolved。
v2 对常用公式直接求值 —— 求值器覆盖真实台账高频场景:

- 算术: + - * / ^ %、一元负、括号、文本连接 &
- 比较: = <> < > <= >=(结果供 IF/条件用)
- 引用: A1 / $A$1 / A1:B3 / A:A 整列 / Sheet1!A1 / '工作 表'!A1:B2(跨 sheet)
- 函数: SUM AVERAGE COUNT COUNTA COUNTBLANK MAX MIN IF AND OR NOT
        SUMIF SUMIFS COUNTIF COUNTIFS
        ROUND ROUNDUP ROUNDDOWN INT ABS MOD
        LEFT RIGHT MID LEN TRIM UPPER LOWER SUBSTITUTE CONCATENATE
        TODAY NOW TRUE FALSE
        ISNUMBER ISBLANK ISTEXT ISNONTEXT ISLOGICAL

原则:
- 能算的算出真值;算不了(未知函数/VLOOKUP/循环引用/外部链接)由调用方保留原文并记录原因
- 日期按 Excel 1900 序列号参与运算,输出还原为 ISO 日期
- 不引入任何第三方依赖,纯标准库

调用方需提供 FormulaEnv:
    cell(sheet, row, col)      -> 该坐标的值:number|str|bool|date|None(空)或抛 FormulaError
    sheet_max_row(sheet)       -> 整列区域上限(取实际最大行)
    sheet_exists(sheet)        -> 工作表是否存在
未捕获 FormulaError 即"算不出",由调用方决定兜底。
"""

from __future__ import annotations

import datetime as _dt
import fnmatch
import math
import re

# 值域统一为: number(float) / text(str) / boolean(bool) / date(datetime.date)
_EMPTY = object()          # 空单元格
_DATE_EPOCH = _dt.date(1899, 12, 30)   # Excel 1900 序列号基准


class FormulaError(Exception):
    """求值失败(引用错误/除零/不支持函数/循环等)。message 面向调用方展示。"""


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
  | (?P<op><>|<=|>=|=|<|>|&|\+|-|\*|/|\^|%|\(|\)|,|:|!|\$)
""", re.X)

# 引用语法里会出现: $ A1 : ! '名称' 等,由 parser 在 token 流上直接拼
_NUM_FUNCS = {"ABS", "INT", "MOD", "ROUND", "ROUNDUP", "ROUNDDOWN"}
_TEXT_FUNCS = {"LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER",
               "SUBSTITUTE", "CONCATENATE"}
_LOG_FUNCS = {"AND", "OR", "NOT"}
_AGG_FUNCS = {"SUM", "AVERAGE", "COUNT", "COUNTA", "COUNTBLANK", "MAX", "MIN",
              "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS"}
_DATE_FUNCS = {"TODAY", "NOW", "DATE"}
_INFO_FUNCS = {"ISNUMBER", "ISBLANK", "ISTEXT", "ISNONTEXT", "ISLOGICAL"}


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
            # 函数调用: ident 紧跟 '('(消解 列名/函数歧义: SUM( 是函数,裸 SUM 是列)
            nxt = self.toks[self.i + 1] if self.i + 1 < len(self.toks) else None
            if nxt is not None and nxt.text == "(":
                self.next()
                self.next()               # 消费 '('
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
        if name == "IF":
            return self._if(args)
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
        def truth(a):
            v = self.value_of(a)
            if isinstance(v, bool):
                return v
            if v is None:
                return False
            if isinstance(v, (int, float)):
                return v != 0
            if isinstance(v, _dt.date):
                return True
            t = v.strip().lower()
            if t == "true":
                return True
            if t == "false":
                return False
            return len(t) > 0
        if name == "NOT":
            return not truth(args[0])
        vals = [truth(a) for a in args]
        return all(vals) if name == "AND" else any(vals)

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

    def _if(self, args):
        if len(args) not in (2, 3):
            raise FormulaError("IF() 需要 2~3 个参数")
        v = self.value_of(args[0])
        if isinstance(v, bool):
            cond = v
        elif isinstance(v, (int, float)):
            cond = v != 0
        elif v is None:
            cond = False
        else:
            cond = _text_of(v).strip().lower() not in ("", "false", "0")
        return self.value_of(args[1]) if cond else (self.value_of(args[2]) if len(args) == 3 else False)


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
            if self.pattern is not None:     # 通配条件
                return self.pattern.match(s) is not None
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

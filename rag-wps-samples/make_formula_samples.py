# -*- coding: utf-8 -*-
"""生成"公式全覆盖"外部 xlsx 物料(openpyxl 直接写公式, 无缓存值)。

用途: 验证 office rag prep 的内置公式求值器。刻意不使用 office-cli 自产,
避免自嗨(公式全部无缓存 → 只有求值器算得出来)。

用法: python rag-wps-samples/make_formula_samples.py [输出目录]
"""
import datetime as dt
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

OUT_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(OUT_DIR, "公式覆盖物料.xlsx")

# 物料: 编码 / 名称 / 单价 / 数量(M-07 起不在等级表 → 触发 IFERROR)
MATERIALS = [
    ("M-01", "深沟球轴承 6204", 32.5, 120),
    ("M-02", "不锈钢内六角螺栓 M8", 1.85, 500),
    ("M-03", "液压油 46#", 28.0, 60),
    ("M-04", "同步带 5M-450", 46.2, 35),
    ("M-05", "气动电磁阀 4V210", 118.0, 24),
    ("M-06", "PLC 扩展模块 FX2N-16EX", 385.0, 12),
    ("M-08", "接近开关 LJ12A3", 42.0, 90),
    ("M-09", "变频器 VFD-M 2.2KW", 1560.0, 6),
    ("M-10", "耐油橡胶垫 3mm", 15.5, 200),
    ("M-12", "时间继电器 ST3PA-B", 68.0, 45),
]
GRADES = [
    ("M-01", "A", 0.97),
    ("M-02", "B", 1.00),
    ("M-03", "B", 1.00),
    ("M-04", "C", 1.05),
    ("M-05", "A", 0.95),
    ("M-06", "A", 0.92),
]

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E1F2")
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center")


def _head(ws, row, labels, widths=None):
    for j, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=j, value=text)
        c.font = BOLD
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[chr(64 + j)].width = w


def build() -> Workbook:
    wb = Workbook()

    # ---------------- Sheet1 物料台账 ----------------
    ws = wb.active
    ws.title = "物料台账"
    ws["A1"] = "2025 年度备件采购台账"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:L1")
    ws["A2"] = "编制部门: 设备动力部"
    ws.merge_cells("A2:L2")
    _head(ws, 3, ["物料编码", "物料名称", "单价(元)", "数量", "金额(元)",
                  "含税金额", "等级", "数量级别", "累计金额", "折扣",
                  "折后金额", "单价级别"],
          [10, 24, 10, 8, 11, 11, 8, 10, 11, 7, 11, 9])
    r0 = 4
    for i, (code, name, price, qty) in enumerate(MATERIALS):
        r = r0 + i
        ws.cell(row=r, column=1, value=code)
        ws.cell(row=r, column=2, value=name)
        pc = ws.cell(row=r, column=3, value=price)
        pc.number_format = "#,##0.00"
        ws.cell(row=r, column=4, value=qty)
        ws.cell(row=r, column=5, value=f"=ROUND(C{r}*D{r},2)")
        ws.cell(row=r, column=6, value=f"=ROUND(E{r}*1.13,2)")
        ws.cell(row=r, column=7,
                value=f'=IFERROR(VLOOKUP(A{r},等级表!$A$2:$C$7,2,FALSE),"未登记")')
        ws.cell(row=r, column=8,
                value=f'=IFS(D{r}>=100,"大批量",D{r}>=50,"中批量",TRUE,"小批量")')
        ws.cell(row=r, column=9, value=f"=ROUND(SUM($E${r0}:E{r}),2)")
        ws.cell(row=r, column=10,
                value=f"=IFERROR(INDEX(等级表!$C$2:$C$7,"
                      f"MATCH(A{r},等级表!$A$2:$A$7,0)),1)")
        ws.cell(row=r, column=11, value=f"=ROUND(E{r}*J{r},2)")
        ws.cell(row=r, column=12,
                value=f'=IF(C{r}>=100,"高",IF(C{r}>=50,"中","低"))')
        for j in range(1, 13):
            ws.cell(row=r, column=j).border = BORDER
    r_end = r0 + len(MATERIALS) - 1
    # 汇总行
    rows = [
        ("金额合计", f"=ROUND(SUM(E{r0}:E{r_end}),2)"),
        ("含税合计", f"=ROUND(SUM(F{r0}:F{r_end}),2)"),
        ("平均单价", f"=ROUND(AVERAGE(C{r0}:C{r_end}),2)"),
        ("最大/最小金额", f"=ROUND(MAX(E{r0}:E{r_end})-MIN(E{r0}:E{r_end}),2)"),
        ("金额中位数", f"=ROUND(MEDIAN(E{r0}:E{r_end}),2)"),
        ("加权平均单价(SUMPRODUCT)",
         f"=ROUND(SUMPRODUCT(C{r0}:C{r_end},D{r0}:D{r_end})"
         f"/SUM(D{r0}:D{r_end}),2)"),
        ("数量≥50 金额合计(SUMIF)", f'=ROUND(SUMIF(D{r0}:D{r_end},">=50",E{r0}:E{r_end}),2)'),
        ("等级=A 金额合计(SUMIFS)",
         f'=ROUND(SUMIFS(E{r0}:E{r_end},G{r0}:G{r_end},"A"),2)'),
        ("等级=A 平均金额(AVERAGEIFS)",
         f'=ROUND(AVERAGEIFS(E{r0}:E{r_end},G{r0}:G{r_end},"A"),2)'),
        ("最大单笔(MAXIFS)", f'=MAXIFS(E{r0}:E{r_end},D{r0}:D{r_end},">=50")'),
        ("最小单笔(MINIFS)", f'=MINIFS(E{r0}:E{r_end},D{r0}:D{r_end},">=50")'),
        ("第 2 高金额(LARGE)", f"=LARGE(E{r0}:E{r_end},2)"),
        ("第 3 低金额(SMALL)", f"=SMALL(E{r0}:E{r_end},3)"),
        ("M-05 金额排名(RANK)", f'=RANK(INDEX(E{r0}:E{r_end},MATCH("M-05",A{r0}:A{r_end},0)),E{r0}:E{r_end})'),
        ("金额标准差(STDEV)", f"=ROUND(STDEV(E{r0}:E{r_end}),2)"),
        ("90 分位(PERCENTILE)", f"=ROUND(PERCENTILE(E{r0}:E{r_end},0.9),2)"),
        ("单价>50 种类(COUNTIF)", f'=COUNTIF(C{r0}:C{r_end},">50")'),
        ("非 A 级数量(COUNTIF <>)", f'=COUNTIF(G{r0}:G{r_end},"<>A")'),
        ("M-07 名称(XLOOKUP 默认值)",
         f'=XLOOKUP("M-07",A{r0}:A{r_end},B{r0}:B{r_end},"无此物料")'),
        ("M-06 名称(INDEX+MATCH)",
         f'=INDEX(B{r0}:B{r_end},MATCH("M-06",A{r0}:A{r_end},0))'),
        ("前 5 行金额(SUM+OFFSET)", f"=ROUND(SUM(OFFSET(E{r0},0,0,5,1)),2)"),
        ("明细行数(ROWS)", f"=ROWS(A{r0}:A{r_end})"),
        ("月供测算(PMT)", "=ROUND(PMT(0.05/12,36,60000),2)"),
        ("前 4 笔净现值(NPV)", f"=ROUND(NPV(0.08,E{r0}:E{r0+3}),2)"),
        ("月份文本(TEXT)", f'=TEXT(DATE(2025,3,1),"yyyy年m月")'),
        ("除零兜底(IFERROR)", '=IFERROR(1/0,"已兜底")'),
        ("查无编码(IFNA)", '=IFNA(VLOOKUP("M-99",A4:C13,2,FALSE),"查无编码")'),
    ]
    r = r_end + 2
    for label, formula in rows:
        ws.cell(row=r, column=1, value=label).font = BOLD
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws.cell(row=r, column=3, value=formula)
        r += 1

    # ---------------- Sheet2 等级表 ----------------
    ws2 = wb.create_sheet("等级表")
    _head(ws2, 1, ["物料编码", "等级", "折扣"], [12, 8, 8])
    for i, (code, grade, disc) in enumerate(GRADES, start=2):
        ws2.cell(row=i, column=1, value=code)
        ws2.cell(row=i, column=2, value=grade)
        ws2.cell(row=i, column=3, value=disc)

    # ---------------- Sheet3 统计(日期/文本/财务) ----------------
    ws3 = wb.create_sheet("统计")
    _head(ws3, 1, ["项目", "值", "备注"], [26, 18, 30])
    d1 = dt.date(2025, 3, 17)
    d2 = dt.date(2026, 1, 31)
    lines = [
        ("合同起始日", d1, "", None),
        ("合同到期日", d2, "", None),
        ("合同天数(DAYS)", "=DAYS(B3,B2)", "日历天数", None),
        ("工作日数(NETWORKDAYS)", "=NETWORKDAYS(B2,B3)", "排除周末", None),
        ("3 个月后月末(EOMONTH)", "=EOMONTH(B2,3)", "日期序列", "yyyy-mm-dd"),
        ("到期前 45 个工作日(WORKDAY)", "=WORKDAY(B2,45)", "工期推算", "yyyy-mm-dd"),
        ("已执行年数(DATEDIF Y)", '=DATEDIF(B2,B3,"Y")', "整年", None),
        ("已执行月数(DATEDIF M)", '=DATEDIF(B2,B3,"M")', "整月", None),
        ("起始年份(YEAR)", "=YEAR(B2)", "", None),
        ("起始月份(MONTH)", "=MONTH(B2)", "", None),
        ("起始日(WEEKDAY)", "=WEEKDAY(B2,2)", "周一=1", None),
        ("月份文本(TEXT)", '=TEXT(B2,"yyyy年m月")', "", None),
        ("拼接(TEXTJOIN)", '=TEXTJOIN("-",TRUE,"合同",B4,"天")', "", None),
        ("金额平方和(SUMSQ)", "=ROUND(SUMSQ(物料台账!E4:E13),2)", "", None),
        ("单价几何均值(GEOMEAN 不支持→SQRT)",
         "=ROUND(SQRT(PRODUCT(物料台账!C4:C6)),2)", "乘积开方", None),
        ("金额取整(CEILING)", "=CEILING(物料台账!E4,10)", "", None),
        ("金额截断(TRUNC)", "=TRUNC(物料台账!E4,0)", "", None),
        ("单价四舍五入到角(MROUND)", "=MROUND(物料台账!C5,0.1)", "", None),
        ("金额符号(SIGN)", "=SIGN(物料台账!E4-1000)", "正=1 负=-1", None),
        ("编号示例(REPT/PROPER)", '=PROPER("no.")&REPT("0",3)&"7"', "", None),
        ("查字符(CHAR 65)", "=CHAR(65)", "", None),
        ("字符码(CODE A)", '=CODE("A")', "", None),
        ("文本转数值(VALUE)", '=VALUE("¥1,234.50")', "", None),
        ("首段文本(FIND/LEFT)", '=LEFT(物料台账!B4,FIND(" ",物料台账!B4)-1)', "", None),
    ]
    r = 2
    for item, val, note, fmt in lines:
        ws3.cell(row=r, column=1, value=item)
        c = ws3.cell(row=r, column=2, value=val)
        if fmt:
            c.number_format = fmt
        if note:
            ws3.cell(row=r, column=3, value=note)
        r += 1

    # ---------------- Sheet4 说明 ----------------
    ws4 = wb.create_sheet("说明")
    ws4["A1"] = "本文件由 openpyxl 生成, 公式无缓存值 —— 用于验证 office rag prep 的内置求值器。"
    ws4.merge_cells("A1:H1")
    ws4["A2"] = "数组表达式(如 SUMPRODUCT((条件)*(区域)))不在支持范围, 预期记为 formula_unresolved。"
    ws4.merge_cells("A2:H2")
    return wb


def main() -> None:
    wb = build()
    wb.save(OUT)
    print("生成:", OUT)


if __name__ == "__main__":
    main()

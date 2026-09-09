# -*- coding: utf-8 -*-
"""用 WPS Office(COM)生成 RAG 功能的真实复杂测试物料(非 office-cli 自产)。

物料:
- 采购台账.xlsx       四张表: 多级表头+公式/百分比/日期/合并/小计/多块/表尾注/隐藏表
- 采购台账_legacy.xls 同结构的 .xls 老二进制格式(验证自动升级读取)

用法: python make_wps_samples.py   (需要本机安装 WPS Office)
"""
import datetime as _dt
import os
import sys

try:
    import win32com.client
except ImportError:  # pragma: no cover
    print("缺少 pywin32: pip install pywin32")
    sys.exit(2)

HERE = os.path.dirname(os.path.abspath(__file__))
XLSX_OUT = os.path.join(HERE, "采购台账.xlsx")
XLS_OUT = os.path.join(HERE, "采购台账_legacy.xls")

# Sheet1 数据: (日期, 供应商, 品名, 单价, 数量)
ROWS = [
    (_dt.datetime(2024, 1, 12), "无锡振华轴承厂", "深沟球轴承 6204", 11.8, 300),
    (_dt.datetime(2024, 1, 19), "上海精工机电", "角接触球轴承 7005", 45.2, 120),
    (_dt.datetime(2024, 2, 6), "宁波传动部件", "圆柱滚子轴承 NU205", 52.0, 80),
    (_dt.datetime(2024, 2, 15), "无锡振华轴承厂", "深沟球轴承 6205", 13.6, 250),
    (_dt.datetime(2024, 2, 28), "洛阳轴承集团", "深沟球轴承 6204", 12.2, 200),
    (_dt.datetime(2024, 3, 5), "上海精工机电", "角接触球轴承 7005", 46.0, 90),
    (_dt.datetime(2024, 3, 12), "宁波传动部件", "圆柱滚子轴承 NU205", 53.5, 60),
    (_dt.datetime(2024, 3, 18), "无锡振华轴承厂", "深沟球轴承 6204", 12.0, 350),
    (_dt.datetime(2024, 3, 25), "洛阳轴承集团", "深沟球轴承 6205", 13.9, 180),
    (_dt.datetime(2024, 4, 2), "上海精工机电", "角接触球轴承 7005", 45.8, 110),
]


def col(rng):  # COM 区域合并
    rng.Merge()
    return rng


def build_sheet1(ws, full: bool):
    """主表「采购台账」: full=False 时只输出到合计行(供 .xls 版)。"""
    ws.Cells(1, 1).Value = "华东供应链 2024 年度采购台账"
    col(ws.Range(ws.Cells(1, 1), ws.Cells(1, 10)))
    ws.Cells(2, 1).Value = "(单位:人民币元;本表由采购部每月 5 日前报送)"
    col(ws.Range(ws.Cells(2, 1), ws.Cells(2, 10)))

    # ---- 两级表头(行 4/5): 纵向 序号/日期/供应商/品名/状态 + 横向组 核算/价款 ----
    heads = [(1, "序号"), (2, "采购日期"), (3, "供应商"), (4, "品名"), (10, "到货状态")]
    for colc, t in heads:
        ws.Cells(4, colc).Value = t
        col(ws.Range(ws.Cells(4, colc), ws.Cells(5, colc)))
    ws.Cells(4, 5).Value = "核算"
    col(ws.Range(ws.Cells(4, 5), ws.Cells(4, 6)))
    ws.Cells(4, 7).Value = "价款(元)"
    col(ws.Range(ws.Cells(4, 7), ws.Cells(4, 9)))
    ws.Cells(5, 5).Value = "单价"
    ws.Cells(5, 6).Value = "数量(件)"
    ws.Cells(5, 7).Value = "金额"
    ws.Cells(5, 8).Value = "税率"
    ws.Cells(5, 9).Value = "税额"
    # 列 10(J)被「到货状态」纵向块占据: 行 5 的 J 属块,无需再写

    for i, (dt, vendor, name, price, qty) in enumerate(ROWS):
        r = 6 + i
        ws.Cells(r, 1).Value = i + 1
        # pywin32 将 naive datetime 按 UTC 传 COM DATE(减 8h),补回东八区使文件落整点
        ws.Cells(r, 2).Value = dt + _dt.timedelta(hours=8)
        ws.Cells(r, 2).NumberFormat = "yyyy-mm-dd"
        ws.Cells(r, 3).Value = vendor
        ws.Cells(r, 4).Value = name
        ws.Cells(r, 5).Value = price
        ws.Cells(r, 6).Value = qty
        ws.Cells(r, 7).Formula = f"=ROUND(E{r}*F{r},2)"
        ws.Cells(r, 8).Value = 0.13
        ws.Cells(r, 8).NumberFormat = "0%"
        ws.Cells(r, 9).Formula = f"=ROUND(G{r}*H{r},2)"
        ws.Cells(r, 10).Formula = f'=IF(F{r}>=100,"批量","零散")'

    r = 16
    ws.Cells(r, 1).Value = "合  计"
    col(ws.Range(ws.Cells(r, 1), ws.Cells(r, 6)))
    ws.Cells(r, 7).Formula = "=SUM(G6:G15)"
    ws.Cells(r, 9).Formula = "=SUM(I6:I15)"
    if not full:
        return

    # ---- 第二块: 不合格品退回登记 ----
    ws.Cells(18, 1).Value = "不合格品退回登记(2024 年一季度)"
    col(ws.Range(ws.Cells(18, 1), ws.Cells(18, 6)))
    for c, t in enumerate(["退回日期", "品名", "数量", "退回原因", "处理结果"], start=1):
        ws.Cells(19, c).Value = t
    returns = [
        ("2024-01-20", "深沟球轴承 6204", 6, "外圈滚道划伤", "退回供应商换货"),
        ("2024-02-08", "圆柱滚子轴承 NU205", 2, "保持架变形", "供应商补发"),
        ("2024-03-16", "角接触球轴承 7005", 4, "游隙超差", "质检复核后让步接收"),
    ]
    for i, (dt, name, qty, why, how) in enumerate(returns):
        r = 20 + i
        ws.Cells(r, 1).Value = dt
        ws.Cells(r, 2).Value = name
        ws.Cells(r, 3).Value = qty
        ws.Cells(r, 4).Value = why
        ws.Cells(r, 5).Value = how

    ws.Cells(24, 1).Value = "注:退回品由质检部复核处理;本表单价均为不含税到厂价。"
    col(ws.Range(ws.Cells(24, 1), ws.Cells(24, 10)))


def build_sheet2(ws):
    ws.Cells(1, 1).Value = "物料价格参考与年度汇总"
    col(ws.Range(ws.Cells(1, 1), ws.Cells(1, 3)))
    ws.Cells(2, 1).Value = "(数据由供应链中心维护,更新至 2024-03)"
    col(ws.Range(ws.Cells(2, 1), ws.Cells(2, 3)))
    for c, t in enumerate(["品名", "基准单价(元)", "备注"], start=1):
        ws.Cells(4, c).Value = t
    refs = [
        ("深沟球轴承 6204", 11.8, "月均采购 ≥200 件可下浮 3%"),
        ("深沟球轴承 6205", 13.6, "含税价另加 13%"),
        ("角接触球轴承 7005", 45.2, "免维护型,货期 20 天"),
        ("圆柱滚子轴承 NU205", 52.0, "进口钢材,货期 30 天"),
    ]
    for i, (name, price, note) in enumerate(refs):
        r = 5 + i
        ws.Cells(r, 1).Value = name
        ws.Cells(r, 2).Value = price
        ws.Cells(r, 3).Value = note

    ws.Cells(11, 1).Value = "2024 年度采购汇总(截至 3 月)"
    col(ws.Range(ws.Cells(11, 1), ws.Cells(11, 3)))
    for c, t in enumerate(["品类", "采购金额(元)", "金额占比"], start=1):
        ws.Cells(12, c).Value = t
    fams = [("深沟球轴承", 13), ("角接触球轴承", 14), ("圆柱滚子轴承", 15)]
    for fam, r in fams:
        ws.Cells(r, 1).Value = fam
        ws.Cells(r, 2).Formula = (
            f'=SUMIF(采购台账!$D$6:$D$15,"{fam}*",采购台账!$G$6:$G$15)')
        ws.Cells(r, 3).Formula = f"=B{r}/$B$16"
        ws.Cells(r, 3).NumberFormat = "0.0%"
    ws.Cells(16, 1).Value = "合计"
    ws.Cells(16, 2).Formula = "=SUM(B13:B15)"
    ws.Cells(17, 1).Value = "注:占比按采购金额合计计算;表内金额均为不含税口径。"


def build_sheet3(ws):
    paras = [
        "编制说明",
        "一、本台账由采购部依据 ERP 采购订单与到货验收单逐笔登记,月结后与财务对账。",
        "二、“核算”组内的金额、税额为公式自动计算,严禁手工改数;税率统一 13%。",
        "三、退回登记仅记录检验判定为不合格的批次,处理结果由质检部在 3 个工作日内回填。",
    ]
    for i, t in enumerate(paras, start=1):
        ws.Cells(i, 1).Value = t
        if i >= 2:   # 说明正文占整行宽
            col(ws.Range(ws.Cells(i, 1), ws.Cells(i, 8)))


def make_xlsx():
    app = win32com.client.DispatchEx("KET.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Add()
        ws1 = wb.Worksheets(1)
        ws1.Name = "采购台账"
        build_sheet1(ws1, full=True)
        ws2 = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        ws2.Name = "价格参考与汇总"
        build_sheet2(ws2)
        ws3 = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        ws3.Name = "编制说明"
        build_sheet3(ws3)
        ws4 = wb.Worksheets.Add(None, wb.Worksheets(wb.Worksheets.Count))
        ws4.Name = "汇总底稿"
        ws4.Visible = 0          # 隐藏表: 期望被跳过
        ws4.Cells(1, 1).Value = "底稿"
        ws4.Cells(2, 1).Value = 1
        app.Calculate()
        wb.SaveAs(XLSX_OUT, 51)
        wb.Close(False)
    finally:
        app.Quit()
    print("written:", XLSX_OUT, os.path.getsize(XLSX_OUT), "bytes")


def make_xls():
    app = win32com.client.DispatchEx("KET.Application")
    try:
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Add()
        ws1 = wb.Worksheets(1)
        ws1.Name = "采购台账"
        build_sheet1(ws1, full=False)
        app.Calculate()
        wb.SaveAs(XLS_OUT, 56)   # 56 = xlExcel8(.xls 老格式)
        wb.Close(False)
    finally:
        app.Quit()
    print("written:", XLS_OUT, os.path.getsize(XLS_OUT), "bytes")


if __name__ == "__main__":
    make_xlsx()
    make_xls()

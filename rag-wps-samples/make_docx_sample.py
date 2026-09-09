# -*- coding: utf-8 -*-
"""用 python-docx(非 office-cli)生成 Word 测试物料: 含合并单元格的评审纪要。

物料: 评审纪要.docx —— 标题/正文/带 gridSpan+vMerge 合并表格/编号列表。
"""
import os

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "评审纪要.docx")


def main():
    doc = Document()
    doc.add_heading("自动化产线改造项目招标评审会议纪要", 0)
    doc.add_heading("一、会议概况", 1)
    doc.add_paragraph(
        "会议时间:2024 年 6 月 18 日 14:00;会议地点:集团 3 号会议室。"
        "主持人:王副总;参加人员:评标委员会全体 7 名成员及各投标单位代表。")
    doc.add_paragraph(
        "本次招标共收到有效投标文件 3 份,经资格预审均合格。评标采用综合评分法,"
        "价格权重 40%、技术与商务权重 60%,全程录音录像并留存评审记录。")

    doc.add_heading("二、评审结果", 1)
    t1 = doc.add_table(rows=4, cols=6)
    t1.style = "Table Grid"
    t1.alignment = WD_TABLE_ALIGNMENT.CENTER
    heads = ["序号", "投标单位", "报价(万元)", "技术评分", "商务评分", "综合意见"]
    for j, h in enumerate(heads):
        t1.cell(0, j).text = h
    data = [
        ("1", "中科自动化装备", 1280.0, 92.5, 88, ""),
        ("2", "华东智能科技", 1196.0, 87.0, 91, ""),
        ("3", "华信机电工程", 1330.0, 90.0, 79, ""),
    ]
    for i, row in enumerate(data, start=1):
        for j, v in enumerate(row):
            if j == 5:
                continue          # 末列留给纵向合并
            t1.cell(i, j).text = str(v)
    # 末列纵向合并 3 行(评审结论)
    merged = t1.cell(1, 5).merge(t1.cell(3, 5))
    merged.text = "综合排序第一,建议中标"
    p = merged.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph("注:报价含税;技术评分由 5 名专家取平均。")

    doc.add_heading("三、后续安排", 1)
    doc.add_paragraph("中标结果公示 5 个工作日,无异议后发出中标通知书;",
                      style="List Number")
    doc.add_paragraph("合同谈判定于 6 月 28 日,重点核对交付节点与质保条款;",
                      style="List Number")
    doc.add_paragraph("未中标单位保证金于 7 月 10 日前原路退回。",
                      style="List Number")
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para.add_run("评标委员会(签章)")
    para2 = doc.add_paragraph()
    para2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    para2.add_run("2024 年 6 月 18 日")

    doc.save(OUT)
    print("written:", OUT, os.path.getsize(OUT), "bytes")


if __name__ == "__main__":
    main()

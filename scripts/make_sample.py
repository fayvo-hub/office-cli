"""生成端到端测试用的样本文件(直接使用 openpyxl 等底层库,不走 office 命令,保证样本真实可控)。

产物(samples/ 目录):
  demo.xlsx       多 sheet:公式/日期/合并/样式/中文
  demo_cached.xlsx 同上,但公式格注入了计算缓存值(模拟 Excel 保存的文件)
  demo_pivot.xlsx 含一个数据透视表定义
  gbk.csv         中文 GBK 编码 CSV(Excel 另存风格)
  logo.png        小图片(image 命令测试用)
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.pivot.cache import CacheDefinition, CacheSource, WorksheetSource
from openpyxl.pivot.table import Location, TableDefinition
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(os.path.dirname(HERE), "samples")


def make_demo_xlsx() -> str:
    os.makedirs(SAMPLES, exist_ok=True)
    path = os.path.join(SAMPLES, "demo.xlsx")
    wb = Workbook()

    # ---------- 销售表:公式 + 日期 + 样式 + 合并 ----------
    ws = wb.active
    ws.title = "销售"
    headers = ["月份", "销量", "单价", "金额", "日期"]
    ws.append(headers)
    for i in range(12):
        row = i + 2
        ws.append([
            f"2024-{i+1:02d}",
            100 + i * 25,
            round(9.9 + i * 0.5, 1),
            f"=B{row}*C{row}",
            datetime(2024, i + 1, 15),
        ])
    # 表头样式
    fill = PatternFill("solid", fgColor="FF4472C4")
    font = Font(bold=True, color="FFFFFFFF")
    thin = Border(*[Side(style="thin", color="FFB0B0B0")] * 4)
    for col in range(1, 6):
        c = ws.cell(row=1, column=col)
        c.fill, c.font = fill, font
        c.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = 12
    # 日期列格式
    for i in range(12):
        ws.cell(row=i + 2, column=5).number_format = "yyyy-mm-dd"
    # 合计区 + 合并
    ws["A15"] = "销量合计(仅A列示例)"
    ws["B15"] = "=SUM(B2:B13)"
    ws["B15"].number_format = "#,##0"
    ws.merge_cells("A16:B16")
    ws["A16"] = "金额总计"
    ws["C16"] = "=SUM(D2:D13)"
    ws["C16"].font = Font(bold=True)
    # 给数据区加细边框(便于验证样式读回)
    for row in ws.iter_rows(min_row=1, max_row=13, min_col=1, max_col=5):
        for c in row:
            c.border = thin
    # 一个注释 + 超链接
    from openpyxl.comments import Comment
    ws["B2"].comment = Comment("1 月销量偏低", "tester")
    ws["A1"].hyperlink = "https://example.com/help"
    ws.freeze_panes = "A2"

    # ---------- 人员表:纯中文数据 ----------
    ws2 = wb.create_sheet("人员")
    ws2.append(["姓名", "城市", "部门", "入职日期"])
    people = [
        ("张三", "北京", "销售部", "2021-03-01"),
        ("李四", "上海", "研发部", "2022-07-15"),
        ("王五", "广州", "市场部", "2020-01-20"),
        ("赵六", "深圳", "销售部", "2023-11-02"),
        ("钱七", "北京", "研发部", "2019-05-30"),
    ]
    for name, city, dept, date_s in people:
        ws2.append([name, city, dept, datetime.strptime(date_s, "%Y-%m-%d")])
        ws2.cell(row=ws2.max_row, column=4).number_format = "yyyy-mm-dd"
    ws2.append(["孙八", "成都", "市场部"])  # 无日期,测试空尾列

    # ---------- 空表 ----------
    wb.create_sheet("空表")

    wb.save(path)

    # 复制一份注入公式缓存值 -> demo_cached.xlsx(模拟 Excel 保存)
    cached_path = os.path.join(SAMPLES, "demo_cached.xlsx")
    _inject_formula_cache(path, cached_path)

    # 透视表样本(openpyxl 3.1 可写回已有透视表定义)
    make_pivot_sample()
    return path


def _inject_formula_cache(src: str, dst: str) -> None:
    """给公式单元格注入缓存值 <v>,模拟 Excel 保存的文件(openpyxl 写文件不带缓存)。"""
    vals = {}  # ref -> 缓存值
    for row in range(2, 14):
        vals[f"D{row}"] = (100 + (row - 2) * 25) * round(9.9 + (row - 2) * 0.5, 1)
    vals["B15"] = sum(100 + i * 25 for i in range(12))
    vals["C16"] = sum(vals[f"D{r}"] for r in range(2, 14))
    n = 0

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w") as zout:
        for item in zin.namelist():
            data = zin.read(item)
            if item.endswith("worksheets/sheet1.xml"):
                text = data.decode("utf-8")
                for ref, v in vals.items():
                    # <c r="D2"><f>B2*C2</f><v></v></c> -> 填入缓存值 <v>xx</v>
                    pat = re.compile(
                        rf'(<c r="{ref}"[^>]*>)(<f>[^<]*</f>)<v></v>(</c>)')
                    text, k = pat.subn(
                        lambda m: f"{m.group(1)}{m.group(2)}<v>{v}</v>{m.group(3)}", text)
                    n += k
                data = text.encode("utf-8")
            zout.writestr(item, data)
    print(f"injected {n} cached values -> {os.path.basename(dst)}")


def make_pivot_sample() -> None:
    wb = Workbook()
    src = wb.active
    src.title = "源数据"
    src.append(["产品", "地区", "金额"])
    for i in range(30):
        src.append([f"产品{i % 5 + 1}", ("华东", "华南", "华北", "西南")[i % 4], (i + 1) * 100])
    ws = wb.create_sheet("透视")
    cache = CacheDefinition(
        cacheSource=CacheSource(type="worksheet",
                                worksheetSource=WorksheetSource(ref="$A$1:$C$31")),
        refreshOnLoad=True,
    )
    p = TableDefinition(name="产品透视", cacheId=1,
                        location=Location(ref="A3", firstHeaderRow=1,
                                          firstDataRow=1, firstDataCol=1),
                        dataCaption="求和项:金额")
    p.cache = cache
    ws.add_pivot(p)
    wb.save(os.path.join(SAMPLES, "demo_pivot.xlsx"))
    print("pivot sample saved")


def make_csvs() -> None:
    import csv as csvmod
    rows = [["姓名", "城市", "销售额", "是否达标"],
            ["张三", "北京", "1200.5", "TRUE"],
            ["李四", "上海", "800", "FALSE"],
            ["王五", "广州", "1500", "TRUE"]]
    # GBK(Excel 中文另存常见编码)
    with open(os.path.join(SAMPLES, "gbk.csv"), "w", encoding="gb18030",
              newline="") as fh:
        csvmod.writer(fh).writerows(rows)
    # UTF-8 with BOM
    with open(os.path.join(SAMPLES, "utf8.csv"), "w", encoding="utf-8-sig",
              newline="") as fh:
        csvmod.writer(fh).writerows(rows)


def make_logo() -> None:
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        return
    img = Image.new("RGB", (60, 24), "#4472C4")
    ImageDraw.Draw(img).text((8, 4), "AI", fill="white")
    img.save(os.path.join(SAMPLES, "logo.png"))


def make_sample_docx() -> str:
    """用 office CLI 的 word write 生成 sample.docx(样本本身即产物)。"""
    import json
    import subprocess
    import sys
    from pathlib import Path

    blocks = [
        {"type": "h", "level": 1, "text": "Word 样本文档"},
        {"type": "p", "text": "这是由 office word write 生成的样本,包含下列元素。"},
        {"type": "h", "level": 2, "text": "混合段落"},
        {"type": "p", "runs": [
            {"text": "普通文本+"},
            {"text": "加粗", "bold": True},
            {"text": "与"},
            {"text": "红字", "color": "FF0000"},
            {"text": "。"}]},
        {"type": "code", "text": "def add(a, b):\n    return a + b"},
        {"type": "quote", "text": "引用样式段落。"},
        {"type": "h", "level": 2, "text": "表格"},
        {"type": "table", "header": True,
         "rows": [["姓名", "城市"], ["张三", "北京"], ["李四", "上海"]]},
        {"type": "h", "level": 2, "text": "图片"},
        {"type": "img", "path": str(Path(SAMPLES) / "logo.png"), "width": 6},
    ]
    cfg = os.path.join(SAMPLES, "..", ".test-tmp", "sample_blocks.json")
    os.makedirs(os.path.dirname(cfg), exist_ok=True)
    with open(cfg, "w", encoding="utf-8") as fh:
        json.dump({"blocks": blocks}, fh, ensure_ascii=False)
    out = os.path.join(SAMPLES, "sample.docx")
    if os.path.exists(out):
        os.remove(out)  # word write --create 对已存在文件是追加语义, 先清旧
    proc = subprocess.run(
        [sys.executable, "-m", "office", "word", "write", "-f", out,
         "--data-file", cfg, "--create"],
        capture_output=True, text=True, encoding="utf-8",
        cwd=os.path.dirname(HERE), timeout=60)
    os.remove(cfg)
    if proc.returncode != 0:
        raise RuntimeError(f"word write 失败: {proc.stdout} {proc.stderr}")
    print("sample.docx saved")
    return out


def make_sample_pptx() -> str:
    """用 python-pptx 生成 sample.pptx:标题页/要点(含层级+备注)/表格/图片。"""
    from pptx import Presentation
    from pptx.util import Inches

    out = os.path.join(SAMPLES, "sample.pptx")
    prs = Presentation()

    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "演示文稿样本"

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "进展摘要"
    body = s.placeholders[1].text_frame
    body.text = "需求评审完成"
    p = body.add_paragraph()
    p.text = "覆盖 3 个部门"
    p.level = 1
    p = body.add_paragraph()
    p.text = "原型已确认"
    s.notes_slide.notes_text_frame.text = "强调时间点"

    s = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    s.shapes.title.text = "数据一览"
    gt = s.shapes.add_table(3, 3, Inches(0.6), Inches(1.8),
                            Inches(8.8), Inches(1.2)).table
    for r, row in enumerate((("指标", "Q1", "Q2"),
                             ("收入", "100", "150"),
                             ("用户", "10", "22"))):
        for c, v in enumerate(row):
            gt.cell(r, c).text = str(v)

    s = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    box = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(8.8), Inches(1))
    box.text_frame.text = "感谢"
    s.shapes.add_picture(os.path.join(SAMPLES, "logo.png"),
                         Inches(4.2), Inches(2.6), width=Inches(1.5))

    prs.save(out)
    print("sample.pptx saved")
    return out


def make_legacy_samples() -> None:
    """用 WPS 把新格式另存为 .xls/.doc/.ppt 老格式样本(WPS 可用时)。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        print("skip legacy samples (no pywin32)")
        return
    pythoncom.CoInitialize()
    xls_dst = os.path.join(SAMPLES, "legacy.xls")
    doc_dst = os.path.join(SAMPLES, "legacy.doc")
    ppt_dst = os.path.join(SAMPLES, "legacy.ppt")
    try:
        app = win32com.client.DispatchEx("KET.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            wb = app.Workbooks.Open(os.path.join(SAMPLES, "demo.xlsx"),
                                    ReadOnly=True)
            wb.SaveAs(xls_dst, FileFormat=56)  # 56 = .xls
            wb.Close(False)
        finally:
            app.Quit()
        print("legacy.xls saved")
    except Exception as e:
        print(f"skip legacy.xls (WPS KET 不可用: {e})")
    try:
        app = win32com.client.DispatchEx("KWPS.Application")
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            d = app.Documents.Open(os.path.join(SAMPLES, "sample.docx"),
                                   ReadOnly=True)
            d.SaveAs(doc_dst, FileFormat=0)  # 0 = .doc
            d.Close(False)
        finally:
            app.Quit()
        print("legacy.doc saved")
    except Exception as e:
        print(f"skip legacy.doc (WPS KWPS 不可用: {e})")
    try:
        import time
        app = None
        pres = None
        # WPS 三个组件共享宿主进程,前两个 Quit 后 wps.exe 延迟退出;
        # 期间激活 KWPP 会失败,故重试等待残留进程退出
        for attempt in range(6):
            try:
                app = win32com.client.DispatchEx("KWPP.Application")
                app.Visible = False
                app.DisplayAlerts = 0
                break
            except Exception:
                app = None
                if attempt < 5:
                    time.sleep(2)
        if app is None:
            raise RuntimeError("KWPP 连续 6 次启动失败(WPS 残留进程未退出?)")
        pres = app.Presentations.Open(
            os.path.join(SAMPLES, "sample.pptx"), ReadOnly=True)
        pres.SaveAs(ppt_dst, 1)  # 1 = .ppt
        pres.Close()
        pres = None
        print("legacy.ppt saved")
    except Exception as e:
        print(f"skip legacy.ppt (WPS KWPP 不可用: {e})")
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass


if __name__ == "__main__":
    make_demo_xlsx()
    make_csvs()
    make_logo()
    make_sample_docx()
    make_sample_pptx()
    make_legacy_samples()
    print("samples ready in", SAMPLES)

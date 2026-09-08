"""生成端到端测试用的样本文件(直接使用 openpyxl,不走 xcli,保证样本真实可控)。

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


if __name__ == "__main__":
    make_demo_xlsx()
    make_csvs()
    make_logo()
    print("samples ready in", SAMPLES)

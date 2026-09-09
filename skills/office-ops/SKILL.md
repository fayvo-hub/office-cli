---
name: office-ops
description: 通过 office-cli 命令读写/转换 Excel、Word、PDF、CSV/JSON 与 Markdown 排版文档。当用户要求生成/修改/读取 Excel 表(任意版本含 .xls)、合并拆分单元格、设置样式、插入图表图片、批量读写数据、处理 CSV/JSON 表格、把 Excel/CSV 导出为 Markdown 表格或纯文本、读写/创建 Word 文档(标题/段落/表格/图片/代码块)、操作 PDF(合并/拆分/旋转/加密/解密/水印/提取文本/表格/图片/转图片)、把 Markdown 转成 PDF/Word/HTML、文档格式互转(含 .doc/.docx、老 .xls)时触发。
---

# office-ops 技能:办公文档全格式操作(office-cli)

本技能操作 **本机已安装的 office-cli**。若 `office` 命令不可用,先按下方
「自动安装」一节装好;然后 `office info` 确认环境与可用引擎,再按需调用。
命令形态:`office <组> <子命令> -f 文件 [参数]`;成功=stdout JSON + 退出码 0;业务错误=stderr JSON `{"error":{"code","message"}}` + 退出码 3(参数错误=2)。所有输出 UTF-8 JSON,直接解析即可。

## 自动安装(首次使用/office 命令不存在时)

公开仓库: https://github.com/fayvo-hub/office-cli
安装目标:**生成 `office` 命令**(Python 的 Scripts 目录 / 或解压目录下的 office.exe)。
原则:有 Python 走 pip(功能最全、升级最方便);无 Python 下载免安装 exe;
git 不是必需(pip 可直接装 zip)。全部命令幂等,可重复执行。

### 第 1 步:探测现状

```cmd
office --version   :: 能输出版本 → 已装好,直接跳到最后一步验证
```

### 第 2 步:找 Python(任一可用即可,需 3.9+)

依次尝试:`py -3.13` / `py -3.12` / `py -3.11` / `py -3.10` / `py -3.9` / `python`,
取第一个能跑 `--version` 的解释器,记作 PY(后续安装命令都用它)。
Windows 下 `py` launcher 与 `python` 通常指向同一环境,装完即可共用。

### 第 3 步:安装(按以下顺序,命中即止)

**A. 有 Python + git(升级体验最好)**
```cmd
%PY% -m pip install --upgrade "office-cli[mdpdf,wps,images] @ git+https://github.com/fayvo-hub/office-cli.git"
```

**B. 有 Python、无 git(pip 直装 GitHub zip,推荐主力路径)**
```cmd
%PY% -m pip install --upgrade "office-cli[mdpdf,wps,images] @ https://github.com/fayvo-hub/office-cli/archive/refs/heads/master.zip"
```

**C. 有 Python 但 pip 装不上(网络/构建异常)——免安装兑底**
```cmd
:: 下载 zip 解压到固定目录(如 %LOCALAPPDATA%\office-cli),之后用 python -m office 代替 office
powershell -Command "Invoke-WebRequest -Uri https://github.com/fayvo-hub/office-cli/archive/refs/heads/master.zip -OutFile %TEMP%\office-cli.zip; Expand-Archive %TEMP%\office-cli.zip %LOCALAPPDATA%\office-cli -Force"
cd /d %LOCALAPPDATA%\office-cli\office-cli-master && python -m office info
:: 注:python -m office 与 office 命令完全等价,只是前缀长;装好 pip 后可随时补 A/B
```

**D. 完全没有 Python——下载免安装单文件版(约 120MB,来自 GitHub Release)**
```cmd
powershell -Command "Invoke-WebRequest -Uri https://github.com/fayvo-hub/office-cli/releases/latest/download/office.exe -OutFile %LOCALAPPDATA%\office-cli\office.exe"
:: 之后用完整路径调用,或把 %LOCALAPPDATA%\office-cli 加入 PATH
%LOCALAPPDATA%\office-cli\office.exe info
```
exe 版限制:md→pdf 不可用(缺 playwright,报错会提示装完整版);WPS 老格式升级/排版导出需要本机装 WPS Office。

**E. 连 Python 都没有、也不想用 exe** → 先引导装 Python 再重试:
```cmd
winget install -e --id Python.Python.3.13
:: 或官网 https://www.python.org/downloads/ 下载安装(勾选 Add to PATH)
```

### 第 4 步:验证(必做)

```cmd
office info
```
读输出 JSON:`ok` 必须为 true;`mode` = pip|exe;`missing` 期望为 `[]`(pip 全量)或
仅 `["playwright"]`(exe 版,此时不要调用 md to-pdf);`engines.wps` 为 false 时
.xls/.doc 老格式与 docx→pdf 不可用(需装 WPS Office)。
**若以上步骤均失败,直接告诉用户安装失败原因与第 E 步/手动安装指引,不要反复重试。**

## 通用约定

- 写操作默认**原地改写**原文件;结果 JSON 里 `file` 指向实际写到的文件。
- 老格式:`.xls`/`.doc` 输入会自动经 WPS 升级为 `.xlsx`/`.docx` 再操作,原文件不动;升级会出现在 `warnings`。
- 编码/分隔符:CSV 读取自动识别(utf-8-sig/utf-8/gb18030 与 `,` `;` `\t` `|`);xlsx→csv 带 BOM(Excel 可直接打开)。
- 复杂 JSON 数据一律写临时文件再 `--data-file`,避免 Shell 转义。

## 常用命令速查

### Excel(组前缀 excel)
```bash
office excel list -f 表.xlsx --preview 5         # 概览: 表名/维度/合并区 + 每表预览前 N 行(默认 3,0 关闭)
office excel read -f 表.xlsx --range A1:F20       # 行数组 JSON; --sheet 选表; --cached 公式取缓存值
office excel read -f 表.xlsx --range A1:F20 --cells   # 逐格输出(含类型/number_format); --limit N 行数上限(默认10000,0不限); --drop-empty-rows
office excel write -f 表.xlsx --cell B2 --data 42
office excel write -f 表.xlsx --cell A1 --data '[[1,2],[3,4]]'
office excel write -f 表.xlsx --data '[["总分","=SUM(B2:B5)"]]' --sheet 成绩
office excel write -f 新.xlsx --create --data '[[...]]'   # 新建(存在则报错,加 --overwrite 覆盖)
office excel sheet add -f 表.xlsx --name 新表 --index 0   # add/rename/remove/copy/active; --index 0 起(默认末尾); 最后一个表不可 remove
office excel style -f 表.xlsx --range A1:E1 --bold --fill '#FFF2CC' --align center --wrap
office excel merge -f 表.xlsx --range A1:C1 --center     # --unmerge 取消; --discard-others 丢弃区域内其他值
office excel chart add -f 表.xlsx --type bar --data A1:B6 --at F2 --title 季度   # bar|col|line|pie; --at 默认数据右下
office excel image add -f 表.xlsx --image logo.png --at D2 --scale 0.5   # --width px 与 --scale 二选一; 默认锚 A1
office excel pivot list -f 表.xlsx               # 透视表清单(名称/所在表/数据源范围); 只能操作已存在的透视表,不能新建
office excel pivot set-source -f 表.xlsx --name 表 --ref A1:F100  # 改数据源(须含表头); 默认设 refreshOnLoad, --no-refresh 关闭
```
要点:
- write 是**覆盖式**:从 `--cell` 起点(默认 A1)直接写,会**覆盖原区域内容**;想追加到末尾须先 read 算出空行再指定 `--cell`。一维数组自上而下成列,二维按矩形写。
- 值约定:null 清格;true/false 写布尔;`=` 开头的字符串按公式(如 `"=SUM(A1:A9)"`),想存为文本加 `--literal`;日期给 ISO 字符串(如 `2024-03-05T00:00:00`)。
- 合并区内写入报 `merged_cell`(只能写左上角);read 的公式格默认输出公式文本,加 `--cached` 取缓存值(openpyxl 写出的文件一般无缓存,取不到仍回退公式文本)。
- style 全部参数: `--font-name --font-size --bold --italic --underline --font-color(hex) --fill(hex|none 清除) --align --valign --wrap --border --border-color --num-format`;布尔参数可写成 `--bold` 或 `--bold true/false`。

### Word(组前缀 word)
```bash
office word read -f 报告.docx                     # 段落/表格/图片元数据; --limit N(默认500,0不限); --no-tables 跳过表格; --save-images DIR 导出图片(默认只列图片清单)
office word write -f 报告.docx --create --text '标题内容'     # 不存在则新建; 存在=末尾追加
office word write -f 报告.docx --data-file blocks.json
```
blocks JSON(顶层 `{"blocks": [...]}`,块类型缺省 p):
```json
{"blocks": [
  {"type": "h", "level": 1, "text": "一级标题"},
  {"type": "p", "text": "普通段落", "align": "center", "color": "#333333", "size": 12},
  {"type": "p", "runs": [{"text": "重点", "bold": true, "color": "E00000"}, {"text": "普通", "italic": true}]},
  {"type": "code", "text": "print(1)"},
  {"type": "quote", "text": "引用语"},
  {"type": "table", "rows": [["列A", "列B"], [1, 2]]},
  {"type": "img", "path": "logo.png", "width": 8},
  {"type": "pagebreak"}
]}
```
块说明:h=标题(level 1-9);p 可 `text` 或 `runs[]`(逐段富文本:bold/italic/code 等宽/color 6位hex/size pt,也可块级统一设);code=等宽灰底;quote=斜体+左缩进;table=二维 rows、网格边框、**首行默认加粗**(表头,加 `--no-header` 取消);img=`path`+`width`(单位厘米,默认 15);pagebreak=分页。

### PDF(组前缀 pdf)
```bash
office pdf info -f 文档.pdf                       # 页数/尺寸/加密(加密文件页数=null, 需先解密)
office pdf read -f 文档.pdf --pages 1-3 --tables  # 文本提取(可选逐页表格); --max-chars 每页字符上限(默认8000,0不限)
office pdf merge -f a.pdf b.pdf --out c.pdf       # 按给出顺序合并
office pdf split -f a.pdf --pages 1-3,5 --out part.pdf   # 抽页成新 PDF; --out-dir d/ 则每页一文件(a_p001.pdf); --pages 缺省全部,支持开区间 5-
office pdf rotate -f a.pdf --angle 90             # 90/180/270; --page 1-3,5 只转指定页(缺省全部); 默认原地, --out 另存
office pdf encrypt -f a.pdf --password 'pw' --out e.pdf   # AES-256; --owner 另设所有者密码(默认同用户密码); decrypt 同构
office pdf watermark -f a.pdf --text '内部资料' --opacity 0.15 --out w.pdf   # --size pt(默认按页高12%自动); --no-rotate 不旋转
office pdf images -f a.pdf --out-dir imgs/        # 抽内嵌图(命名 p页码-序号); 无图报 no_images
office pdf to-image -f a.pdf --out-dir png/ --dpi 150 --pages 1   # 整页转 PNG(默认 150dpi, --pages 缺省全部)
```
错误码: 已加密操作未解密文件=encrypted;解密密码错=bad_password;文件未加密却 decrypt=not_encrypted。

### Markdown 排版(md 组,样式/代码高亮/mermaid 图)
```bash
office md to-pdf -f 说明.md --out 说明.pdf        # A4、页码; --css '自定义样式串' 可追加样式; 需本机 Chrome/Edge
office md to-docx -f 说明.md --out 说明.docx
office md to-html -f 说明.md --out 说明.html
```
支持标题层级/表格/引用/图片(md 同目录相对路径)/代码块高亮/mermaid 代码块;A4 排版带页码。

### 跨格式转换(顶层 convert,方向按 --out 后缀自动判定,也可 --to 显式指定)
```bash
office convert -f 表.xlsx --out 表.csv            # xlsx→csv(带表头+BOM)
office convert -f 表.csv --out 表.xlsx            # csv→xlsx
office convert -f 表.xlsx --to json --out 表.json # json 互转同理; 所有 xlsx 源转换都支持 --sheet/--range/--cached
office convert -f 表.xlsx --out 表.md            # 导出 Markdown 表格(首行作表头)
office convert -f 表.xlsx --out 表.txt           # 导出 TSV 制表符文本(信息无损)
office convert -f 旧.xls --out 新.xlsx            # 老格式升级(WPS,原文件不动)
office convert -f 报告.docx --out 报告.pdf        # Word→PDF 走 WPS 保真排版(--engine 可选,默认 wps)
office convert -f 文档.pdf --to docx --out 文档.docx   # pdf2docx 版面还原
office convert -f 报告.docx --to md --out 报告.md  # docx→md 结构近似
office convert -f 说明.md --out 说明.pdf           # md→pdf 默认样式(精细控制用 md to-pdf);也可 --css 文件路径定制
office convert -f 说明.md --out 说明.html          # md→html 同样通 convert
```
支持矩阵: xlsx/xlsm/xls→csv|json|xlsx|md|txt(xls 自动升级);csv→xlsx|json|md|txt;json→xlsx|md|txt;doc/docx→pdf(doc 自动升级)、docx→md、doc→docx;pdf→docx;md→pdf|docx|html。CSV→xlsx 会做类型推断(数字/布尔/日期),`--no-infer` 关闭;`--delimiter`/`--encoding` 可强制;输出已存在直接覆盖,输入输出同路径报 same_file。

## 工作流建议

1. **先摸底**: `office excel list` / `office pdf info` / `office word read` 看结构再动手;改动前先 read 相关区域。
2. **小步快跑**: 大批量/多步操作拆成多条命令,每条校验返回 JSON 的 `ok` 与 `warnings`;出错看 `error.code` 与中文 `message` 自行修正(如 no_file/merged_cell/same_file)。
3. **临时数据**: 生成 JSON 到临时文件再 `--data-file`;不要 inline 长 JSON。
4. **样本练习**: `office-cli/samples/` 下有 demo.xlsx(含公式缓存版 demo_cached.xlsx)、demo_pivot.xlsx、sample.docx、sample.md、legacy.xls、legacy.doc、gbk.csv——需要真实素材验证时用它们。
5. **权限说明**: 文件被 Excel/WPS 占用时写操作报 `file_busy`,让用户先关闭程序重试。

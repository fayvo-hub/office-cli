# office-cli

**面向 AI / 自动化脚本的办公文档命令行工具**:一行命令读写 Excel(.xlsx/.xlsm/.xls)、CSV/JSON 互转、Word(.docx/.doc)、PDF 全套操作与 markdown 排版渲染(PDF/Word/HTML)。

一切输入输出均为 **UTF-8 JSON**(含错误),因此无需"解析终端文本",任何语言、任何 AI 都能稳定调用。

```bash
office excel read -f 报表.xlsx --range A1:F20 --cached
office excel write -f 表.xlsx --data '[["姓名","分数"],["张三",92]]' --cell A1
office convert -f 表.xlsx --out 表.csv --sheet 成绩   # 自动识别分隔符/编码
office word write -f 周报.docx --text '本周完成…'      # 不存在自动新建
office md to-pdf -f 说明.md --out 说明.pdf             # 保留样式 + mermaid 渲染
office pdf merge -f a.pdf b.pdf --out c.pdf
office pdf encrypt -f c.pdf --password '1234'
office pdf watermark -f c.pdf --text 内部资料
office pdf to-image -f c.pdf --out-dir 页图/ --dpi 150
```

## 安装

公开仓库: https://github.com/fayvo-hub/office-cli

**需要 Python 3.9+;git 不是必需**(pip 支持直接装 GitHub zip)。安装后命令行
可用 `office`,也可以用 `python -m office`。

### 一键安装(用户)

```bash
# 有 git(升级方便,同样命令可重复执行)
python -m pip install --upgrade "office-cli[mdpdf,wps,images] @ git+https://github.com/fayvo-hub/office-cli.git"
# 无 git: pip 直装 GitHub zip
python -m pip install --upgrade "office-cli[mdpdf,wps,images] @ https://github.com/fayvo-hub/office-cli/archive/refs/heads/master.zip"
# 免安装兑底: 下载解压后用 python -m office 调用(等价,前缀长一些)
```

### 免 Python:单文件版 office.exe

GitHub Releases 提供免安装单文件版(约 120MB):下载 `office.exe` 后直接运行。
受限项:md→pdf 需 playwright 库(单文件版不含,调用时会提示装完整版);
.xls/.doc 老格式升级与 docx→pdf 排版导出依赖本机办公引擎(**WPS Office** 或 **LibreOffice** 任一)。

### 开发安装(本仓库)

```bash
pip install -e ".[mdpdf,wps,images]"   # mdpdf=md 渲染;wps=WPS COM;images=图片
pip install -e ".[formula]"            # 可选公式引擎(PyPI formulas, 跨平台)
python scripts/build_exe.py             # 重新打包单文件 exe(需要 pip install pyinstaller)
```

核心依赖(openpyxl/python-docx/PyMuPDF/pdfplumber/pypdf/pdf2docx 含其依赖)
随包自动安装,Excel/Word/PDF 全套开箱即用。

### 环境需求

- **md to-pdf** 需要本机 Chrome 或 Edge(用系统浏览器渲染)+ playwright 库;
- **.xls / .doc / .ppt 老格式**与 Office→pdf、xlsx 公式重算需要办公引擎:
  **WPS Office**(Windows 快)或 **LibreOffice**(跨平台,Linux/macOS/容器);
  两者都没有时会在被用到的命令上报 `engine_unavailable`(纯 Python 功能不受影响);
- 不确定环境先跑 `office info`(自检依赖/引擎/运行模式 pip|exe)。

### 引擎(WPS / LibreOffice)

只有三类操作需要外部办公套件: 旧格式升级(`.xls/.doc/.rtf/.ppt` → 新格式)、
Office→PDF 导出(`docx/pptx→pdf`)、xlsx 公式重算(`rag prep --recalc`)。
它们统一走 `office/engine.py` 门面(调用方不感知后端),选择规则:

| 设置 | 行为 |
|---|---|
| 默认(auto) | Windows 有 WPS 用 WPS,否则 LibreOffice;macOS/Linux 反之 |
| `OFFICE_ENGINE=wps` | 只用 WPS(不可用时 `engine_unavailable`) |
| `OFFICE_ENGINE=lo` | 只用 LibreOffice(跨平台一致,适合服务器/容器) |
| `OFFICE_ENGINE=none` | 禁用外部引擎(纯 Python 路径,需要时明确报错) |
| `LIBREOFFICE_SOFFICE` | 手动指定 soffice 可执行文件路径 |

```bash
office info                                       # 看 engines.active(WPS/LibreOffice/none)
OFFICE_ENGINE=lo office word to-pdf -f 报告.docx   # 指定用 LibreOffice 导出
OFFICE_ENGINE=lo office rag prep -f 表.xlsx --recalc   # 用 LibreOffice 重算公式
# Linux 容器: apt install libreoffice fonts-noto-cjk(macOS: brew install --cask libreoffice)
```

版本要求: LibreOffice ≥ 7.3、WPS Office ≥ 2019(Windows)。

### 打印机隔离(Windows)

WPS / LibreOffice / Chrome 启动时都会初始化 Windows 打印子系统并读取“默认打印机”能力;
若默认打印机是 WSD/IP 网络打印机,系统就会去连它(命令变慢、托盘出现“连接打印机”),
而 office-cli 从不需要打印。所以引擎调用期间会把默认打印机临时换成本地虚拟打印机
(`Microsoft Print to PDF`),用完立即恢复原值(实测切换 19ms/恢复 4ms);默认打印机本身
就是本地端口时不做任何操作。

| 设置 | 说明 |
|---|---|
| `OFFICE_PRINTER_ISOLATE=0` | 关闭隔离(默认开启) |
| `OFFICE_VIRTUAL_PRINTER=名称` | 指定替代用的本地虚拟打印机 |
| `office info` → `engines.printer` | 看当前默认打印机/端口/是否网络机/替代目标 |

非 Windows 平台不适用(macOS/Linux 不走 winspool)。隔离只影响引擎进程启动那几秒,
不改变你的日常默认打印机。

## 通用协议

| 约定 | 说明 |
|---|---|
| 成功 | stdout 输出一个 JSON 对象 `{"ok": true, ...}`,退出码 **0** |
| 参数错误 | 退出码 **2**(argparse),用法打印到 stderr |
| 业务错误 | stderr 输出 `{"error": {"code": "...", "message": "中文原因"}}`,退出码 **3** |
| 内部错误 | 退出码 **4** |
| 文件参数 | `-f/--file`;**不含扩展名时由内容判定**(.xls/.doc 自动升级) |
| 警告 | 结果里 `"warnings": [...]`(如合并单元格丢数据、格式升级提示) |
| 列范围 | `A1:B3`、`A:C`、`1:3`、`B2` 均可 |
| 数据 | `--data` 内联 JSON 或 `--data-file` 文件(UTF-8);`null`=清空,`true/false`=布尔,`=` 开头=公式 |

### 错误码速查

`no_file` 文件不存在 · `bad_json` 数据不是 JSON · `unsupported_format` 不支持的格式/方向 ·
`no_sheet`/`dup_sheet` sheet 不存在/重名 · `merged_cell` 目标在合并区内 · `too_large` 超过上限 ·
`file_busy` 文件被占用(自动重试后仍失败) · `encrypted`/`bad_password`/`not_encrypted` PDF 密码相关 ·
`no_images` PDF 无内嵌图 · `same_file` 输入输出同文件 · `need_dep`/`need_chrome` 缺依赖 ·
`wps_unavailable`/`wps_timeout` 办公引擎不可用/超时(兼容旧码) ·
`engine_unavailable` 需要引擎但本机无 WPS/LibreOffice · `unsupported_format` 未知扩展名

## 命令总览

```
office excel list          # 工作簿信息: sheet 列表/行列数/合并区
office excel read          # 读范围/整表: rows 或 cells, 支持公式缓存值
office excel write         # 写值/公式/布尔, 追加行, 新建或原地改写
office excel sheet         # add/rename/remove/copy/active
office excel style         # 字体/对齐/填充/边框/数字格式(按范围)
office excel merge         # 合并/取消合并单元格
office excel chart         # bar/col/line/pie 图表
office excel image         # 插入/删除图片
office excel pivot         # list/set-source 透视表
office excel layout        # 版式: 列宽/行高/冻结窗格/自动筛选/隐藏行列
office excel cond-format   # 条件格式 add/clear(比较/区间/包含/重复值高亮)
office excel comment       # 批注 set/clear
office excel validate      # 数据验证(下拉) add/clear
office excel insert        # 插入行列(合并区/表/筛选自动检查)
office excel delete        # 删除整行/整列

office excel replace       # 查找替换(--regex/--in-formulas 可选)
office convert             # 跨格式互转, 见下表
office word read           # 段落/表格/图片元数据(可导出图片)
office word write          # 新建/追加段落、标题、表格、代码块、图片、分页
office word replace        # 替换/模板占位符填充(正文/表格/页眉, 跨 run 合并)
office ppt read            # 页/标题/文本层级/表格/图片/备注(可导出图片)
office ppt write           # 新建/追加幻灯片(标题/要点/表格/图片/备注)
office ppt to-pdf          # PPT → PDF(办公引擎: WPS/LibreOffice)
office pdf info            # 页数/尺寸/加密/元数据(加密文件也能查)
office pdf read            # 按页提取文本/表格(可选 --tables)
office pdf merge           # 合并多个 PDF
office pdf split           # 按页拆分(单文件或 --out-dir 每页一文件)
office pdf rotate          # 旋转 90/180/270
office pdf encrypt         # 设置密码(AES-256)  office pdf decrypt
office pdf watermark       # 对角线文字水印(支持中文)
office pdf images          # 抽取内嵌图片  office pdf to-image  # 整页转 PNG
office pdf from-images     # 图片(组)合成 PDF(jpg/png/webp 混排)
office pdf footer          # 页脚/页码(中英文本, 可指定页)
office pdf search          # 全文查找(页/命中数/上下文)
office convert -f 页.html --to pdf  # HTML → PDF(Chrome 渲染)
office md to-pdf           # Markdown → PDF(样式/代码高亮/mermaid)
office md to-docx          # Markdown → Word
office md to-html          # Markdown → 自包含 HTML
office rag prep            # RAG 清洗: Office/PDF → 语义重建 md+JSON(见下)
office info                # 环境自检: 依赖/引擎/样本路径
```

### excel 组

```bash
office excel list -f 表.xlsx
office excel read -f 表.xlsx --sheet 成绩 --range A1:F20 --cached   # 公式=缓存值
office excel read -f 表.xlsx --cells C3 --limit 500
office excel write -f 表.xlsx --cell A1 --data '[[1,2],[3,4]]'       # 原地写入
office excel write -f 表.xlsx --data '[["总分","=SUM(B2:B5)"]]'      # 公式
office excel write -f 新表.xlsx --data '[["a","b"]]' --create        # 新建
office excel sheet copy -f 表.xlsx --name 成绩 --new-name 成绩备份   # 复制 sheet
office excel style -f 表.xlsx --range A1:E1 --bold --fill '#FFF2CC' --align center
office excel merge -f 表.xlsx --range A1:C1 --center
office excel layout -f 表.xlsx --col-width 'A=18,C:E=22'     # 列宽(字符数); 行高同理 --row-height '1=28,3:5=20'
office excel layout -f 表.xlsx --freeze B2                  # 冻结窗格(A2=冻首行, B1=冻A列); --unfreeze 取消

office excel layout -f 表.xlsx --filter A1:F100             # 表头自动筛选; --hide-cols 'C:F'/--show-rows 等显隐

office excel cond-format add -f 表.xlsx --range C2:C20 --op between --value '50,100'  # --op: gt/lt/gte/lte/eq/neq/between/not-between/contains/duplicates

office excel cond-format add -f 表.xlsx --range D2:D20 --op duplicates   # 重复值高亮; clear 用 --range 或 --all

office excel comment set -f 表.xlsx --cell B2 --text '备注' --author 张三   # set 新建/覆盖; clear 删除

office excel validate add -f 表.xlsx --range D2:D20 --list '甲,乙,丙'      # 选项含逗号用 --source '选项表!A1:A5'

office excel validate clear -f 表.xlsx --range D2:D20        # 移除与区域相交的下拉验证

office excel insert -f 表.xlsx --rows 3 --count 2            # 第 3 行前插 2 行; --cols C 插列

office excel delete -f 表.xlsx --rows 3-5                    # 删第 3~5 行; --cols B:D 删列

office excel replace -f 表.xlsx --find 旧值 --replace 新值   # 可加 --regex/--match-case/--in-formulas/--range
office excel chart add -f 表.xlsx --type bar --data A1:B6 --at F2 --title 季度
office excel image add -f 表.xlsx --image logo.png --at D2 --scale 0.5
office excel pivot set-source -f 表.xlsx --name 数据透视表1 --ref A1:F100
```

读表输出 `mode: "rows"` 时逐行数组(跳过空行),`"cells"` 逐格含 `row/col/ref/value/type`。
写表默认"追加到首个空行";指定 `--cell` 则从该格开始。`--create` 建新文件(已存在时报错,加 `--overwrite` 才覆盖);不带 `--create` 时对已存在文件原地改写。合并单元格区域内的写入会报 `merged_cell`(提示用左上角)。

### word 组

```bash
office word write -f 报告.docx --create --text '第一行'
office word write -f 报告.docx --data-file blocks.json   # 追加
```

`blocks.json`:

```json
{"blocks": [
  {"type": "h", "level": 1, "text": "季度总结"},
  {"type": "p", "text": "正文", "bold": false, "color": "#333333"},
  {"type": "p", "runs": [{"text": "重点", "bold": true}, {"text": "其余"}]},
  {"type": "code", "text": "print(1)"},
  {"type": "quote", "text": "引用"},
  {"type": "table", "rows": [["品名", "数量"], ["苹果", 3]]},
  {"type": "img", "path": "logo.png", "width_cm": 6},
  {"type": "pagebreak"}
]}
```

`word read` 输出段落(`text/style/level`)、表格(带表头样式名)、图片元数据;`--save-images DIR` 导出正文图片。
`word replace` 两种用法: `--find 旧词 --replace 新词`(单对)或 `--data-file repl.json` 批量模板替换(JSON 对象 `{"{{金额}}": "12800 元"}`,值 null=删空);加 `--ignore-case` 忽略大小写。命中正文/表格/页眉页脚,优先 run 级精确替换保留样式,跨 run 文本自动合并重建(样式取首 run);`summary` 返回替换数/重建段数。
`.doc` 文件会先经办公引擎(WPS/LibreOffice)升级为 `.docx` 再读/写,原文件不动。

### ppt 组

```bash
office ppt read -f 演示.pptx                      # 每页 title/texts(含层级)/tables/pictures/notes
office ppt read -f 演示.pptx --slide 2 --save-images imgs/
office ppt write -f 演示.pptx --create --data-file slides.json
office ppt to-pdf -f 演示.pptx --out 演示.pdf    # 走办公引擎
```

`slides.json`:`{"slides": [{"layout": "title", "title": "季度汇报"}, {"layout": "title-content", "title": "进展", "bullets": ["一", {"text": "子项", "level": 1}], "notes": "备注"}]}`;
`layout`: `title`(仅标题) / `title-content`(标题+要点,自动选版式) / `blank`(自由摆放,支持 `texts`/`table`/`picture` 坐标尺寸,单位英寸)。
无 `--create` 时在现有演示文稿末尾追加。`.ppt` 老格式经办公引擎升级后读/追加,原文件不动。

### pdf 组

```bash
office pdf info -f 文档.pdf                      # 加密文件返回 encrypted:true + 页尺寸
office pdf read -f 文档.pdf --pages 1-3 --tables
office pdf split -f 文档.pdf --pages 1,3,5-8 --out-dir out/
office pdf rotate -f 文档.pdf --angle 90        # 原地旋转
office pdf watermark -f 文档.pdf --text '机密' --opacity 0.15
office pdf images -f 文档.pdf --out-dir imgs/
office pdf to-image -f 文档.pdf --dpi 200 --pages 1
office pdf footer -f 文档.pdf --text '第 {page} 页 / 共 {pages} 页'   # 原地加页脚; {page}{pages} 占位

office pdf footer -f 文档.pdf --text '机密' --pages 2-4 --color 990000   # 指定页/字号 --size/下边距 --margin

office pdf search -f 文档.pdf --find '季度'                       # 命中页码+上下文片段(无命中 total=0)

office pdf from-images -f 扫描件.pdf --out 扫描.pdf a.jpg b.png   # 图片为位置参数, 每张一页适配 A4
```

加密文件先 `decrypt` 后才能 read/merge/split/rotate/watermark。

### rag 组(RAG 数据清洗)

```bash
# 单文件: 同目录生成 <名>.md + <名>.json
office rag prep -f 报表.xlsx
office rag prep -f 报表.xls --out-dir clean/        # 老格式自动升级
# 批量目录: 另写 qa.json 质检汇总(坏文件/公式未解析上库前暴露)
office rag prep -f 文档目录/ --out-dir clean/ --recursive
# xlsx 表头判定手动兜底: auto 自动(默认) / N 固定行数 / 0 无表头
office rag prep -f 表.xlsx --header-rows 2
# 公式保真兜底: 数组公式等内置求值器算不出的, 用办公引擎重算后解析(需 WPS 或 LibreOffice)
# 显式指定后端: --engine lo(LibreOffice)/wps(WPS)/formulas(PyPI 包)/builtin(默认内置)
office rag prep -f 表.xlsx --recalc
# 或改用跨平台 PyPI formulas 引擎(需 pip install "office-cli[formula]")
office rag prep -f 表.xlsx --engine formulas
```

支持 .xlsx/.xlsm/.xls/.csv/.docx/.doc/.rtf/.pptx/.ppt/.pdf/.txt(目录批量;.xls/.doc/.rtf/.ppt 老格式自动经办公引擎升级)。
与普通转换的关键差异(xlsx 语义重建):

- **多 sheet 全量导出**(隐藏 sheet/空表跳过并说明),不做静默截断
- **合并单元格展开**: 纵向/横向/二维块均按 Excel 显示语义整块铺开(数据行带完整归属维度,
  二维块坐标记 warning);顶部整行合并标题识别为标题文本、短合并标题(<60% 列宽)抽为副标题,不混入表头
- **一个 sheet 多张表**(空行分隔)+ 备注/说明段: 输出 blocks 数组(table/notes),表尾“注:…”全宽行不误当标题
- **多行表头按列合成层级标签**(如 `2023年 / 上半年`),纵向合并导致的重复自动去重
- **内置公式求值器(90+ 函数, 纯标准库零依赖)**: 无缓存公式直接计算 ——
  统计 SUM/AVERAGE/COUNT(A)/MAX/MIN/MEDIAN/LARGE/SMALL/RANK/STDEV/VAR/PERCENTILE/SUMPRODUCT,
  条件聚合 SUMIF(S)/COUNTIF(S)/AVERAGEIF(S)/MAXIFS/MINIFS(通配符与比较运算符),
  查找 VLOOKUP/HLOOKUP/XLOOKUP/LOOKUP/INDEX/MATCH/OFFSET/ROW/COLUMN/ROWS/COLUMNS,
  逻辑 IF/IFS/AND/OR/NOT/CHOOSE/SWITCH/IFERROR/IFNA/IS* 系列,
  文本 LEFT/RIGHT/MID/LEN/TRIM/大小写/PROPER/SUBSTITUTE/REPLACE/FIND/SEARCH/TEXT/TEXTJOIN/CONCAT/EXACT/VALUE,
  日期 TODAY/NOW/DATE/日期拆分/EOMONTH/EDATE/DATEDIF/NETWORKDAYS/WORKDAY/DAYS/WEEKDAY/DATEVALUE,
  数学与财务 SQRT/POWER/LOG/INT/MOD/ROUND 系列/CEILING/FLOOR/TRUNC/SUMSQ/PRODUCT/PMT/FV/PV/NPV/IRR;
  算不了回退缓存值, 再无缓存保留原文并记入 `formula_unresolved`
- **公式保真三级**: ① 文件自带缓存值(最快) → ② 内置求值器(零依赖, 常见公式全覆盖) →
  ③ 兜底引擎二选一: `--engine formulas`(PyPI `formulas` 包, 跨平台、数组表达式也能算,
  需 `pip install "office-cli[formula]"`;每个文件加载 1~10s)或 `--engine wps`/`--recalc`
  (本机办公引擎重算, 保真度最高;WPS 与 LibreOffice 均可, 跨平台用 LibreOffice)
- **数值按 number_format 渲染**: 百分比(13%→"13%")、日期 ISO(2024-01-12),与 Excel 显示一致
- CSV 自动识别编码(utf-8-sig/utf-8/gb18030)与分隔符;txt 输出为段落 notes
- **PPTX 结构重建**(纯 python-pptx,不需引擎): 每页一个 sheet(「第 N 页 标题」),
  标题/要点保留缩进层级,页内表格→table 块,演讲者备注单独列出;.ppt 旧格式自动升级
- **RTF/DOC 同 docx 管道**: 经办公引擎升级后走标题层级/表格保留逻辑,原文件不改动
- **去噪**: PDF 页脚/页码整行去除(含 ⻚ 兼容字符)、表格内容不再与正文重复出现;docx 标题层级保留、
  表格合并单元格展开不重复(缺 `<w:tblGrid>` 的脏表按各行 `w:tc`/`w:gridSpan` 推断列数,不整表丢弃);
  pptx 页标题与正文首行重复时只保留标题
- 目录模式 `qa.json` 汇总每个文件的行数/表数/公式未解析/告警/大小/耗时;单文件失败不中断批量

### html → pdf

```bash
office convert -f 页面.html --to pdf --out 页面.pdf   # Chrome 渲染(同 md to-pdf 引擎)
```

### md 组(排版渲染)

```bash
office md to-pdf -f README.md --out README.pdf        # A4, 中文字体, 页码
office md to-docx -f README.md --out README.docx
office md to-html -f README.md --out README.html
```

支持:标题层级、表格、代码高亮、引用、图片(相对 md 目录)、**mermaid 图**(graph TD 等,依赖本地 Chrome 渲染);A4 页面带页码。

### convert 跨格式矩阵

| 从 \ 到 | 说明 |
|---|---|
| xlsx/csv/json ↔ xlsx/csv/json | 纯 Python,自动识别编码(utf-8-sig/utf-8/gb18030)与分隔符 |
| xlsx/csv/json → md/txt | 导出 Markdown 表格 / TSV 纯文本(支持 --sheet/--range/--cached) |
| .xls → xlsx/csv/json 等 | 经办公引擎升级后转换(需装 WPS 或 LibreOffice) |
| docx/doc → docx、docx/doc → pdf | 办公引擎保真排版(Word 级质量;WPS/LibreOffice) |
| docx → md | 结构近似转换:标题/表格/代码/粗斜体 |
| pdf → docx | pdf2docx 版面还原(可编辑) |
| md → pdf/docx/html | 见 md 组 |

```bash
office convert -f 表.xlsx --out 表.csv                  # xlsx→csv(带表头)
office convert -f 表.xlsx --out 表.md                   # 导出 Markdown 表格(首行作表头)
office convert -f 表.xlsx --out 表.txt                  # 导出 TSV 文本(制表符分隔,可无损读回)
office convert -f 表.xlsx --to json --out 表.json       # 显式指定目标类型
office convert -f 说明.md --to pdf --out 说明.pdf
office convert -f 旧报告.doc --to pdf --out 旧报告.pdf  # 走办公引擎
office convert -f 文档.pdf --to docx --out 文档.docx    # pdf2docx
```

## AI / skill 集成建议

1. 让模型先 `office info` 确认环境,再 `office excel read`/`pdf info` 摸底;
2. 写操作一律先读后写;大批量改动拆小步,每步校验返回 JSON;
3. 老格式文件(扩展名 .xls/.doc)操作时留意返回的 `warnings`(升级提示);
4. 中文字符串作为参数直接传(UTF-8),无需转义;JSON 用 `--data-file` 避免命令行转义地狱;
5. 不确定参数就 `office excel write -h`(每个子命令都有中文帮助)。

## 开发与测试

```bash
python scripts/make_sample.py     # 生成 samples/: demo.xlsx、sample.docx/md、legacy.xls/.doc、logo.png
python tests/run_tests.py         # 端到端黑盒测试(调真实 CLI), 当前 175 passed
```

测试样本也可作为 skill 的练习材料:`samples/` 下 `demo.xlsx`(公式缓存值版 `demo_cached.xlsx`)、
`demo_pivot.xlsx`、`sample.docx`(标题+表格+图片)、`sample.md`(表格/代码/mermaid/引用)、
`legacy.xls`、`legacy.doc`(老格式,测自动升级)。

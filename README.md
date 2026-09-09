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
.xls/.doc 老格式升级与 docx→pdf 排版导出仍依赖本机 **WPS Office**。

### 开发安装(本仓库)

```bash
pip install -e ".[mdpdf,wps,images]"   # mdpdf=md 渲染;wps=WPS COM;images=图片
python scripts/build_exe.py             # 重新打包单文件 exe(需要 pip install pyinstaller)
```

核心依赖(openpyxl/python-docx/PyMuPDF/pdfplumber/pypdf/pdf2docx 含其依赖)
随包自动安装,Excel/Word/PDF 全套开箱即用。

### 环境需求

- **md to-pdf** 需要本机 Chrome 或 Edge(用系统浏览器渲染)+ playwright 库;
- **.xls / .doc 老格式**与 docx→pdf 需要本机 **WPS Office**(自动升级为 .xlsx/.docx 后操作,原文件不动);
- 不确定环境先跑 `office info`(自检依赖/引擎/运行模式 pip|exe)。

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
`wps_unavailable`/`wps_timeout` WPS 引擎不可用/超时 · `unsupported_format` 未知扩展名

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
office excel layout        # 布局: 行列尺寸/显隐/冻结窗格/打印区域
office excel cond-format   # 条件格式(色阶/数据条/图标/规则)
office excel comment       # 批注增改删
office excel validate      # 数据验证(下拉)增/查/清
office excel insert        # 插入行列(合并区/表/筛选自动检查)
office excel delete        # 删除行列(内容/整行/整列)
office excel replace       # 查找替换(值/公式/批注)
office convert             # 跨格式互转, 见下表
office word read           # 段落/表格/图片元数据(可导出图片)
office word write          # 新建/追加段落、标题、表格、代码块、图片、分页
office word replace        # 查找替换(正文/表格/页眉, 跨 run 合并)
office ppt read            # 页/标题/文本层级/表格/图片/备注(可导出图片)
office ppt write           # 新建/追加幻灯片(标题/要点/表格/图片/备注)
office ppt to-pdf          # PPT → PDF(WPS 引擎)
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
office excel layout -f 表.xlsx --range A:B --col-width 20     # 列宽英寸
office excel layout -f 表.xlsx --range 1:1 --row-height 30    # 行高磅
office excel cond-format -f 表.xlsx --range C2:C20 --type color-scale
office excel comment add -f 表.xlsx --cell B2 --text '备注'
office excel validate add -f 表.xlsx --range D2:D20 --list '甲,乙,丙'
office excel insert -f 表.xlsx --axis rows --at 3 --count 2
office excel delete -f 表.xlsx --axis rows --at 3 --count 2   # 连行列内容
office excel replace -f 表.xlsx --find 旧值 --replace 新值    # 值/公式/批注
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
`word replace` 用 `--find/--replace`(JSON 文件传字符串数组做多对),支持加粗/斜体等 run 级精确替换;跨 run 文本会自动合并段落重排;命中表格与页眉页脚;`.doc` 同样先升级再操作。
`.doc` 文件会先经 WPS 升级为 `.docx` 再读/写,原文件不动。

### ppt 组

```bash
office ppt read -f 演示.pptx                      # 每页 title/texts(含层级)/tables/pictures/notes
office ppt read -f 演示.pptx --slide 2 --save-images imgs/
office ppt write -f 演示.pptx --create --data-file slides.json
office ppt to-pdf -f 演示.pptx --out 演示.pdf    # 走 WPS 引擎
```

`slides.json`:`{"slides": [{"layout": "title", "title": "季度汇报"}, {"layout": "title-content", "title": "进展", "bullets": ["一", {"text": "子项", "level": 1}], "notes": "备注"}]}`;
`layout`: `title`(仅标题) / `title-content`(标题+要点,自动选版式) / `blank`(自由摆放,支持 `texts`/`table`/`picture` 坐标尺寸,单位英寸)。
无 `--create` 时在现有演示文稿末尾追加。`.ppt` 老格式经 WPS 升级后读/追加,原文件不动。

### pdf 组

```bash
office pdf info -f 文档.pdf                      # 加密文件返回 encrypted:true + 页尺寸
office pdf read -f 文档.pdf --pages 1-3 --tables
office pdf split -f 文档.pdf --pages 1,3,5-8 --out-dir out/
office pdf rotate -f 文档.pdf --angle 90        # 原地旋转
office pdf watermark -f 文档.pdf --text '机密' --opacity 0.15
office pdf images -f 文档.pdf --out-dir imgs/
office pdf to-image -f 文档.pdf --dpi 200 --pages 1
office pdf footer -f 文档.pdf --text '第 {n} 页 / 共 {N} 页'   # {n}{N} 占位
office pdf footer -f 文档.pdf --text '机密' --pages 2-4 --top      # 指定页/顶部
office pdf search -f 文档.pdf --text '季度'                       # 命中的页与上下文
office pdf from-images -f 扫描件.pdf --images a.jpg b.png --dpi 150
```

加密文件先 `decrypt` 后才能 read/merge/split/rotate/watermark。

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
| .xls → xlsx/csv/json 等 | WPS 升级后转换(WPS Office 需已安装) |
| docx/doc → docx、docx/doc → pdf | WPS 引擎保真排版(Word 级质量) |
| docx → md | 结构近似转换:标题/表格/代码/粗斜体 |
| pdf → docx | pdf2docx 版面还原(可编辑) |
| md → pdf/docx/html | 见 md 组 |

```bash
office convert -f 表.xlsx --out 表.csv                  # xlsx→csv(带表头)
office convert -f 表.xlsx --out 表.md                   # 导出 Markdown 表格(首行作表头)
office convert -f 表.xlsx --out 表.txt                  # 导出 TSV 文本(制表符分隔,可无损读回)
office convert -f 表.xlsx --to json --out 表.json       # 显式指定目标类型
office convert -f 说明.md --to pdf --out 说明.pdf
office convert -f 旧报告.doc --to pdf --out 旧报告.pdf  # 走 WPS
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

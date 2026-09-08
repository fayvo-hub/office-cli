---
name: office-ops
description: 通过 office-cli 命令读写/转换 Excel、Word、PDF、CSV/JSON 与 Markdown 排版文档。当用户要求生成/修改/读取 Excel 表(任意版本含 .xls)、合并拆分单元格、设置样式、插入图表图片、批量读写数据、处理 CSV/JSON 表格、读写 Word 文档、操作 PDF(合并/拆分/旋转/加密/解密/水印/提取文本/表格/图片/转图片)、把 Markdown 转成 PDF/Word/HTML、文档格式互转(含 .doc/.docx、老 .xls)时触发。
---

# office-ops 技能:办公文档全格式操作(office-cli)

本技能操作 **本机已安装的 office-cli**。先 `office info` 确认环境与可用引擎,再按需调用。
命令形态:`office <组> <子命令> -f 文件 [参数]`;成功=stdout JSON + 退出码 0;业务错误=stderr JSON `{"error":{"code","message"}}` + 退出码 3(参数错误=2)。所有输出 UTF-8 JSON,直接解析即可。

## 通用约定

- 写操作默认**原地改写**原文件;结果 JSON 里 `file` 指向实际写到的文件。
- 老格式:`.xls`/`.doc` 输入会自动经 WPS 升级为 `.xlsx`/`.docx` 再操作,原文件不动;升级会出现在 `warnings`。
- 编码/分隔符:CSV 读取自动识别(utf-8-sig/utf-8/gb18030 与 `,` `;` `\t` `|`);xlsx→csv 带 BOM(Excel 可直接打开)。
- 复杂 JSON 数据一律写临时文件再 `--data-file`,避免 Shell 转义。

## 常用命令速查

### Excel(组前缀 excel)
```bash
office excel list -f 表.xlsx                      # 概览: 表名/维度/合并区
office excel read -f 表.xlsx --range A1:F20       # --cached 公式取缓存值; --sheet 选表
office excel write -f 表.xlsx --cell A1 --data '[[1,2],[3,4]]'
office excel write -f 表.xlsx --data '[["总分","=SUM(B2:B5)"]]' --sheet 成绩
office excel write -f 新.xlsx --create --data '[[...]]'   # 新建(存在则报错,加 --overwrite 覆盖)
office excel sheet add|rename|remove|copy|active ...
office excel style -f 表.xlsx --range A1:E1 --bold --fill '#FFF2CC' --align center --wrap
office excel merge -f 表.xlsx --range A1:C1 --center     # --unmerge 取消
office excel chart add -f 表.xlsx --type bar --data A1:B6 --at F2 --title 季度   # bar|col|line|pie
office excel image add -f 表.xlsx --image logo.png --at D2 --scale 0.5   # 或 --width 300(px)
office excel pivot list -f 表.xlsx                 # 透视表清单; set-source --name 表 --ref A1:F100
```
要点:默认写首个空行;合并区内写入报 `merged_cell`(只能写左上角);公式格默认读出公式文本(想拿结果加 `--cached`,但 openpyxl 写的文件无缓存);布尔写 true/false,清格写 null。

### Word(组前缀 word)
```bash
office word read -f 报告.docx                     # 段落/表格/图片元数据; --save-images dir 导出图片
office word write -f 报告.docx --create --text '标题内容'     # 不存在则新建; 存在=追加
office word write -f 报告.docx --data-file blocks.json
```
blocks JSON: `{"blocks":[{"type":"h","level":1,"text":"…"},{"type":"p","text":"…","bold":true,"color":"#333333"},{"type":"p","runs":[{"text":"重点","bold":true}]},{"type":"code","text":"…"},{"type":"quote","text":"…"},{"type":"table","rows":[[…]]},{"type":"img","path":"logo.png","width_cm":6},{"type":"pagebreak"}]}`。表格首行默认加粗(表头),`--no-header` 取消。

### PDF(组前缀 pdf)
```bash
office pdf info -f 文档.pdf                       # 页数/尺寸/加密(加密文件页数=null, 需先解密)
office pdf read -f 文档.pdf --pages 1-3 --tables  # 文本提取(可选表格)
office pdf merge -f a.pdf b.pdf --out c.pdf       # 合并
office pdf split -f a.pdf --pages 1-3,5 --out-dir d/   # 或 --out 单文件; 缺省全部页
office pdf rotate -f a.pdf --angle 90             # 90/180/270, 默认原地
office pdf encrypt -f a.pdf --password 'pw' --out e.pdf   # AES-256; decrypt 同构
office pdf watermark -f a.pdf --text '内部资料' --opacity 0.15 --out w.pdf
office pdf images -f a.pdf --out-dir imgs/        # 抽内嵌图; 无图报 no_images
office pdf to-image -f a.pdf --out-dir png/ --dpi 150 --pages 1   # 整页转 PNG
```
错误码: 已加密操作未解密文件=encrypted;解密密码错=bad_password;文件未加密却 decrypt=not_encrypted。

### Markdown 排版(md 组,样式/代码高亮/mermaid 图)
```bash
office md to-pdf -f 说明.md --out 说明.pdf        # A4、页码; 需本机 Chrome/Edge
office md to-docx -f 说明.md --out 说明.docx
office md to-html -f 说明.md --out 说明.html
office convert -f 说明.md --to pdf --out 说明.pdf # 等价
```
支持标题层级/表格/引用/图片(md 同目录相对路径)/代码块高亮/mermaid 代码块;A4 排版带页码。

### 跨格式转换(顶层 convert,方向自动判定)
```bash
office convert -f 表.xlsx --out 表.csv            # xlsx→csv(带表头+BOM)
office convert -f 表.csv --out 表.xlsx            # csv→xlsx
office convert -f 表.xlsx --to json --out 表.json # json 互转同理; xlsx→csv 支持 --sheet/--range
office convert -f 旧.xls --out 新.xlsx            # 老格式升级(WPS)
office convert -f 报告.docx --out 报告.pdf        # Word→PDF 走 WPS 保真排版
office convert -f 文档.pdf --to docx --out 文档.docx   # pdf2docx 版面还原
office convert -f 报告.docx --to md --out 报告.md  # docx→md 结构近似
office convert -f 说明.md --to pdf --out 说明.pdf
```
CSV→xlsx 会做类型推断(数字/布尔/日期),`--no-infer` 关闭;`--delimiter`/`--encoding` 可强制。

## 工作流建议

1. **先摸底**: `office excel list` / `office pdf info` / `office word read` 看结构再动手;改动前先 read 相关区域。
2. **小步快跑**: 大批量/多步操作拆成多条命令,每条校验返回 JSON 的 `ok` 与 `warnings`;出错看 `error.code` 与中文 `message` 自行修正(如 no_file/merged_cell/same_file)。
3. **临时数据**: 生成 JSON 到临时文件再 `--data-file`;不要 inline 长 JSON。
4. **样本练习**: `office-cli/samples/` 下有 demo.xlsx(含公式缓存版 demo_cached.xlsx)、demo_pivot.xlsx、sample.docx、sample.md、legacy.xls、legacy.doc、gbk.csv——需要真实素材验证时用它们。
5. **权限说明**: 文件被 Excel/WPS 占用时写操作报 `file_busy`,让用户先关闭程序重试。

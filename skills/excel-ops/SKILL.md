---
name: excel-ops
description: 操作 Excel 工作簿(.xlsx/.xlsm):读取单元格与区域、写入数据与公式、管理工作表、设置样式、合并单元格、格式转换(csv/json)、插入图表图片、调整透视表数据源。当用户要求查看/修改/分析 Excel 或 xlsx 文件、把表格数据转成 csv/json 时使用。所有操作走 JSON 协议,适合大表与批量修改。
---

# Excel 操作(xcli)

基于 `excel-xcli`(openpyxl CLI 封装)。**永远输出 JSON,错误也是 JSON,绝不自由发挥。**

## 运行方式

```bash
cd C:/Users/Administrator/AppData/Roaming/pi-desktop/chat-workspace/excel-cli
python -m xcli <命令> [参数]
```

- 成功:stdout 输出 JSON,退出码 0
- 失败:stderr 输出 `{"error": {"code": "...", "message": "中文可操作提示"}}`,退出码非 0
  → **先读 message 修正参数再重试,不要硬猜**

## 标准操作流程

1. **先侦察再动手**(必做,禁止凭空猜表名):
   `list -f 文件.xlsx` → 拿到全部表名、数据范围、前 3 行预览
2. **读取数据**:`read -f 文件.xlsx --sheet 表名`(全表)或加 `--range A1:F13` 限定
3. **修改数据/结构**,每步校验返回的 `ok` 与警告
4. 需要计算值而非公式文本时加 `--cached`(无缓存会得到 null,属正常)

## 命令速查

| 需求 | 命令 |
|---|---|
| 看结构/表名/预览 | `list -f a.xlsx` |
| 读数据(JSON) | `read -f a.xlsx --sheet 销售 --range A1:F13` |
| 逐格读(类型/格式/公式) | `read -f a.xlsx --range D2 --cells --cached` |
| 写数据/公式 | `write -f a.xlsx --cell A1 --data-file 数据.json` |
| 新建文件 | `write -f 新.xlsx --cell A1 --data 1 --create` |
| 增/删/改名/复制表 | `sheet add/remove/rename/copy -f a.xlsx --name ...` |
| 表头样式 | `style -f a.xlsx --range A1:F1 --bold --fill FFC000 --align center` |
| 数字格式 | `style -f a.xlsx --range C2:C9 --num-format '0.00%'` |
| 合并/拆分 | `merge -f a.xlsx --range A1:B2 [--center] [--unmerge]` |
| xlsx→csv/json | `convert -f a.xlsx --out a.csv` |
| csv/gbk→xlsx | `convert -f a.csv --out a.xlsx` |
| 图表 | `chart add -f a.xlsx --type col --data A1:B13 --title 标题` |
| 图片 | `image add -f a.xlsx --image p.png --at B2` |
| 透视表改数据源 | `pivot set-source -f a.xlsx --name 透视表1 --ref A1:D500` |

详细说明:`python -m xcli <命令> --help`(每个命令的 help 含 JSON 结构与约定)。

## 数据约定(必须遵守)

- **read 输出**:`rows` 二维数组与表头对齐;日期是 ISO 字符串;公式格默认输出
  公式文本;`merged_cells` 列出合并区(数据只在左上角格,展示时按需填充)
- **write 输入**(--data 或 --data-file,UTF-8 JSON):
  - 二维数组 `[[...]]` → 按行列写;一维 `[1,2]` → 写一列;标量 → 写单格
  - `null` → 清空格;`true/false` → 布尔;数字 → 数字
  - 字符串以 `=` 开头 → 按**公式**写入;要写字面量加 `--literal`
- **中文/路径**:直接传路径与表名即可,内部 UTF-8 处理

## 铁律(避免踩坑)

1. 不知道表名/结构 → 先 `list`,禁止猜测
2. 覆盖性修改前,先用 `read` 把原数据取出来(或提示用户备份)
3. openpyxl **不会重算公式**:写入公式后的计算结果需 Excel 打开才刷新,
   不要把无缓存文件的 `--cached` 结果当公式真值
4. 数据透视表**只能改已有的**,不能创建;报 `no_pivot` 时改用 Excel 模板思路
5. `read` 大表默认截断 1 万行并警告;确认要全量再加 `--limit 0`
6. `merge` 会清除区域内非左上角的值(openpyxl 行为),默认会警告
7. 文件被 Excel/WPS 打开时会写不进(退出码 3 `file_busy`),提示用户先关闭
8. 不支持 `.xls` 旧格式;不支持图表/透视的**创建**(Excel 能力边界)

## 典型组合示例

**把 CSV 数据填进已有模板并美化**:
```
convert -f 数据.csv --out 数据.xlsx
style -f 数据.xlsx --range A1:F1 --bold --fill 4472C4 --font-color FFFFFF --align center
style -f 数据.xlsx --range A2:F100 --border thin
merge -f 数据.xlsx --range A1:F1 ...   # 视需求
read -f 数据.xlsx   # 校验结果
```

**给 AI 的自我检查**:每个修改命令返回后,对照 JSON 的 `ok`/`warnings` 判断是否
符合预期;需要确认视觉效果时提示用户用 Excel 打开查看。

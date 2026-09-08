# excel-xcli — 面向 AI 调用的 Excel 操作 CLI

Python + openpyxl 实现。专为 **AI skill 场景**设计:所有命令输出固定结构的 JSON,
失败时向 stderr 输出 `{"error": {"code", "message"}}`(中文、可操作)并返回非零退出码。
无交互、无弹窗,可被 AI 安全地循环调用。

## 安装

```bash
# 依赖:Python 3.9+ 与 openpyxl 3.1+(本机已有可跳过)
pip install -r requirements.txt        # 或 pip install -e . 注册 xcli 命令

# 可选:插入图片需要 Pillow
pip install Pillow
```

无需安装也可直接运行:`python -m xcli ...`(在项目根目录)。

## 命令一览

| 命令 | 作用 | 示例 |
|---|---|---|
| `list` | 列出全部工作表结构(含前 N 行预览) | `xcli list -f demo.xlsx` |
| `read` | 读取单元格/区域/整表 → JSON | `xcli read -f demo.xlsx --sheet 销售 --range A1:F13` |
| `write` | 写入数据/公式(JSON 入参) | `xcli write -f a.xlsx --cell A1 --data-file v.json` |
| `sheet` | 增/删/改名/复制表、设默认表 | `xcli sheet add -f a.xlsx --name 新表` |
| `style` | 字体/填充/对齐/边框/数字格式 | `xcli style -f a.xlsx --range A1:F1 --bold --fill FFC000` |
| `merge` | 合并/取消合并单元格 | `xcli merge -f a.xlsx --range A1:B2 --center` |
| `convert` | xlsx ↔ csv/json 数据搬运 | `xcli convert -f a.xlsx --out a.csv --cached` |
| `chart` | 插入柱状/折线/饼图 | `xcli chart add -f a.xlsx --type col --data A1:C8` |
| `image` | 插入图片(需 Pillow) | `xcli image add -f a.xlsx --image logo.png --at B2` |
| `pivot` | 透视表:列出/改数据源范围 | `xcli pivot set-source -f a.xlsx --name PT1 --ref A1:D500` |

每个命令的详细说明见 `xcli <命令> --help`(含 JSON 结构说明与约定)。

## 给 AI 的约定(读这一节就够)

1. **操作前先侦察**:先 `list -f 文件`(看有几个表、表名、结构、预览行),再决定
   read/write 的目标表与区域,避免猜错表名。
2. **公式格**:默认读出的是公式文本(如 `=B2*C2`);要计算值加 `--cached`
   (依赖文件保存时的缓存;openpyxl/程序生成的文件无缓存,读到 null 属正常)。
3. **日期**:统一输出 ISO 字符串(`2024-03-05T10:00:00` / `2024-03-05`)。
4. **合并单元格**:数据只在左上角格,其余为 null;`read` 会返回 `merged_cells`
   列表,需要铺开显示时按它填充。
5. **写入用文件传数据**:`--data-file` 传 UTF-8 JSON(二维数组按行列写、一维写一列、
   标量写单格),避免 shell 转义踩坑。`null` 清空格;`=` 开头字符串按公式写
   (`--literal` 可强作文本)。
6. **错误即反馈**:任何失败都有 `error.code` + 中文 `message`,把 message 读给用户
   或据此修正参数重试,不要臆测。
7. **透视表无法凭空创建**:只能改已有透视表的数据源。要"新透视表"请用 Excel 建好
   模板(哪怕空模板),再 set-source 套用。
8. 大文件/大区域:read 默认截断 1 万行(带警告),需要全量加 `--limit 0`;
   样式区域上限 5 万格。

## 测试与样例

```bash
python scripts/make_sample.py   # 生成 samples/ 演示文件(含透视表样本)
python tests/run_tests.py       # 端到端测试(104 例,全绿)
```

## 项目结构

```
xcli/
  cli.py         命令路由、JSON 输出、错误协议(退出码:0 成功 / 3 业务错误 / 4 内部错误)
  xlutil.py      范围解析(A1:B3/A:C/1:3)、值序列化、原子保存(自动重试,防杀软锁)
  cmds/          10 个命令模块,每个含 NAME/HELP/DESCRIPTION/register/run
scripts/         样本生成(含公式缓存注入、透视表样本)
tests/           端到端测试
skills/          给 AI 用的 skill 示例模板
```

## 已知限制

- 不支持旧版 `.xls`(openpyxl 限制),请先另存为 .xlsx
- 不保留宏:`sheet remove` 之外的修改对 .xlsm 会尽量保留 VBA(keep_vba),但复杂
  文件建议先备份;含 ActiveX 控件等高级对象的文件请用 Excel 验证后再信任结果
- openpyxl 不重算公式:写入公式后要计算结果,请用 Excel/WPS 打开(或 --cached 读旧缓存)
- 透视表只能操作已有的,不能创建;openpyxl 写回的透视表需 Excel 打开时刷新
- 样式操作上限 5 万格、read 默认上限 1 万行(防呆,均可放宽)

# rag-wps-samples — RAG 真实测试物料

本目录的物料**不是 office-cli 自产**,由独立工具(WPS Office COM / python-docx /
Python 标准库)生成,用于对 `office rag prep` 做真实回归(避免"自己造料自己测"的自嗨):

| 文件 | 生成器 | 覆盖点 |
| --- | --- | --- |
| 采购台账.xlsx | make_wps_samples.py(WPS KET) | 两级表头(纵向+横向合并组)、公式(ROUND/SUM/IF/跨表 SUMIF 通配)、税率百分比格、日期格式、整行合并合计行(公式续接)、同 sheet 第二张表、表尾全宽"注:"行、纯文本说明 sheet、隐藏 sheet(应剔除)、中文表名裸引用 |
| 采购台账_legacy.xls | 同上(另存 FileFormat=56) | .xls 老格式 → prep 时经 WPS 自动升级读取,校验内容一致 |
| 评审纪要.docx | make_docx_sample.py(python-docx) | 标题层级、正文、表格纵向合并(建议中标只出现一次,后续行不重复)、编号列表 |
| 设备巡检记录.csv | make_docx_sample.py(标准库 csv) | utf-8-sig BOM + 逗号分隔自动识别 |
| 交接班记录.txt | 同上 | 纯文本 → notes 段落 |

预期质检: `office rag prep -f rag-wps-samples/ --out-dir out/` → 5/5 成功,
公式未解析 0,隐藏表/空表告警仅提示。

生成器坑位备忘(重造物料时注意):
- pywin32 传 naive datetime 给 COM 会按 UTC 折算(-8h),写入前补 `+8h` 再传,
  否则文件里日期偏移一天;
- WPS 中须**先写值再 Merge**(锚格留值);
- 老格式 .xls 中文 sheet 名/合并/百分比格式均兼容,SaveAs FileFormat=56。

"""office — 面向 AI 调用的办公文档 CLI(Excel/Word/PDF/Markdown)。

设计原则:
- 所有命令输出固定结构的 JSON(stdout),错误输出 JSON 到 stderr + 非零退出码
- 无交互、无弹窗,适合作为 AI skill 的执行后端
- 老格式(.xls/.doc)自动经本机 WPS Office 无损升级后操作,原文件不改动
"""

__version__ = "0.3.0"

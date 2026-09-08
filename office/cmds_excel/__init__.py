"""子命令模块。每个模块约定:

NAME        — 顶层子命令名(必填)
HELP        — argparse 一行帮助
DESCRIPTION — 多行详细说明(写入 --help)
register(sp) — 注册参数
run(args)    — 执行并返回可 JSON 序列化的 dict(供 cli.main 打印)
"""

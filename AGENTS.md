# office-cli 项目规范(AGENTS.md)

面向在此仓库协作的 AI 代理与人类开发者。所有交流使用中文。

## 项目是什么

`office-cli` 是一个**面向 AI 调用**的办公文档 CLI:Excel/Word/PDF/Markdown/PowerPoint 读写与互转。所有成功输出为 stdout JSON(UTF-8,中文原文),错误输出 stderr JSON。协议与完整命令手册见 `README.md` 与 `skills/office-ops/SKILL.md`(AI 技能源,两者必须始终与真实命令一致,见「文档铁律」)。

- 退出码:0 成功 / 2 参数错误(argparse)/ 3 业务错误(`{error:{code,message}}`)。
- 错误码约定:`bad_args` `no_file` `bad_json` `file_busy` `merged_cell` `layout_conflict` `encrypted` `bad_password` `not_encrypted` `no_images` `no_match` `same_file` `need_dep` `need_chrome` `wps_unavailable` `wps_timeout` `unsupported_format` `cannot_open` `write_failed`。新错误码须先在 `office/errors.py` 定义。
- 代码与错误文案使用中文;命令参数名用英文。

## 目录结构

```
office/                  # 主包
  cli.py                 # 顶层入口:convert/info + _GROUP_MODS 组注册(importlib 动态加载,缺模块自动跳过)
  cmds_excel/ cmds_word/ cmds_pdf/ cmds_ppt/   # 各组命令模块,一组一目录
  xlutil.py docx2md.py pdfutil.py mdutil.py    # 格式工具库
  ioplan.py              # 老格式升级方案(io plan):.xls/.doc/.ppt 先经 WPS 升级再操作,原文件不动
  wps.py + wps_worker.py # WPS COM 探测与子进程执行器(COM 不能跨进程复用,全部走子进程)
  errors.py              # CliError
tests/run_tests.py       # 黑盒子进程测试(唯一测试入口),期望全绿
scripts/                 # build_exe.py(发布打包) exe_launcher.py make_sample.py(重建 samples/)
samples/                 # 测试与演示样本(demo.xlsx/sample.docx/sample.pptx/sample.md/legacy.*/logo.png)
skills/office-ops/       # office-ops 技能的仓库源(全局副本在 ~/.pi/agent/skills/office-ops/)
.pi/skills/release-office-cli/  # 发布流程技能(发布请调用,全局副本在 ~/.pi/agent/skills/release-office-cli/)
```

- 新命令规范:在对应 `cmds_<组>/` 新建模块,定义 `NAME/HELP/DESCRIPTION` + `register(sp)` + `run(args)->dict`,然后把它加进 `office/cli.py` 的 `_GROUP_MODS` 对应组列表(动态加载按 `name.replace('-','_')` 找模块,`ModuleNotFoundError` 静默跳过——命令不注册等于不存在)。**每个新命令必须**配套:README 总览+示例、SKILL.md 速查、tests/run_tests.py 断言(excel 组新命令同时加入 `_EXCEL_CMDS` 集合以便自动加 `excel` 前缀)。
- 写文件的命令走 ioplan(老格式升级)并原子保存(工具库内已封装 mkstemp+os.replace,别绕过)。

## 文档铁律(血泪教训)

1. **文档必须与实现逐参数一致**。曾整轮凭记忆把 cond-format/layout/comment/validate/insert/delete/replace/footer/search/from-images 的参数写进 README/SKILL.md,冒烟时发现与真实 argparse 完全不符,只能逐文件对照源码重写。
2. 写文档前**先读命令模块源码的 `DESCRIPTION` docstring 与 `register()`**,那里有权威用法示例;禁止凭印象推断 `--xxx` 的存在与语义。
3. SKILL.md(仓库版)改后必须 `cp` 同步到全局 `C:/Users/Administrator/.pi/agent/skills/office-ops/SKILL.md`,否则 PI 加载的是旧版。
4. README/SKILL 与实现冲突时以实现为准,立即修文档,不要迁就文档。

## 版本与发布

- 版本号**两处**同步:`pyproject.toml` 的 `version` 与 `office/__init__.py` 的 `__version__`(遗漏会导致 exe 报旧版,需重打)。
- 发布完整流程调用 `.pi/skills/release-office-cli` 技能(bump→测试→build_exe.py→.smoke 冒烟→gh release create/upload→SHA256 校验)。`office.exe` 单文件约 123MB,不含 playwright(md→pdf 在 exe 内不可用属预期,报 need_dep)。
- Git:单远程 `origin` = https://github.com/fayvo-hub/office-cli.git。推送用 `git push origin master`;直连 GitHub 被重置时用 `git -c http.proxy=http://127.0.0.1:7897 push origin master`(`-c` 在子命令前)。

## 测试

- 唯一入口:`python tests/run_tests.py`(当前 277 passed / 0 failed)。它每次自建并清理 `.test-data/`,子进程真实跑 `python -m office`。
- **禁止直接对 `samples/` 里的样本做写操作测试**(会把样本改脏导致大面积假 FAIL);先复制到 `.test-data/` 或 `.smoke/` 再操作。样本脏了用 `python scripts/make_sample.py` 重建(其中 legacy.xls/.doc/.ppt 由 WPS COM 生成,失败会打印 skip 不中断)。
- 本机 WPS KWPP(演示组件)COM 已损坏:`legacy.ppt` 相关断言带 `if os.path.exists(legacy_ppt)` guard 属正常设计;`ppt to-pdf` 与老格式 ppt 相关能力依赖 WPS,不可用时测试自动跳过,不要当失败去修代码。
- exe 冒烟同规则:复制样本到 `.smoke/`(该目录已在 .gitignore,勿 `git add -A` 误提交)。

## 环境事实

- 开发解释器:D:\Python313(python -m pip install -e ".[mdpdf,wps,images]" 安装,可编辑模式,改码即时生效)。
- 本机无 MS Office/LibreOffice,有 WPS Office 12.1(KWPS/KET COM 正常,KWPP 损坏);WPS COM 必须传 Windows 绝对路径,`/regserver` 与注册表手工修复对 KWPP 无效(环境问题非代码)。
- md→pdf / html→pdf 渲染走系统 Chrome(playwright),exe 版不可用。
- 依赖已在 pyproject 声明(openpyxl/python-docx/PyMuPDF/pdfplumber/pypdf/pdf2docx/python-pptx/Pillow/markdown/pygments/playwright/pywin32),新增依赖必须同步 pyproject 与安装命令。

## 常识约束

- 本仓库不引入 web 框架;CLI 保持零交互(所有输入靠参数/--data-file JSON)。
- 修改前先 `git status`;提交粒度按「一件事一次提交」;消息用中文或英文均可但保持风格一致。
- 交互式弹窗/询问一律用 ask_question;大任务先列 todo。

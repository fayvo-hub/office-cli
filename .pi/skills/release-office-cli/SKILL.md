---
name: release-office-cli
description: 发布 office-cli 新版本(重打包单文件 exe + 发 GitHub Release)。当需要发布/发版 office-cli、bump 版本号、重新打包 office.exe、创建或更新 GitHub Release、修复已发布 exe 的版本号时使用。覆盖版本同步(两处)、测试基线、PyInstaller 打包、exe 冒烟、Release 创建/上传、SHA256 校验全流程。
---

# office-cli 发布流程

在仓库根 `C:/Users/Administrator/AppData/Roaming/pi-desktop/chat-workspace/office-cli` 执行(或任意 office-cli clone 的根)。全程用 Git Bash。仓库有两个远程:`github`(https://github.com/fayvo-hub/office-cli.git,发布目标,master 分支)与 `origin`(阿里云 codeup,旧,勿推)。本机直连 GitHub 会被重置,推送统一用 `git -c http.proxy=http://127.0.0.1:7897 push github master`(`-c` 必须放子命令**前**,`git push -c` 会报 unknown switch)。

## 0. 前置检查

```bash
git status --short          # 工作区干净(排除已忽略的 build/dist/.smoke 等)
python tests/run_tests.py   # 全量黑盒测试,期望 NNN passed / 0 failed(当前 277)
```

- 若改了 skills/office-ops/SKILL.md(仓库版,技能源),先 `cp skills/office-ops/SKILL.md "C:/Users/Administrator/.pi/agent/skills/office-ops/SKILL.md"` 同步全局再发布。
- 本机 WPS KWPP(演示 COM)已损坏:`legacy.ppt` 相关测试被 guard 跳过属正常;`ppt to-pdf` 需 WPS,若不可用测试会自行跳过,勿当失败。

## 1. bump 版本(两处必须同步,否则 exe 报旧版本)

```bash
# pyproject.toml 的 version = "x.y.z"
# office/__init__.py 的 __version__ = "x.y.z"
git add pyproject.toml office/__init__.py && git commit -m "chore: bump x.y.z"
git -c http.proxy=http://127.0.0.1:7897 push github master
```

教训:曾只改 pyproject 忘改 `__init__.py`,打包后 `office --version` 仍报旧版,需重打。

## 2. 打包

```bash
python scripts/build_exe.py    # PyInstaller 单文件,约 95 秒,产物 dist/office.exe(~123 MB)
```

scripts/build_exe.py 已自动处理:动态 importlib 命令模块 hidden-import(含 cmds_ppt 等)、排除本机 ML 大库(cv2 除外,pdf2docx 依赖)、排除 playwright。新增 `office/cmds_<组>/<名>.py` 后无需改 build_exe(glob 自动收集),但需确认命令已注册进 `office/cli.py` 的 `_GROUP_MODS` 对应组列表;若新增 excel 组命令,把命令名加进 `tests/run_tests.py` 的 `_EXCEL_CMDS` 集合(该组测试会按名字自动加 `excel` 前缀)。

## 3. exe 冒烟(勿直接操作 samples/,先复制到 .smoke/)

```bash
rm -rf .smoke && mkdir .smoke && cp samples/sample.pptx samples/demo.xlsx samples/sample.docx samples/sample.md .smoke/ && cd .smoke
../dist/office.exe --version                        # 必须等于新版本号
../dist/office.exe info | grep -E '"version"|"missing"'   # missing 仅 ["playwright"] 属正常
../dist/office.exe excel cond-format add -f demo.xlsx --sheet 销售 --range C2:C20 --op between --value '50,100'
../dist/office.exe excel comment set -f demo.xlsx --sheet 销售 --cell B2 --text 测试
printf '{"slides":[{"layout":"title","title":"冒烟"}]}' > s.json
../dist/office.exe ppt write -f new.pptx --create --data-file s.json && ../dist/office.exe ppt read -f new.pptx --slide 1
../dist/office.exe word replace -f sample.docx --find 模板 --replace demo
../dist/office.exe pdf from-images --out 图.pdf ../samples/logo.png
cd .. && rm -rf .smoke
```

预期:`--version` 为新版本;exe 内 `md to-pdf` 报 `need_dep`(缺 playwright)是**已知限制不是失败**;`insert/delete` 遇合并区报 `layout_conflict` 是保护生效。所有新命令的**真实参数以源码 argparse/docstring 为准**(见 §5)。

## 4. 发 GitHub Release

```bash
gh release create vX.Y.Z --title "office-cli vX.Y.Z" --notes-file /tmp/notes.md   # 标题带 v 前缀
gh release upload vX.Y.Z dist/office.exe --clobber    # 128MB,需一两分钟;若 create 900s 超时留了 draft,upload 后 gh release edit vX.Y.Z --draft=false
```

notes 按模板写:标题、新增(分组件列出新命令)、其他(测试数、文档同步说明、已知限制如 KWPP)。直连若被重置,给 gh 加 `HTTPS_PROXY=http://127.0.0.1:7897` 环境变量重试。

## 5. 验证(必做)

```bash
curl -sL -x http://127.0.0.1:7897 -o /tmp/check.exe https://github.com/fayvo-hub/office-cli/releases/latest/download/office.exe
sha256sum /tmp/check.exe  # 与本地 dist/office.exe 完全一致
gh release view vX.Y.Z --json isDraft,isPrerelease,assets
```

- 直连下载常被重置,curl 必须走 `-x http://127.0.0.1:7897`。
- 发布瞬间资产可能假 404(CDN 延迟)或返回元数据 JSON(需 `Accept: application/octet-stream`),多试一次并用 SHA256 定论。
- 不要推 origin(codeup);文档提交与发布都只走 github remote。

## 6. 收尾

- 确认 git log 干净、master 已包含全部提交(远程 HEAD = 本地 HEAD)。
- 若本轮有任何新教训或新命令,更新本 SKILL 与本仓库 README/SKILL.md 并同步全局。

## 常见问题速查

| 现象 | 原因/处理 |
|---|---|
| `office --version` 旧版本 | `__init__.py` 没同步,重打前先改 |
| push 卡住/重置 | 用 `git -c http.proxy=http://127.0.0.1:7897 push github master` 语法 |
| exe 里某命令消失(26MB 级小 exe) | 缺 hidden-import;新模块要能被 build_exe 的 glob 收到 |
| 测试大面积失败 | 先 `python scripts/make_sample.py` 重建 samples(勿用被改脏的) |
| md to-pdf 在 exe 报 need_dep | 正常限制,exe 不含 playwright,用 pip 完整版 |
| .smoke/.test-data 出现在 git status | 已忽略;`git rm -r --cached` 清理后勿再 add -A 冒烟目录 |

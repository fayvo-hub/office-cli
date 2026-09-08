"""xcli 端到端测试:每个用例用 subprocess 调用真实 CLI,校验退出码与 JSON 输出。

运行: python tests/run_tests.py   (在项目根目录)
输出: 每行 PASS/FAIL,最后汇总;失败可加环境变量 VERBOSE=1 打印输出详情。
"""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLES = os.path.join(ROOT, "samples")
WORK = os.path.join(ROOT, ".test-data")

# excel 组的命令名(旧测试直接以命令名开头,自动补组前缀)
_EXCEL_CMDS = {"list", "read", "write", "sheet", "style", "merge",
               "chart", "image", "pivot"}

VERBOSE = os.environ.get("VERBOSE") == "1"
_passed = 0
_failed = 0
_failures = []


def run(*argv: str) -> tuple[int, dict | None, dict | None]:
    """调用 office;返回 (退出码, stdout JSON, stderr JSON)。

    excel 组命令(list/read/write/...)自动加 "excel" 前缀(兼容旧式调用)。
    """
    if argv and argv[0] in _EXCEL_CMDS:
        argv = ("excel",) + argv
    proc = subprocess.run(
        [sys.executable, "-m", "office", *argv],
        capture_output=True, text=True, encoding="utf-8",
        cwd=ROOT, timeout=60,
    )
    out = err = None
    try:
        out = json.loads(proc.stdout) if proc.stdout.strip() else None
    except json.JSONDecodeError:
        out = {"__raw__": proc.stdout}
    try:
        err = json.loads(proc.stderr) if proc.stderr.strip() else None
    except json.JSONDecodeError:
        err = {"__raw__": proc.stderr}
    return proc.returncode, out, err


def check(name: str, cond: bool, detail: str = "") -> None:
    global _passed, _failed
    if cond:
        _passed += 1
        print(f"PASS  {name}")
    else:
        _failed += 1
        _failures.append(name)
        print(f"FAIL  {name}  {detail[:400]}")


def expect_ok(name, *argv, **kw):
    code, out, err = run(*argv)
    ok = code == 0 and out is not None and out.get("ok") is True
    detail = f"code={code} out={out} err={err}" if not ok or VERBOSE else ""
    check(name, ok, detail)
    return out


def expect_err(name, errcode, *argv):
    code, out, err = run(*argv)
    got = (err is not None and err.get("error", {}).get("code") == errcode
           and code != 0)
    detail = f"code={code} err={err}" if not got or VERBOSE else ""
    check(name, got, detail)


def wpath(name: str) -> str:
    """工作副本路径(测试用,随便改)"""
    return os.path.join(WORK, name)


def copy_sample(name: str) -> str:
    shutil.copy(os.path.join(SAMPLES, name), wpath(name))
    return wpath(name)


# ---------------------------------------------------------------------------

def main() -> int:
    global _passed, _failed
    shutil.rmtree(WORK, ignore_errors=True)
    os.makedirs(WORK)
    demo = copy_sample("demo.xlsx")
    cached = copy_sample("demo_cached.xlsx")
    pivot_f = copy_sample("demo_pivot.xlsx")

    # ---------- list ----------
    out = expect_ok("list 基本", "list", "-f", demo)
    if out:
        names = [s["name"] for s in out["sheets"]]
        check("list 三个表", names == ["销售", "人员", "空表"], str(names))
        s1 = out["sheets"][0]
        check("list 销售表信息", s1["max_row"] == 16 and s1["max_column"] == 5
              and "A16:B16" in s1["merged_cells"] and s1["charts"] == 0,
              json.dumps(s1, ensure_ascii=False)[:300])
        check("list 预览含表头", s1["preview_rows"][0][0] == "月份",
              str(s1["preview_rows"][:1]))
    expect_err("list 文件不存在", "no_file", "list", "-f", wpath("nope.xlsx"))

    # ---------- read ----------
    out = expect_ok("read 全表(默认激活表)", "read", "-f", demo)
    check("read rows=16 行", out and out["row_count"] == 16, str(out))
    check("read D2 是公式文本", out and out["rows"][1][3] == "=B2*C2",
          str(out and out["rows"][1]))
    check("read 日期是 ISO", out and out["rows"][1][4].startswith("2024-01-15"),
          str(out and out["rows"][1][4]))
    check("read 合并区报告", out and "A16:B16" in out["merged_cells"],
          str(out and out["merged_cells"]))

    out = expect_ok("read 指定表", "read", "-f", demo, "--sheet", "人员")
    check("read 人员中文数据", out and out["rows"][1][0] == "张三"
          and out["rows"][1][1] == "北京", str(out and out["rows"][1]))

    out = expect_ok("read 单格", "read", "-f", demo, "--range", "B2")
    check("read 单格值", out and out["rows"] == [[100]], str(out and out["rows"]))

    out = expect_ok("read 区域 + cached", "read", "-f", cached, "--range", "A1:D4",
                    "--cached")
    check("read cached 公式=计算值", out and out["rows"][1][3] == 990.0,
          str(out and out["rows"][1]))

    out = expect_ok("read cells 模式", "read", "-f", cached, "--range", "D2",
                    "--cells", "--cached")
    c = out["cells"][0] if out else {}
    check("read cells 元数据", c.get("type") == "formula" and c.get("formula") == "=B2*C2"
          and c.get("value") == 990.0 and c.get("cached_available") is True,
          json.dumps(c, ensure_ascii=False))

    out = expect_ok("read cells 注释与超链接", "read", "-f", demo, "--range", "B2",
                    "--cells")
    c = out["cells"][0] if out else {}
    check("read comment/hyperlink", "comment" in c and c["comment"].startswith("1 月销量"),
          json.dumps(c, ensure_ascii=False)[:200])

    out = expect_ok("read 整列 + limit", "read", "-f", demo, "--range", "A:A",
                    "--limit", "3")
    check("read 截断警告", out and out["row_count"] == 3
          and any("limit" in w for w in out["warnings"]),
          str(out and (out["row_count"], out["warnings"])))

    out = expect_ok("read 空表", "read", "-f", demo, "--sheet", "空表")
    check("read 空表 rows=[]", out and out["rows"] == [], str(out))

    expect_err("read 表不存在", "no_sheet", "read", "-f", demo, "--sheet", "不存在")
    expect_err("read 坏范围", "bad_range", "read", "-f", demo, "--range", "A0")

    # ---------- write ----------
    out = expect_ok("write 单值", "write", "-f", demo, "--cell", "C15", "--data", "123.5")
    check("write 返回范围", out and out["range"] == "C15", str(out))
    out = expect_ok("write 后读回", "read", "-f", demo, "--range", "C15")
    check("write 值生效", out and out["rows"] == [[123.5]], str(out and out["rows"]))

    out = expect_ok("write 二维+公式", "write", "-f", demo, "--cell", "F1",
                    "--data", '[[1, "=F1*2"], [2, "文字"]]')
    check("write 范围 F1:G2", out and out["range"] == "F1:G2"
          and out["values_written"] == 4, str(out))
    out = expect_ok("write 公式读回", "read", "-f", demo, "--range", "G1", "--cells")
    check("write 公式格", out and out["cells"][0]["type"] == "formula", str(out))

    out = expect_ok("write --literal", "write", "-f", demo, "--cell", "F4",
                    "--data", '"=NOT_A_FORMULA"', "--literal")
    out = expect_ok("write literal 读回", "read", "-f", demo, "--range", "F4",
                    "--cells")
    check("write literal 是文本", out and out["cells"][0]["type"] == "str"
          and out["cells"][0]["value"] == "=NOT_A_FORMULA", json.dumps(out, ensure_ascii=False))

    jf = wpath("vals.json")
    with open(jf, "w", encoding="utf-8") as fh:
        json.dump([["甲", 1], ["乙", 2]], fh, ensure_ascii=False)
    expect_ok("write --data-file", "write", "-f", demo, "--cell", "H1",
              "--data-file", jf)

    created = wpath("created.xlsx")
    expect_ok("write --create 新建", "write", "-f", created, "--cell", "A1",
              "--data", "1", "--create")
    expect_err("write --create 撞已存在", "file_exists", "write", "-f", created,
               "--cell", "A1", "--data", "1", "--create")
    expect_ok("write --create --overwrite", "write", "-f", created, "--cell", "A1",
              "--data", '"x"', "--create", "--overwrite")
    expect_err("write 文件不存在(无 create)", "no_file", "write", "-f",
               wpath("missing.xlsx"), "--data", "1")
    expect_err("write 坏 JSON", "bad_json", "write", "-f", created, "--data", "{oops")
    expect_err("write 越界", "too_large", "write", "-f", created, "--cell", "A1048570",
               "--data", '[[1],[2],[3],[4],[5],[6],[7],[8],[9],[10]]')

    # ---------- sheet ----------
    out = expect_ok("sheet add", "sheet", "add", "-f", demo, "--name", "新表")
    check("sheet add 出现", out and "新表" in out["sheets"], str(out))
    expect_err("sheet add 重名", "dup_sheet", "sheet", "add", "-f", demo,
               "--name", "销售")
    expect_ok("sheet rename", "sheet", "rename", "-f", demo, "--name", "新表",
              "--new-name", "改后")
    out = expect_ok("sheet copy", "sheet", "copy", "-f", demo, "--name", "销售",
                    "--new-name", "销售副本")
    check("sheet copy 含副本", out and "销售副本" in out["sheets"], str(out))
    out = expect_ok("sheet remove", "sheet", "remove", "-f", demo, "--name", "改后")
    check("sheet remove 生效", out and "改后" not in out["sheets"], str(out))
    out = expect_ok("sheet remove 销售副本", "sheet", "remove", "-f", demo,
                    "--name", "销售副本")

    # active 用独立副本,避免改变 demo 的默认激活表影响后续用例
    active_sh = copy_sample("demo.xlsx")
    expect_ok("sheet active", "sheet", "active", "-f", active_sh, "--name", "人员")
    out = expect_ok("sheet active 后 list", "list", "-f", active_sh)
    wb_tmp = load_workbook(active_sh)
    check("sheet active 生效", wb_tmp.active.title == "人员", wb_tmp.active.title)

    # ---------- style ----------
    st = wpath("style.xlsx")
    shutil.copy(demo, st)
    out = expect_ok("style 表头", "style", "-f", st, "--sheet", "销售",
                    "--range", "A1:E1",
                    "--bold", "--fill", "FFC000", "--align", "center",
                    "--font-size", "13", "--font-color", "FFFFFF",
                    "--border", "thin", "--num-format", "0.0")
    check("style cells_affected=5", out and out["cells_affected"] == 5, str(out))
    wb = load_workbook(st)
    c = wb["销售"]["A1"]
    check("style 字体生效", c.font.bold and c.font.size == 13, str(c.font))
    check("style 填充生效", c.fill.fgColor.rgb in ("00FFC000", "FFFFC000"),
          str(c.fill.fgColor.rgb))
    check("style 数字格式生效", c.number_format == "0.0", c.number_format)
    expect_err("style 颜色非法", "bad_color", "style", "-f", st, "--range", "A1",
               "--fill", "red")
    expect_err("style 区域过大", "too_large", "style", "-f", st,
               "--range", "A1:C20000", "--bold")

    # ---------- merge ----------
    mg = wpath("merge.xlsx")
    shutil.copy(demo, mg)
    S = ["--sheet", "销售"]
    out = expect_ok("merge 合并", "merge", "-f", mg, *S, "--range", "C17:D17")
    check("merge 报告合并区", out and "C17:D17" in out["merged_cells"], str(out))
    # 向合并区非左上角写入 -> 业务错误 merged_cell(而非 internal)
    expect_err("write 合并区非左上格", "merged_cell", "write", "-f", mg, *S,
               "--cell", "D17", "--data", "1")
    # 合并区左上角仍可写(数据只在左上角)
    expect_ok("write 合并区左上角", "write", "-f", mg, *S, "--cell", "C17",
              "--data", '"更新"')
    expect_err("merge 重叠", "overlap", "merge", "-f", mg, *S, "--range", "D17:E17")
    out = expect_ok("write 准备合并警告", "write", "-f", mg, *S, "--cell", "A18",
                    "--data", '"保留"')
    expect_ok("write 准备合并警告2", "write", "-f", mg, *S, "--cell", "B18",
              "--data", '"丢弃我"')
    out = expect_ok("merge 非左上值警告", "merge", "-f", mg, *S, "--range", "A18:B18")
    check("merge 警告存在", out and any("B18" in w for w in out["warnings"]),
          str(out and out["warnings"]))
    out = expect_ok("merge unmerge", "merge", "-f", mg, *S, "--range", "C17:D17",
                    "--unmerge")
    check("merge unmerge 生效", out and "C17:D17" not in out["merged_cells"], str(out))
    expect_err("merge unmerge 非合并区", "not_merged", "merge", "-f", mg, *S,
               "--range", "E1:F1", "--unmerge")

    # ---------- convert ----------
    o_csv = wpath("out.csv")
    expect_ok("convert xlsx->csv", "convert", "-f", cached, "--sheet", "销售",
              "--out", o_csv, "--cached")
    with open(o_csv, encoding="utf-8-sig") as fh:
        lines = fh.readlines()
    check("convert csv BOM+公式值", lines[1].startswith("2024-01,")
          and "990.0" in lines[1], lines[1] if lines else "")

    o_json = wpath("out.json")
    out = expect_ok("convert xlsx->json", "convert", "-f", cached, "--sheet", "销售",
                    "--out", o_json, "--cached")
    with open(o_json, encoding="utf-8") as fh:
        doc = json.load(fh)
    check("convert json 结构", doc["sheet"] == "销售" and len(doc["rows"]) == 16
          and doc["rows"][1][3] == 990.0, str(doc["rows"][1] if "rows" in doc else doc))

    x1 = wpath("from_gbk.xlsx")
    expect_ok("convert gbk csv->xlsx", "convert", "-f",
              os.path.join(SAMPLES, "gbk.csv"), "--out", x1)
    out = expect_ok("read gbk 转换结果", "read", "-f", x1)
    check("convert gbk 中文+推断", out and out["rows"][1][0] == "张三"
          and out["rows"][1][2] == 1200.5 and out["rows"][1][3] is True,
          str(out and out["rows"][1]))

    x2 = wpath("from_json.xlsx")
    expect_ok("convert json->xlsx", "convert", "-f", o_json, "--out", x2)
    out = expect_ok("read json 转换结果", "read", "-f", x2, "--sheet", "Sheet1")
    check("convert json 回写", out and out["rows"][0][0] == "月份", str(out and out["rows"][0]))

    o_csv2 = wpath("out2.json")
    expect_ok("convert csv->json", "convert", "-f", os.path.join(SAMPLES, "utf8.csv"),
              "--out", o_csv2)
    expect_err("convert 不支持方向", "unsupported", "convert", "-f", demo,
               "--out", wpath("a.xls"))
    expect_err("convert 同路径", "same_file", "convert", "-f", demo, "--out", demo)

    # ---------- chart ----------
    ch = wpath("chart.xlsx")
    shutil.copy(cached, ch)
    out = expect_ok("chart add col", "chart", "add", "-f", ch, "--type", "col",
                    "--data", "A1:D13", "--at", "H2", "--title", "月度销售")
    check("chart series=3", out and out["series"] == 3 and out["anchor"] == "H2", str(out))
    expect_ok("chart add pie", "chart", "add", "-f", ch, "--type", "pie",
              "--data", "A1:B13")
    expect_err("chart pie 多系列报错", "bad_args", "chart", "add", "-f", ch,
               "--type", "pie", "--data", "A1:C13")
    expect_err("chart 单列报错", "bad_args", "chart", "add", "-f", ch,
               "--type", "col", "--data", "A1:A13")
    wb = load_workbook(ch)
    check("chart 保存后可重开", wb is not None, "")

    # ---------- image ----------
    im = wpath("image.xlsx")
    shutil.copy(demo, im)
    out = expect_ok("image add", "image", "add", "-f", im, "--image",
                    os.path.join(SAMPLES, "logo.png"), "--at", "B2", "--width", "120")
    check("image 尺寸", out and out["width_px"] == 120 and out["height_px"] == 48,
          str(out))
    expect_err("image 文件不存在", "no_file", "image", "add", "-f", im,
               "--image", wpath("no.png"), "--at", "B2")

    # ---------- pivot ----------
    out = expect_ok("pivot list", "pivot", "list", "-f", pivot_f)
    p0 = out["pivots"][0] if out and out["pivots"] else {}
    check("pivot list 内容", p0.get("name") == "产品透视"
          and p0.get("source_range") == "$A$1:$C$31" and p0.get("sheet") == "透视",
          json.dumps(p0, ensure_ascii=False))
    out = expect_ok("pivot set-source", "pivot", "set-source", "-f", pivot_f,
                    "--name", "产品透视", "--ref", "A1:C40")
    check("pivot set-source 范围", out and out["source_range"] == "$A$1:$C$40"
          and out["refresh_on_load"] is True, str(out))
    out = expect_ok("pivot list 验证", "pivot", "list", "-f", pivot_f)
    p0 = out["pivots"][0] if out and out["pivots"] else {}
    check("pivot 修改已落盘", p0.get("source_range") == "$A$1:$C$40",
          json.dumps(p0, ensure_ascii=False))
    expect_err("pivot 名字不存在", "no_pivot", "pivot", "set-source", "-f", pivot_f,
               "--name", "不存在", "--ref", "A1:B2")
    out = expect_ok("pivot list 无透视表文件", "pivot", "list", "-f", demo)
    check("pivot list 空", out and out["pivots"] == [], str(out))

    # ---------- 汇总 ----------
    print()
    print(f"==== {_passed} passed, {_failed} failed ====")
    if _failures:
        print("failures:", *_failures, sep="\n  - ")
    return 1 if _failed else 0


if __name__ == "__main__":
    sys.exit(main())

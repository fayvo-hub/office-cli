"""WPS COM 转换 worker —— 在子进程内执行,防止 COM 卡死主 CLI。

用法: python -m office.wps_worker '<job-json>'
job: {"op": "xls_to_xlsx"|"doc_to_docx"|"docx_to_pdf"|"doc_to_pdf"|"probe",
      "src": "...", "dst": "..."}
输出到 stdout 的 JSON: {"ok": true} 或 {"ok": false, "error": "..."}
"""

# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import sys


def _app(progid: str):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx(progid)
    try:
        app.Visible = False
    except Exception:
        pass
    try:
        app.DisplayAlerts = 0
    except Exception:
        pass
    return app


def _quit(app) -> None:
    try:
        app.Quit()
    except Exception:
        pass


def xls_to_xlsx(src: str, dst: str) -> None:
    app = _app("KET.Application")
    wb = None
    try:
        wb = app.Workbooks.Open(src, ReadOnly=False)
        # 51 = xlOpenXMLWorkbook (.xlsx); 覆盖已存在文件
        wb.SaveAs(dst, FileFormat=51)
        wb.Close(False)
        wb = None
    finally:
        _quit(app)


def doc_to_docx(src: str, dst: str) -> None:
    app = _app("KWPS.Application")
    doc = None
    try:
        doc = app.Documents.Open(src, ReadOnly=False)
        # 16 = wdFormatXMLDocument (.docx)
        doc.SaveAs2(dst, FileFormat=16)
        doc.Close(False)
        doc = None
    finally:
        _quit(app)


def docx_to_pdf(src: str, dst: str) -> None:
    app = _app("KWPS.Application")
    doc = None
    try:
        doc = app.Documents.Open(src, ReadOnly=True)
        # 0 = wdExportFormatPDF
        doc.ExportAsFixedFormat(0, dst)
        doc.Close(False)
        doc = None
    finally:
        _quit(app)


def doc_to_pdf(src: str, dst: str) -> None:
    """.doc -> .pdf: 先转 docx 临时文件再导出(避免老格式兼容问题)。"""
    import tempfile

    fd, tmp_docx = tempfile.mkstemp(prefix=".office-doc-", suffix=".docx")
    os.close(fd)
    try:
        app = _app("KWPS.Application")
        doc = None
        try:
            doc = app.Documents.Open(src, ReadOnly=True)
            doc.SaveAs2(tmp_docx, FileFormat=16)
            doc.Close(False)
            doc = None
            doc2 = app.Documents.Open(tmp_docx, ReadOnly=True)
            doc2.ExportAsFixedFormat(0, dst)
            doc2.Close(False)
            doc2 = None
        finally:
            _quit(app)
    finally:
        try:
            os.remove(tmp_docx)
        except OSError:
            pass


_OPS = {
    "xls_to_xlsx": xls_to_xlsx,
    "doc_to_docx": doc_to_docx,
    "docx_to_pdf": docx_to_pdf,
    "doc_to_pdf": doc_to_pdf,
}


def main() -> int:
    job = json.loads(sys.argv[1])
    op = job["op"]
    if op == "probe":
        # 探测:能启动应用即可
        for progid in ("KWPS.Application", "KET.Application"):
            try:
                app = _app(progid)
                _quit(app)
                print(json.dumps({"ok": True}))
                return 0
            except Exception:
                continue
        print(json.dumps({"ok": False, "error": "WPS COM 不可用"}))
        return 1
    fn = _OPS.get(op)
    if fn is None:
        print(json.dumps({"ok": False, "error": f"未知 op: {op}"}))
        return 1
    try:
        fn(job["src"], job["dst"])
        print(json.dumps({"ok": True}))
        return 0
    except Exception as e:  # noqa: BLE001 - worker 出口,全部转 JSON
        print(json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"}))
        return 1


if __name__ == "__main__":
    sys.exit(main())

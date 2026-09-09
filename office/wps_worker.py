"""WPS COM 转换 worker —— 在子进程内执行,防止 COM 卡死主 CLI。

用法: python -m office.wps_worker '<job-json>'
job: {"op": "xls_to_xlsx"|"doc_to_docx"|"docx_to_pdf"|"doc_to_pdf"|
     "ppt_to_pptx"|"pptx_to_pdf"|"ppt_to_pdf"|"probe",
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
        # (文件名, 17=wdExportFormatPDF) —— WPS KWPS 以文件名为第一参数
        doc.ExportAsFixedFormat(dst, 17)
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
            doc2.ExportAsFixedFormat(dst, 17)
            doc2.Close(False)
            doc2 = None
        finally:
            _quit(app)
    finally:
        try:
            os.remove(tmp_docx)
        except OSError:
            pass


def ppt_to_pptx(src: str, dst: str) -> None:
    """.ppt 老格式 -> .pptx(24 = ppSaveAsOpenXMLPresentation)。"""
    app = _app("KWPP.Application")
    pres = None
    try:
        pres = app.Presentations.Open(os.path.abspath(src), ReadOnly=True)
        pres.SaveAs(os.path.abspath(dst), 24)
        pres.Close()
        pres = None
    finally:
        _quit(app)


def pptx_to_pdf(src: str, dst: str) -> None:
    """.pptx -> .pdf(32 = ppSaveAsPDF)。"""
    _ppt_save_pdf(src, dst)


def ppt_to_pdf(src: str, dst: str) -> None:
    """.ppt 老格式 -> .pdf(直接导出,不经过 pptx)。"""
    _ppt_save_pdf(src, dst)


def _ppt_save_pdf(src: str, dst: str) -> None:
    app = _app("KWPP.Application")
    pres = None
    try:
        # 目标已存在时 SaveAs 可能弹覆盖确认,先删除
        try:
            os.remove(os.path.abspath(dst))
        except OSError:
            pass
        pres = app.Presentations.Open(os.path.abspath(src), ReadOnly=True)
        # KWPP 的 ExportAsFixedFormat 签名与 MS 不同,SaveAs(路径, 32) 实测可行
        pres.SaveAs(os.path.abspath(dst), 32)
        pres.Close()
        pres = None
    finally:
        _quit(app)


_OPS = {
    "xls_to_xlsx": xls_to_xlsx,
    "doc_to_docx": doc_to_docx,
    "docx_to_pdf": docx_to_pdf,
    "doc_to_pdf": doc_to_pdf,
    "ppt_to_pptx": ppt_to_pptx,
    "pptx_to_pdf": pptx_to_pdf,
    "ppt_to_pdf": ppt_to_pdf,
}


def main() -> int:
    return worker_main(sys.argv[1])


def worker_main(payload: dict | str) -> int:
    """执行一次转换任务(payload 为 dict 或 JSON 字符串),返回进程退出码。"""
    job = payload if isinstance(payload, dict) else json.loads(payload)
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

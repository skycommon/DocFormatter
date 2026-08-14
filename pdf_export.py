# -*- coding: utf-8 -*-
"""离线把 .docx 转换为 .pdf。

优先使用本机 Microsoft Word（COM 自动化），其次尝试 LibreOffice 无头转换；
两者都不可用时返回 False，由调用方给出友好提示。
"""

import os
import shutil
import subprocess


def docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    """把 docx 转成 pdf。成功返回 True，否则 False。"""
    docx_path = os.path.abspath(docx_path)
    pdf_path = os.path.abspath(pdf_path)

    # 1) Microsoft Word COM
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        try:
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 = wdFormatPDF
        finally:
            doc.Close(0)
            word.Quit()
        return os.path.exists(pdf_path)
    except Exception:
        pass

    # 2) LibreOffice 无头转换
    soffice = _find_soffice()
    if soffice:
        try:
            outdir = os.path.dirname(pdf_path)
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf",
                 "--outdir", outdir, docx_path],
                check=True, stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL, timeout=180)
            return os.path.exists(pdf_path)
        except Exception:
            pass

    return False


def _find_soffice():
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        r"D:\Program Files\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return shutil.which("soffice") or shutil.which("libreoffice")

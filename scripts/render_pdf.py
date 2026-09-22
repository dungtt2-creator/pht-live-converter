# -*- coding: utf-8 -*-
"""Render docx -> pdf bằng Word COM (verify layout). Usage: python render_pdf.py <docx> [out.pdf]"""
import sys, os, win32com.client

src = os.path.abspath(sys.argv[1])
out = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.splitext(src)[0] + ".pdf"
if os.path.exists(out):
    try: os.remove(out)
    except OSError: pass

word = win32com.client.DispatchEx("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(src, ReadOnly=True)
    doc.SaveAs2(out, FileFormat=17)  # wdFormatPDF
    doc.Close(False)
    print("OK:", out)
finally:
    word.Quit()
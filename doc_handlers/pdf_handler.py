# -*- coding: utf-8 -*-
"""
PDF 文档打印：通过 Windows 默认关联程序打印（如 Adobe / Edge）
"""
import os
import subprocess
import sys

try:
    import win32api
    import win32print
except ImportError:
    win32api = None
    win32print = None


def print_pdf(pdf_path, printer_name=None):
    """
    使用系统默认方式打印 PDF。
    Windows: 使用 ShellExecute 的 print 动词，会调用与 .pdf 关联的程序。
    printer_name 为 None 时使用默认打印机；若需指定打印机，可先临时设置默认打印机再打印。
    """
    path = os.path.abspath(pdf_path)
    if not os.path.isfile(path):
        raise FileNotFoundError(path)
    if sys.platform != "win32":
        raise RuntimeError("当前仅支持 Windows 打印 PDF")
    if win32api is None:
        raise RuntimeError("请安装 pywin32: pip install pywin32")
    try:
        if printer_name and win32print:
            old_default = win32print.GetDefaultPrinter()
            try:
                win32print.SetDefaultPrinter(printer_name)
                win32api.ShellExecute(0, "print", path, None, ".", 0)
            finally:
                win32print.SetDefaultPrinter(old_default)
        else:
            win32api.ShellExecute(0, "print", path, None, ".", 0)
        return True
    except Exception:
        raise

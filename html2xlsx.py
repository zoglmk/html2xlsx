#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2018-10-09 11:58:26
# @Author  : zgm (heykeener@gmail.com)
# @Link    : https://recordme.club
# @Version : 0.1

from win32com.client import Dispatch


def htmltoxlsx(file_path, save_filename):
    """convert a table in html to xlsx
    only on Windows.
    Args:
        file_path: the open file absolute path,including filename and ext.
        save_filename: want to save filename
        eg.'F:\\Python\\test\\example.html'
    Returns:
        save a Excel(.xlsx) file.

    error:
        return the error infomation.
    """
    try:
        excelApp = Dispatch('Excel.Application')
        book = excelApp.Workbooks.open(file_path)
        book.SaveAs(save_filename + '.xlsx', 51)
        book.close
        print("转换完成 / Converted successful.")
    except Exception as e:
        print(e)


htmltoxlsx('F:\\Python\\test\\安徽example.html', 'F:\\Python\\test\\example')

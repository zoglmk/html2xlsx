#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Date    : 2018-10-09 11:58:26
# @Author  : zgm (heykeener@gmail.com)
# @Link    : https://recordme.club
# @Version : 0.1

from win32com.client import Dispatch

excelApp = Dispatch('Excel.Application')
book = excelApp.Workbooks.open('html file path')
book.SaveAs('xlsx save path' + '.xlsx', 51)
book.close

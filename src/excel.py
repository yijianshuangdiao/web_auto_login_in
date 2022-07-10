#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2022/7/10 10:33
# @Author  : YangGuo
# @FileName: excel.py
# @Software: PyCharm

import xlrd


class ExcelReader:

    def __init__(self, file_name):
        self.xls_file = xlrd.open_workbook(file_name)

    def GetSheetsNames(self):
        """返回workbook中所有的表名"""
        return self.xls_file.sheet_names()

    def GetSheetSize(self, sheet_name):
        """返回表名对应的sheet的行数和列数"""
        if not sheet_name in self.GetSheetsNames():
            return 0, 0
        sheet = self.xls_file.sheet_by_name(sheet_name)
        return sheet.nrows, sheet.ncols

    def GetSheetContent(self, sheet_name):
        """以二位列表，返回sheet的内容"""
        if not sheet_name in self.GetSheetsNames():
            return []

        sheet = self.xls_file.sheet_by_name(sheet_name)
        nr, nc = sheet.nrows, sheet.ncols
        return [[sheet.cell(r, c).value for c in range(nc)] for r in range(nr)]

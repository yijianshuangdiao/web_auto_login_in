#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2022/7/10 13:28
# @Author  : YangGuo
# @FileName: model.py
# @Software: PyCharm
from PyQt5 import Qt, QtCore
from PyQt5.QtCore import QAbstractTableModel


class TableModel(QAbstractTableModel):
    """
    表格数据模型MVC模式
    """

    def __init__(self, data, HEADERS):
        super(TableModel, self).__init__()
        self._data = data
        self.headers = HEADERS

    def updateData(self, data):
        """
        (自定义)更新数据
        """
        self.beginResetModel()
        self._data = data
        self.endResetModel()

    def data(self, index, role=None):
        if role == QtCore.Qt.DisplayRole:
            value = self._data[index.row()][index.column()]
            return value

        if role == QtCore.Qt.DecorationRole:
            pass

        if role == QtCore.Qt.TextAlignmentRole:
            return QtCore.Qt.AlignCenter

    def rowCount(self, parent=None, *args, **kwargs):
        """
        行数
        """
        return len(self._data)

    def columnCount(self, parent=None, *args, **kwargs):
        """
        列数
        """
        return len(self.headers)

    def headerData(self, section, orientation, role=QtCore.Qt.DisplayRole):
        """
        标题定义
        """
        if role != QtCore.Qt.DisplayRole:
            return None
        if orientation == QtCore.Qt.Horizontal:
            return self.headers[section]
        return int(section + 1)

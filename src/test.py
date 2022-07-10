#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# @Time    : 2022/7/10 9:40
# @Author  : YangGuo
# @FileName: test.py
# @Software: PyCharm
import os
import sys
import threading
import time

from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication, QMainWindow, QHeaderView
from selenium.webdriver.edge.service import Service

from mainwindow import Ui_MainWindow
from src.excel import ExcelReader
from src.model import TableModel


class MainWindow(Ui_MainWindow, QMainWindow):
    def __init__(self):
        super().__init__()
        self.m_exePath = None
        self.m_data = None
        self.m_tableModel = None
        self.m_resPath = None
        self.m_counts = 0
        self.setupUi(self)
        self.initView()
        self.initConnect()

    def initView(self):
        # 水平方向标签拓展剩下的窗口部分，填满表格
        self.tableView.horizontalHeader().setStretchLastSection(True)
        # 水平方向，表格大小拓展到适当的尺寸
        self.tableView.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 进度条
        self.progressBar.setValue(0)

    def initConnect(self):
        self.exePathButton.clicked.connect(self.selectExe)
        self.resPathButton.clicked.connect(self.selectRes)
        self.loadButton.clicked.connect(self.loadRes)
        self.runButton.clicked.connect(self.runTask)

    def selectExe(self):
        filePath, fileType = QtWidgets.QFileDialog.getOpenFileName(self, "配置工作环境", os.getcwd(), "exe(*exe)")
        self.exeLineEdit.setText(filePath)
        self.m_exePath = filePath

    def selectRes(self):
        filePath, fileType = QtWidgets.QFileDialog.getOpenFileName(self, "选择资源", os.getcwd(), "资源文件(*xlsx)")
        self.resLineEdit.setText(filePath)
        self.m_resPath = filePath

    def loadRes(self):
        excelObject = ExcelReader(self.m_resPath)
        sheetNameList = excelObject.GetSheetsNames()
        nRows, nCols = excelObject.GetSheetSize(sheetNameList[0])
        self.m_counts = nRows
        resData = excelObject.GetSheetContent(sheetNameList[0])
        headers = resData[0]
        resData = resData[1:]
        self.m_data = []
        for name, phone_number in resData:
            self.m_data.append([name, str(int(phone_number))])
        self.m_tableModel = TableModel(data=self.m_data, HEADERS=headers)
        self.tableView.setModel(self.m_tableModel)

    def runTask(self):
        # 进度条
        self.progressBar.setValue(0)
        self.progressBar.setRange(0, self.m_counts - 1)

        threading.Thread(target=self.__run).start()

    def __run(self):
        """用子线程跑核心代码，避免前端卡顿"""

        # 链接网页
        executable_path = self.m_exePath
        browser = webdriver.Edge(service=Service(executable_path))
        browser.get(
            "https://fz2022zsh5.hnbluestorm.com/?code=031Hxp100i1pbO1u2U300u3Oy74Hxp10&state=STATE#/pages/user/download?nID=287792987476087")

        # 点击“已安装上报信息“
        browser.find_element(By.XPATH,
                             "/html/body/uni-app/uni-page/uni-page-wrapper/uni-page-body/uni-view/uni-view[2]/uni-button[2]").click()

        time.sleep(2.0)  # 给浏览器预留2s反应时间
        try:
            name_element = browser.find_element(By.XPATH,
                                                "/html/body/uni-app/uni-page/uni-page-wrapper/uni-page-body/uni-view/uni-view[3]/uni-view[2]/uni-view[1]/uni-scroll-view/div/div/div/uni-view/uni-view[1]/uni-view[2]/uni-input/div/input")
            phone_number_element = browser.find_element(By.XPATH,
                                                        "/html/body/uni-app/uni-page/uni-page-wrapper/uni-page-body/uni-view/uni-view[3]/uni-view[2]/uni-view[1]/uni-scroll-view/div/div/div/uni-view/uni-view[1]/uni-view[3]/uni-input/div/input")
            confirm_element = browser.find_element(By.XPATH,
                                                   "/html/body/uni-app/uni-page/uni-page-wrapper/uni-page-body/uni-view/uni-view[3]/uni-view[2]/uni-view[1]/uni-scroll-view/div/div/div/uni-view/uni-view[2]/uni-view[2]")
        except NoSuchElementException:
            window.close()

        finished_counts = 0
        for name, phone_number in self.m_data:
            name_element.click()
            name_element.clear()
            name_element.send_keys(name)
            phone_number_element.click()
            phone_number_element.clear()
            phone_number_element.send_keys(phone_number)
            confirm_element.click()

            finished_counts += 1
            self.progressBar.setValue(finished_counts)

            time.sleep(0.1)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

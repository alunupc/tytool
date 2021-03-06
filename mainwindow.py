# -*- coding:utf8 -*-
import json
import os
import re
import sys

import camelot
import matplotlib
import numpy as np
import pymysql
import xlrd
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QTreeWidget, QWidget, QTreeWidgetItem, QHBoxLayout, QApplication, QLabel, \
    QLineEdit, QVBoxLayout, QGridLayout, QGroupBox, QPushButton, QMessageBox, QFileDialog, \
    QComboBox, QListView, QTextEdit
from matplotlib import pylab
from matplotlib.pyplot import axis

from downloader import DownLoader
from excel import Excel
from lineedit import LineEdit
from log import Logger
from node import MultiTree, TreeNode
from pageinfo import PageInfo
from utils import process_param, process_file

matplotlib.use("Qt5Agg")


class RunThread(QtCore.QThread):
    resSignal = QtCore.pyqtSignal(str, MultiTree)

    def __init__(self, root, tree, year):
        super(RunThread, self).__init__()
        self.root = root
        self.tree = tree
        self.year = year
        self.log = Logger('init.log', level='error')

    def __del__(self):
        self.wait()

    def init_tree_widget(self, root, count, my_tree):
        if isinstance(my_tree, MultiTree):
            if len(my_tree.tree.children) == 0:
                return
            self.init_tree_widget(root, count + 1, my_tree.tree)
        else:
            child_item = QTreeWidgetItem(root)
            child_item.setText(0, my_tree.data.get('code') + " " + my_tree.data.get('name'))
            for child in my_tree.children:
                self.init_tree_widget(child_item, count + 1, child)

    def generate_tree_by_year(self, year):
        try:
            # 建立连接，使用pymysql中的connect方法
            conn = pymysql.connect(host='192.168.110.201',
                                   user='root',
                                   passwd='tengyun2020',
                                   db='fiscal_tax',
                                   port=3306,
                                   charset='utf8',
                                   cursorclass=pymysql.cursors.DictCursor)
            # DictCursor游标类，返回单个结果为字典，多个结果为嵌套字典的列表, 不添加cursorclass游标嵌套元组的元组
            # 创建游标，使用连接当中的cursor方法， 使用游标对象进行接受cursor
            cursor = conn.cursor()
        except Exception as e:
            # QMessageBox.information(self, "提示", '    数据库连接失败！    ')
            print(e)
            self.log.logger.error(e)

        try:
            query_sql = "SELECT id, pid, code, name, type, gid, flag FROM ft_subject WHERE gid=%s ORDER BY id ASC"
            cursor.execute(query_sql, 0)
            res = cursor.fetchall()
            for row in res:
                self.tree.add(TreeNode(row), row.get('pid'), self.tree.tree)
            # 查询树状结构
            query_sql = """
                    SELECT
                        t.*,
                        CASE
                    WHEN t.pid >= 0 THEN
                        t.code
                    ELSE
                         t.code
                    END code
                    FROM
                        ft_subject t,
                        ft_subject t1
                    WHERE
                        t.pid >= 5
                    AND t.pid = t1.id AND t.date_year=%s 
                    ORDER BY
                        t.pid
                    """
            cursor.execute(query_sql, year)
            conn.commit()
            results = cursor.fetchall()
            for row in results:
                self.tree.add(TreeNode(row), row.get('pid'), self.tree.tree)
            # self.tree.traverse(self.tree)

            self.init_tree_widget(self.root, 0, self.tree.tree)
        except Exception as e:
            print(e)
            self.log.logger.error(e)
            conn.rollback()
            # QMessageBox.information(self, "提示", '    数据库查询失败！    ')
        try:
            # 关闭连接释放资源，先关闭游标
            cursor.close()
            # 再关闭连接对象
            conn.close()
        except Exception as e:
            print(e)
            self.log.logger.error(e)
            # QMessageBox.information(self, "提示", '    关闭数据库失败！    ')

    def run(self):
        print("Access the thread!")
        self.generate_tree_by_year(self.year)
        self.resSignal.emit(self.year, self.tree)  # 任务完成后，发送信号


class DownloadThread(QtCore.QThread):
    resSignal = QtCore.pyqtSignal(str)

    def __init__(self, url, path):
        super(DownloadThread, self).__init__()
        self.url = url
        self.path = path
        self.log = Logger('downloadthread.log', level='error')

    def __del__(self):
        self.wait()

    def run(self):
        # print("")
        try:
            downloader = DownLoader(timeout=30, url=self.url, path=self.path)
            downloader.download_file()
            self.resSignal.emit("ok")  # 任务完成后，发送信号
        except Exception as e:
            self.resSignal.emit("fail")  # 任务完成后，发送信号
            self.log.logger.error(e)
            # self.resSignal.emit(e)  # 任务完成后，发送信号
            # QMessageBox.information(self, "提示", '    ！    ')


class MainWindow(QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__()
        self.pageInfo = PageInfo("", "")
        self.setWindowTitle("河北腾云信息科技有限公司")
        self.icon = QtGui.QIcon()
        self.icon.addPixmap(QtGui.QPixmap("images/logo.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(self.icon)
        self.treeWidget = QTreeWidget()
        self.centralWidget = QWidget()
        self.setCentralWidget(self.centralWidget)
        self.root = QTreeWidgetItem(self.treeWidget)
        self.treeWidget.addTopLevelItem(self.root)
        # self.setStyleSheet(helper.CommonHelper.read_css('./css/style.css'))

        # 设置列数
        self.treeWidget.setColumnCount(1)
        # 设置树形控件头部的标题
        self.treeWidget.setHeaderLabels(['财税大数据平台'])
        self.treeWidget.setFixedWidth(500)

        self.rightWidget = QWidget()
        self.rightLayout = QVBoxLayout()

        self.dataBox = QGroupBox("数据库连接")
        self.dataBox.setStyleSheet("QGroupBox{border: 1px solid black;}")
        self.connectBtn = QPushButton("生成查找树")

        self.yearLab = QLabel("年份")
        self.yearEdit = QLineEdit()

        self.dataLayout = QGridLayout()
        self.dataLayout.setContentsMargins(10, 20, 10, 15)
        self.hostLab = QLabel("主机名或ip地址：")
        self.portLab = QLabel("端口：")
        self.userLab = QLabel("用户名：")
        self.passLab = QLabel("密码")
        self.dbLab = QLabel("数据库名：")
        self.charsetLab = QLabel("字符集：")

        self.hostEdit = QLineEdit()
        self.userEdit = QLineEdit()
        self.portEdit = QLineEdit()
        self.passEdit = QLineEdit()
        self.dbEdit = QLineEdit()
        self.charsetEdit = QLineEdit()

        self.hostEdit.setText("192.168.110.201")
        self.portEdit.setText("3306")
        self.userEdit.setText("root")
        self.passEdit.setEchoMode(QLineEdit.Password)
        self.passEdit.setText("tengyun2020")
        self.dbEdit.setText("fiscal_tax")
        self.charsetEdit.setText("utf8")
        self.yearEdit.setText("2019")

        self.pageTextInfo = []

        self.dataLayout.addWidget(self.hostLab, 0, 0)
        self.dataLayout.addWidget(self.hostEdit, 0, 1)
        self.dataLayout.addWidget(self.portLab, 0, 2)
        self.dataLayout.addWidget(self.portEdit, 0, 3)
        self.dataLayout.addWidget(self.userLab, 0, 4)
        self.dataLayout.addWidget(self.userEdit, 0, 5)
        self.dataLayout.addWidget(self.yearLab, 0, 6)
        self.dataLayout.addWidget(self.yearEdit, 0, 7)
        self.dataLayout.addWidget(self.passLab, 1, 0)
        self.dataLayout.addWidget(self.passEdit, 1, 1)
        self.dataLayout.addWidget(self.dbLab, 1, 2)
        self.dataLayout.addWidget(self.dbEdit, 1, 3)
        self.dataLayout.addWidget(self.charsetLab, 1, 4)
        self.dataLayout.addWidget(self.charsetEdit, 1, 5)
        self.dataLayout.addWidget(self.connectBtn, 1, 7, 1, 1)

        self.dataBox.setLayout(self.dataLayout)

        self.budgetBox = QGroupBox("预决算信息提取")
        self.budgetBox.setStyleSheet("QGroupBox{border: 1px solid black;}")

        self.comboBox = QComboBox()
        self.comboBox.addItem("一般公共预算收支科目")
        self.comboBox.addItem("政府性基金预算收支科目")
        self.comboBox.addItem("国有资本经营预算收支科目")
        self.comboBox.addItem("社会保险基金预算收支科目")
        self.comboBox.addItem("支出经济分类科目")
        # self.comboBox.addItem("一般公共预算收入科目")
        # self.comboBox.addItem("一般公共预算支出科目")
        # self.comboBox.addItem("政府性基金预算收入科目")
        # self.comboBox.addItem("政府性基金预算支出科目")
        # self.comboBox.addItem("国有资本经营预算收入科目")
        # self.comboBox.addItem("国有资本经营预算支出科目")
        # self.comboBox.addItem("社会保险基金预算收入科目")
        # self.comboBox.addItem("社会保险基金预算支出科目")
        self.comboBox.setStyleSheet("""
                    QComboBox {border:none;background:#000000;color:#ffffff;
                    padding-left:30px;font-size:16px "SimHei";}
                    QComboBox QAbstractItemView {background:#000000;color:#ffffff;padding-left:30px;} 
                    QComboBox QAbstractItemView::item {min-height:30px;font-size:16px "SimHei";}
         """)
        self.comboBox.setView(QListView())
        self.pathLab = QLabel("文件路径：")
        self.extractBtn = QPushButton("提取信息")
        self.genBtn = QPushButton("生成Excel")
        self.initBtn = QPushButton("清空文件")
        self.codeLab = QLabel("科目编码：")
        self.subLab = QLabel("科目名称：")
        self.budgetLab = QLabel("预算数：")
        self.actualLab = QLabel("决算数：")
        self.pageLab = QLabel("页码：")
        self.pageInfoBtn = QPushButton("页面信息")
        self.regxLab = QLabel("分割符：")
        self.posLab = QLabel("名称位置：")
        self.targetLab = QLabel("目标文件路径：")
        self.boundLab = QLabel("pdf无边框表格位置：")
        self.currentPageLab = QLabel("无边框表格页码：")

        self.pathEdit = LineEdit()
        self.codeEdit = QLineEdit()
        self.subEdit = QLineEdit()
        self.budgetEdit = QLineEdit()
        self.actualEdit = QLineEdit()
        self.pageEdit = QLineEdit()
        self.regxEdit = QLineEdit()
        self.posEdit = QLineEdit()
        self.targetEdit = LineEdit()
        self.boundEdit = QLineEdit()
        self.currentPageEdit = QLineEdit()

        current_page_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"^[1-9]\d*$"), self.currentPageEdit)
        self.currentPageEdit.setValidator(current_page_validator)

        page_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[1-9]\d*)(-[1-9]\d*)||(^[1-9]\d*)(,[1-9]\d*)+"),
                                                self.pageEdit)
        self.pageEdit.setValidator(page_validator)

        sub_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[1-9]\d*)(,[1-9]\d*)+"), self.subEdit)
        self.subEdit.setValidator(sub_validator)

        code_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[1-9]\d*)(,[1-9]\d*)+"), self.codeEdit)
        self.codeEdit.setValidator(code_validator)

        budget_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[1-9]\d*)(,[1-9]\d*)+"), self.budgetEdit)
        self.budgetEdit.setValidator(budget_validator)

        actual_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[1-9]\d*)(,[1-9]\d*)+"), self.actualEdit)
        self.actualEdit.setValidator(actual_validator)

        bound_validator = QtGui.QRegExpValidator(QtCore.QRegExp(r"(^[0-9]\d*)(,[0-9]\d*)+"), self.boundEdit)
        self.boundEdit.setValidator(bound_validator)
        """
        self.pathEdit.setText("C:/Users/localhost/Desktop/bb/新乐市2019年政府预算公开附表.xlsx")
        self.targetEdit.setText("C:/Users/localhost/Desktop/cc")
        self.subEdit.setText("1")
        self.codeEdit.setText("")
        self.budgetEdit.setText("2")
        self.actualEdit.setText("")
        self.pageEdit.setText("1")
        self.boundEdit.setText("0,700,550,0")
        """
        self.checkBtn = QPushButton("查看")
        self.pathEdit.setObjectName("path")
        self.targetEdit.setObjectName("target")
        self.configEdit = QTextEdit()
        self.configEdit.setStyleSheet("""
            QTextEdit{font-size:14px;background:#3D1140;color:#FFFFFF;}
        """)
        self.configEdit.setFixedHeight(300)
        self.json_str = """
        {
            "一般公共预算收支科目":{
               "parsed":{ 
                    "page":"",
                    "code":"",
                    "budget":"",
                    "actual":"",
                    "type":""
                },
               "unparsed":{ 
                    "page":"",
                    "name":"",
                    "budget":"",
                    "actual":"",
                    "regex":{
                        "sep":"",
                        "pos":""
                    },
                    "type":""
                }
            },
            "政府性基金预算收支科目":{
            "parsed":{ 
                    "page":"",
                    "code":"",
                    "budget":"",
                    "actual":"",
                    "type":""
                },
               "unparsed":{ 
                    "page":"",
                    "name":"",
                    "budget":"",
                    "actual":"",
                    "regex":{
                        "sep":"",
                        "pos":""
                    },
                    "type":""
                }
            },
            "国有资本经营预算收支科目":{
              "parsed":{ 
                    "page":"",
                    "code":"",
                    "budget":"",
                    "actual":"",
                    "type":""
                },
               "unparsed":{ 
                    "page":"",
                    "name":"",
                    "budget":"",
                    "actual":"",
                    "regex":{
                        "sep":"",
                        "pos":""
                    },
                    "type":""
                }
            },
            "社会保险基金预算收支科目":{
               "parsed":{ 
                    "page":"",
                    "code":"",
                    "budget":"",
                    "actual":"",
                    "type":""
                },
               "unparsed":{ 
                    "page":"",
                    "name":"",
                    "budget":"",
                    "actual":"",
                    "regex":{
                        "sep":"",
                        "pos":""
                    },
                    "type":""
                }
            },
            "支出经济分类科目":{
             "parsed":{ 
                    "page":"",
                    "code":"",
                    "budget":"",
                    "actual":"",
                    "type":""
                },
               "unparsed":{ 
                    "page":"",
                    "name":"",
                    "budget":"",
                    "actual":"",
                    "regex":{
                        "sep":"",
                        "pos":""
                    },
                    "type":""
                }
            }
        }
        """
        self.configEdit.setText(json.dumps(json.loads(self.json_str), indent=4, sort_keys=False, ensure_ascii=False))
        self.json = {"一般公共预算收支科目": {}, "政府性基金预算收支科目": {}, "国有资本经营预算收支科目": {}, "社会保险基金预算收支科目": {}, "支出经济分类科目": {}}
        self.page_settings = QPushButton("")

        self.budgetLayout = QGridLayout()
        self.budgetLayout.setContentsMargins(10, 20, 10, 15)
        self.budgetLayout.addWidget(self.pathLab, 0, 0)
        self.budgetLayout.addWidget(self.pathEdit, 0, 1, 1, 3)
        self.budgetLayout.addWidget(self.targetLab, 1, 0)
        self.budgetLayout.addWidget(self.targetEdit, 1, 1, 1, 3)

        self.spiderBox = QGroupBox("文件下载")
        self.urlLab = QLabel("下载页地址：")
        self.urlEdit = QLineEdit()
        self.downloadLab = QLabel("文件下载路径：")
        self.downloadEdit = LineEdit()
        self.downloadEdit.setObjectName("download")

        self.downloadBtn = QPushButton("下载当前页文件")
        self.spiderLayout = QGridLayout()
        self.spiderLayout.addWidget(self.urlLab, 0, 0)
        self.spiderLayout.addWidget(self.urlEdit, 0, 1)
        # self.spiderLayout.addSpacing(20)
        self.spiderLayout.addWidget(self.downloadLab, 1, 0)
        self.spiderLayout.addWidget(self.downloadEdit, 1, 1)
        self.spiderLayout.addWidget(self.downloadBtn, 1, 2)
        self.spiderBox.setLayout(self.spiderLayout)

        url_validator = QtGui.QRegExpValidator(
            QtCore.QRegExp(r"http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+"),
            self.urlEdit)
        self.urlEdit.setValidator(url_validator)

        self.budgetLayout.addWidget(self.spiderBox, 3, 0, 4, 4)
        # self.budgetLayout.addWidget(self.configEdit, 2, 0, 6, 4)
        self.budgetLayout.addWidget(self.comboBox, 8, 0, 1, 4)
        self.budgetLayout.addWidget(self.subLab, 9, 0)
        self.budgetLayout.addWidget(self.subEdit, 9, 1)
        # self.budgetLayout.addWidget(self.regxLab, 3, 2)
        # self.budgetLayout.addWidget(self.regxEdit, 3, 3)
        # self.budgetLayout.addWidget(self.posLab, 4, 2)
        # self.budgetLayout.addWidget(self.posEdit, 4, 3)

        self.budgetLayout.addWidget(self.codeLab, 9, 2)
        self.budgetLayout.addWidget(self.codeEdit, 9, 3)

        self.budgetLayout.addWidget(self.budgetLab, 10, 0)
        self.budgetLayout.addWidget(self.budgetEdit, 10, 1)
        self.budgetLayout.addWidget(self.actualLab, 10, 2)
        self.budgetLayout.addWidget(self.actualEdit, 10, 3)

        self.budgetLayout.addWidget(self.pageLab, 11, 0)
        self.budgetLayout.addWidget(self.pageEdit, 11, 1)
        self.budgetLayout.addWidget(self.pageInfoBtn, 11, 4)

        self.budgetLayout.addWidget(self.boundLab, 12, 0)
        self.budgetLayout.addWidget(self.boundEdit, 12, 1)
        self.budgetLayout.addWidget(self.currentPageLab, 12, 2)
        self.budgetLayout.addWidget(self.currentPageEdit, 12, 3)
        self.budgetLayout.addWidget(self.checkBtn, 12, 4)

        self.btnWidget = QWidget()
        self.btnLayout = QHBoxLayout()
        self.btnLayout.addWidget(self.initBtn)
        self.btnLayout.addSpacing(20)
        self.btnLayout.addWidget(self.extractBtn)
        self.btnLayout.addSpacing(20)
        self.btnLayout.addWidget(self.genBtn)

        # self.testBtn = QPushButton("测试")
        # self.btnLayout.addSpacing(20)
        # self.btnLayout.addWidget(self.testBtn)

        self.btnWidget.setLayout(self.btnLayout)
        self.budgetLayout.addWidget(self.btnWidget, 13, 0, 1, 4)

        # self.budgetLayout.addWidget(self.initBtn, 10, 0)
        # self.budgetLayout.addWidget(self.extractBtn, 10, 1)
        # self.budgetLayout.addWidget(self.genBtn, 10, 3)
        for i in range(14):
            self.budgetLayout.setRowMinimumHeight(i, 25)
        self.budgetBox.setLayout(self.budgetLayout)

        self.rightLayout.addWidget(self.dataBox)
        self.rightLayout.addSpacing(20)
        self.rightLayout.addWidget(self.budgetBox)
        self.rightLayout.addStretch(1)
        self.rightWidget.setLayout(self.rightLayout)

        # 节点全部展开
        self.treeWidget.expandAll()
        self.leftWidget = QWidget()
        self.leftLayout = QHBoxLayout()
        self.leftLayout.addWidget(self.treeWidget)

        self.leftWidget.setLayout(self.leftLayout)
        self.mainLayout = QHBoxLayout()
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.mainLayout.setSpacing(0)

        self.mainLayout.addStretch(1)
        self.mainLayout.addWidget(self.leftWidget)
        self.mainLayout.addStretch(1)
        self.mainLayout.addWidget(self.rightWidget)
        self.mainLayout.addStretch(1)
        self.centralWidget.setLayout(self.mainLayout)
        self.setMinimumSize(1200, 800)
        # 查找树
        self.tree = MultiTree({"id": 0, "pid": -1, "name": "政府科目", "code": "0"})
        # 预声明子线程
        self.initThread = RunThread(self.root, self.tree, "2020")
        self.downloadThread = DownloadThread("", "")
        # 子树对应字典
        self.tree_dict = {}
        self.sheet_name_list = ["一般公共预算收支科目", "政府性基金预算收支科目", "国有资本经营预算收支科目", "社会保险基金预算收支科目", "支出经济分类科目"]

        self.connectBtn.clicked.connect(self.on_connect_clicked)
        self.pathEdit.clicked.connect(self.on_edit_double_clicked)
        self.checkBtn.clicked.connect(self.on_check_btn_clicked)

        self.targetEdit.clicked.connect(self.on_edit_double_clicked)
        self.downloadEdit.clicked.connect(self.on_edit_double_clicked)
        self.extractBtn.clicked.connect(self.on_extract_clicked)
        self.initBtn.clicked.connect(self.on_init_btn_clicked)
        self.genBtn.clicked.connect(self.on_gen_btn_clicked)
        self.downloadBtn.clicked.connect(self.on_download_btn_clicked)
        self.pageInfoBtn.clicked.connect(self.on_page_info_btn_clicked)
        # self.testBtn.clicked.connect(self.find_code_by_name)

    @pyqtSlot()
    def on_page_info_btn_clicked(self):
        if self.pageEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    页码不能为空！    ')
            return
        page = process_param(self.pageEdit.text().strip())
        if len(page) > 0:
            suffix = self.pathEdit.text().split(".")[-1]
            if suffix.lower() == "doc" or suffix.lower() == "docx" or suffix.lower() == "pdf":
                pdf = camelot.read_pdf(self.pathEdit.text().strip().replace("docx", "pdf").replace("doc", "pdf"),
                                       flavor='stream',
                                       pages=str(page[0]))
                print(len(pdf))
                if len(pdf) > 0:
                    info = ""
                    for i, row in enumerate(pdf[0].df.values.tolist()):
                        info += ",".join(row) + "\n"
                    self.pageInfo = PageInfo(page[0], info)
                    self.pageInfo.setWindowModality(QtCore.Qt.ApplicationModal)
                    self.pageInfo.show()
            elif suffix.lower() == "xls" or suffix.lower() == "xlsx":
                wb = xlrd.open_workbook(self.pathEdit.text().strip())
                sheet_names = wb.sheet_names()
                info = "\n"
                for i, sheet_name in enumerate(sheet_names):
                    info += str(i + 1) + "\t" + sheet_name + "\n"
                self.pageInfo = PageInfo(page[0], info)
                self.pageInfo.setWindowModality(QtCore.Qt.ApplicationModal)
                self.pageInfo.show()
            else:
                pass

    @pyqtSlot()
    def on_download_btn_clicked(self):
        if self.urlEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    url地址不能为空！    ')
            return
        if self.downloadEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    文件夹地址不能为空！    ')
            return

        self.downloadThread = DownloadThread(url=self.urlEdit.text().strip(), path=self.downloadEdit.text().strip())
        self.downloadThread.resSignal.connect(lambda _msg: self.on_download_thread(_msg))
        self.downloadThread.start()

        # downloader = DownLoader(timeout=30, url=self.urlEdit.text().strip(), path=self.downloadEdit.text().strip())
        # downloader.download_file()
        # QMessageBox.information(self, "提示", '    文件下载完成！    ')
        # pass

    # @staticmethod
    # def check_excel_is_open(path):
    #     file_name = path.split("\\")[-1]
    #     print(file_name)
    #     print(path)
    #     print(path.replace(file_name, '~$' + file_name))
    #     if os.path.exists(path.replace(file_name, '~$' + file_name)):
    #         print("True")
    #         return True
    #     print("False")
    #     return False

    @pyqtSlot()
    def on_gen_btn_clicked(self):
        if len(self.json) > 0:
            # print(os.path.join(self.targetEdit.text().strip(), "预决算.xls"))
            # if self.check_excel_is_open(os.path.join(self.targetEdit.text().strip().replace("/", "\\"), "预决算.xls")):
            #     QMessageBox.information(self, "提示", '    Excel文件已经打开，请先关闭！    ')
            #     return
            excel_writer = Excel(os.path.join(self.targetEdit.text().strip().replace("/", "\\"), "预决算.xls"),
                                 self.sheet_name_list, self.json)
            try:
                excel_writer.write_excel()
            except:
                QMessageBox.information(self, "提示", '    写入失败查看是否文件已经打开！    ')
                return

            QMessageBox.information(self, "提示", '    Json信息写入Excel成功！    ')

    @pyqtSlot()
    def on_init_btn_clicked(self):
        try:
            if self.targetEdit.text().strip() == "" or not os.path.isdir(self.targetEdit.text().strip()):
                QMessageBox.information(self, "提示", '    输入不能为空，或者路径有错误！    ')
            else:
                self.json.clear()
                for file in os.listdir(self.targetEdit.text().strip().replace("/", "\\")):
                    # print(os.path.join(self.targetEdit.text().strip().replace("/", "\\"), file))
                    if os.path.isfile(os.path.join(self.targetEdit.text().strip().replace("/", "\\"), file)):
                        try:
                            os.remove(os.path.join(self.targetEdit.text().strip().replace("/", "\\"), file))
                        except:
                            QMessageBox.information(self, "提示", '    删除文件失败请查看是否打开！    ')
                            return
                QMessageBox.information(self, "提示", '    清空文件成功！    ')
        except Exception as e:
            QMessageBox.information(self, "提示", e)

    @pyqtSlot()
    def on_check_btn_clicked(self):
        # self.find_code_by_name()
        if self.currentPageEdit.text().strip() == "" or self.pathEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    输入不能为空！    ')
            return
        elif not self.pathEdit.text().strip().endswith(".pdf"):
            QMessageBox.information(self, "提示", '    只有PDF文件需要次操作！    ')
            return
        else:
            print(self.pathEdit.text().strip().replace('.docx', '.pdf').replace(".doc", '.pdf'))
            pdf = camelot.read_pdf(self.pathEdit.text().strip().replace('.docx', '.pdf').replace(".doc", '.pdf'),
                                   flavor='stream',
                                   pages=self.currentPageEdit.text().strip())
            if pdf:
                plt = camelot.plot(pdf[0], kind='textedge')
                plt.show()
                axis('tight')
                fig = pylab.gcf()
                fig.canvas.set_window_title("第" + self.currentPageEdit.text().strip() + "页表格解析示意图")

    @pyqtSlot()
    def on_extract_clicked(self):
        # print(self.tree_dict)
        # self.find_code("")
        #
        if len(self.tree_dict) == 0:
            QMessageBox.information(self, "提示", '    未生成查找树！    ')
            return
        if self.pathEdit.text().strip() == "" or self.targetEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    文件路径和目标路径不能为空！    ')
            return
        if self.pageEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    页码不能为空！    ')
            return
        if self.budgetEdit.text().strip() == "" and self.actualEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    预算数、决算数不能同时为空！    ')
            return
        if self.subEdit.text().strip() == "" and self.codeEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    科目名称和科目编码不能同时为空！    ')
            return
        if self.subEdit.text().strip() != "" and self.codeEdit.text().strip() != "":
            QMessageBox.information(self, "提示", '    科目名称和科目编码不能同时非空！    ')
            return
        try:
            process_file(self.pathEdit.text().strip())
            suffix = self.pathEdit.text().split(".")[-1]
            # print(self.pathEdit.text())
            page = process_param(self.pageEdit.text().strip())
            if suffix.lower() == "doc" or suffix.lower() == "docx" or suffix.lower() == "pdf":
                bound_info = self.boundEdit.text().strip()
                print(bound_info)
                if bound_info != "":
                    if bound_info.endswith(","):
                        bound_info = bound_info.rstrip(",")
                    # print(page)
                    if len(page) > 0:
                        # print((list(map(str, page))))
                        pdf = camelot.read_pdf(self.pathEdit.text().strip(), flavor='stream',
                                               pages=','.join(list(map(str, page))),
                                               table_areas=[bound_info])
                        print(len(pdf))
                        for i in range(len(pdf)):
                            table_list = []
                            for row_data in pdf[i].df.values.tolist():
                                table_list.append(row_data)
                            self.parse(table_list)
                        QMessageBox.information(self, "提示", '    提取信息结束！    ')
                        print(self.json)
                else:
                    """
                    pdf = pdfplumber.open(self.pathEdit.text().strip())
                    # print(pdf.pages)
                    # print(len(pdf.pages))
                    # print(page)
                    for i, _page in enumerate(pdf.pages):
                        if i + 1 in page:
                            table_list = []
                            print(_page.extract_text())
                            for pdf_table in _page.extract_tables():
                                for row in pdf_table:
                                    table_list.append(row)
                                    print(row)
                            print(table_list)
                            self.parse(table_list)
                    """
                    if len(page) > 0:
                        # print((list(map(str, page))))
                        if suffix.lower() == "doc" or suffix.lower() == "docx":
                            pdf = camelot.read_pdf(self.pathEdit.text().strip().replace("." + suffix, ".pdf"),
                                                   flavor='stream',
                                                   pages=','.join(list(map(str, page))))
                        else:
                            pdf = camelot.read_pdf(self.pathEdit.text().strip(), flavor='stream',
                                                   pages=','.join(list(map(str, page))))
                        # print(len(pdf))
                        for i in range(len(pdf)):
                            table_list = []
                            for row_data in pdf[i].df.values.tolist():
                                table_list.append(row_data)
                            self.parse(table_list)
                    QMessageBox.information(self, "提示", '    提取信息结束！    ')

            elif suffix.lower() == "xls" or suffix.lower() == "xlsx":
                wb = xlrd.open_workbook(self.pathEdit.text().strip())
                sheet_names = wb.sheet_names()
                for i, sheet_name in enumerate(sheet_names):
                    # print(sheet_name)
                    if i + 1 in page:
                        table = wb.sheet_by_index(i)
                        table_list = []
                        print(table.nrows)

                        for ii in range(table.nrows):
                            # print(type(table.row_values(ii)))
                            # print(table.row_values(ii))
                            table_list.append(table.row_values(ii))
                        print(table_list)
                        self.parse(table_list)
                QMessageBox.information(self, "提示", '    提取信息结束！    ')
        except:
            QMessageBox.information(self, "提示", "提取信息失败请查看输入是否有误！")

    @pyqtSlot()
    def find_code_by_name(self):
        """
        Test of new method for finding the code
        :param name_list:
        :return:
        """
        pdf = camelot.read_pdf(r'C:\Users\localhost\Desktop\政府性基金预算支出表.pdf', flavor='stream',
                               pages='1')
        for i in range(len(pdf)):
            name_list = []
            for row_data in pdf[i].df.values.tolist():
                name = re.sub(r'\s+', '', row_data[0])
                name = name.split("、")[-1]
                name = name.split("：")[-1]
                if name.strip() != "":
                    name_list.append(name.strip())
            self.tree.prepare_search_name(self.tree_dict.get("政府性基金预算收支科目"))
            res = []
            for name in name_list:
                res.append(self.tree.search_node_by_name(name, 0.8))
            for item in res:
                print(item)

        pass

    def backtracking(self, index, data_list, res=[]):
        if index == len(data_list) - 1:
            return
        if len(data_list[index]) == 0:
            return
        for data in data_list:
            res_temp = res[:]
            if len(res_temp) == 0:
                pass
            for item in res_temp:
                if data.get("pid") == item.get("id"):
                    pass

        pass

    def find_code(self, name_list):
        """
        print(self.comboBox.currentText())
        print(self.tree_dict.get(self.comboBox.currentText()))
        name_list = ["一般公共服务支出", "人大事务", "行政运行", "政协事务", "行政运行", "机关服务", "教育支出", "教育管理事务", "行政运行"]
        name_list = ["我拉个区", "一般公共服务支出", "援助其他地区支出", "垃圾", "一般公共服务", "国防支出", "公共安全支出", "教育支出"]
        name_list = ["一般公共服务支出", "援助其他地区支出", "一般公共服务", "国防支出", "公共安全支出", "教育支出"]
        """
        try:
            self.tree.prepare_search_name(self.tree_dict.get(self.comboBox.currentText()))
            res = []
            mark = ""
            for i, name in enumerate(name_list):
                search_res = self.tree.search_name(name)
                # print(search_res)
                # res.append(search_res)
                if len(search_res) == 1:
                    mark = search_res[0]
                    # print(i, search_res[0])
                    res.append((i, mark))
                elif len(search_res) > 1:
                    search_res.sort(key=lambda j: len(j))
                    for search_item in search_res:
                        if search_item.startswith(mark):
                            # print(i, search_item)
                            mark = search_item
                            res.append((i, mark))
                            break
                        elif len(search_item) == len(mark):
                            # print(i, search_item)
                            mark = search_item
                            res.append((i, mark))
                            break
            return res
        except Exception as e:
            QMessageBox.information(self, "提示", e)

    def parse(self, data_list):
        # print(data_list)
        try:
            if len(data_list) == 0:
                QMessageBox.information(self, "提示", '    表格解析失败！    ')
                return
            budget_num = process_param(self.budgetEdit.text().strip())
            actual_num = process_param(self.actualEdit.text().strip())
            if self.codeEdit.text().strip() != "":
                code_num = process_param(self.codeEdit.text().strip())
                if (len(budget_num) > 0 and len(budget_num) != len(code_num)) or (
                        len(actual_num) > 0 and len(actual_num) != len(code_num)):
                    QMessageBox.information(self, "提示", '    长度不对应！    ')
                    return
                if (len(budget_num) > 0 and budget_num[-1] > len(data_list[0])) or (
                        len(actual_num) > 0 and actual_num[-1] > len(data_list[0])) or code_num[-1] > len(
                    data_list[0]):
                    QMessageBox.information(self, "提示", '    列数越界！    ')
                    return
                for data in data_list:
                    for i in range(len(code_num)):
                        key = re.sub(r'\s+', '', str(data[code_num[i] - 1]))
                        # print(key)
                        if isinstance(data[code_num[i] - 1], float):
                            key = str(int(data[code_num[i] - 1]))
                        if key and key.isdigit():
                            if len(budget_num) > 0:
                                if self.json.get(self.comboBox.currentText().strip()).get(key):
                                    self.json.get(self.comboBox.currentText().strip()).get(key).update(
                                        {"预算数": data[budget_num[i] - 1]})
                                else:
                                    self.json.get(self.comboBox.currentText().strip()).update(
                                        {key: {"预算数": data[budget_num[i] - 1]}})
                            if len(actual_num) > 0:
                                if self.json.get(self.comboBox.currentText().strip()).get(key):
                                    self.json.get(self.comboBox.currentText().strip()).get(key).update(
                                        {"决算数": data[actual_num[i] - 1]})
                                else:
                                    self.json.get(self.comboBox.currentText().strip()).update(
                                        {key: {"决算数": data[actual_num[i] - 1]}})
            else:
                sub_num = process_param(self.subEdit.text().strip())
                if (len(budget_num) > 0 and len(budget_num) != len(sub_num)) or (
                        len(actual_num) > 0 and len(actual_num) != len(sub_num)):
                    QMessageBox.information(self, "提示", '    长度不对应！    ')
                    return
                if (len(budget_num) > 0 and budget_num[-1] > len(data_list[0])) or (
                        len(actual_num) > 0 and actual_num[-1] > len(data_list[0])) or sub_num[-1] > len(
                    data_list[0]):
                    QMessageBox.information(self, "提示", '    列数越界！    ')
                    return
                name_list = []
                for i in range(len(data_list)):
                    row_name = []
                    for j in range(len(sub_num)):
                        name = re.sub(r'\s+', '', data_list[i][sub_num[j] - 1])
                        name = name.split("、")[-1]
                        name = name.split("：")[-1]
                        row_name.append(name)
                    name_list.append(row_name)
                name_array = np.array(name_list)
                for j in range(len(sub_num)):
                    for index_code in self.find_code(name_array[:, j].tolist()):
                        key = index_code[1]
                        if key.isdigit():
                            if len(budget_num) > 0:
                                if self.json.get(self.comboBox.currentText().strip()).get(key):
                                    self.json.get(self.comboBox.currentText().strip()).get(key).update(
                                        {"预算数": data_list[index_code[0]][budget_num[j] - 1]})
                                else:
                                    self.json.get(self.comboBox.currentText().strip()).update(
                                        {key: {"预算数": data_list[index_code[0]][budget_num[j] - 1]}})
                            if len(actual_num) > 0:
                                print("------------>", actual_num[j] - 1, type(actual_num[j]))
                                print("___________", data_list[index_code[0]])
                                if self.json.get(self.comboBox.currentText().strip()).get(key):
                                    self.json.get(self.comboBox.currentText().strip()).get(key).update(
                                        {"决算数": data_list[index_code[0]][actual_num[j] - 1]})
                                else:
                                    self.json.get(self.comboBox.currentText().strip()).update(
                                        {key: {"决算数": data_list[index_code[0]][actual_num[j] - 1]}})
        except Exception as e:
            QMessageBox.information(self, "提示", e)

    # def check_input(self, num, budget, actual, data):
    #     if (len(budget) > 0 and len(budget) != len(num)) or (
    #             len(actual) > 0 and len(actual) != len(num)):
    #         QMessageBox.information(self, "提示", '    长度不对应！    ')
    #         return
    #     if budget[-1] >= len(data[0]) or actual[-1] >= len(data[0]) or num[-1] >= len(
    #             data[0]):
    #         QMessageBox.information(self, "提示", '    列数越界！    ')
    #         return

    @pyqtSlot()
    def on_edit_double_clicked(self):
        if self.sender().objectName() == "path":
            filepath = QFileDialog.getOpenFileName(self,
                                                   "请选择文件路径", "/")
            if filepath:
                self.pathEdit.setText(filepath[0].strip())
        elif self.sender().objectName() == "target":
            directory = QFileDialog.getExistingDirectory(self, "请选择文件夹路径", "/")
            if directory:
                self.targetEdit.setText(directory.strip())
        elif self.sender().objectName() == "download":
            directory = QFileDialog.getExistingDirectory(self, "请选择文件夹路径", "/")
            if directory:
                self.downloadEdit.setText(directory.strip())
        else:
            pass

    @pyqtSlot()
    def on_connect_clicked(self):
        print("connect button is clicked!")
        if self.hostEdit.text().strip() == "" or self.portEdit.text().strip() == "" or self.userEdit.text().strip() == "" or self.yearEdit.text().strip() == "" or self.passEdit.text().strip() == "" or self.dbEdit.text().strip() == "" or self.charsetEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    输入不能为空！    ')

        else:
            year = self.yearEdit.text().strip()
            self.tree = MultiTree({"id": 0, "pid": -1, "name": year + "年政府科目", "code": "0"})

            self.treeWidget.clear()
            self.root = QTreeWidgetItem(self.treeWidget)
            self.root.setText(0, year + "年树状结构生成中...")
            self.treeWidget.addTopLevelItem(self.root)

            self.initThread = RunThread(self.root, self.tree, year)
            self.initThread.resSignal.connect(lambda _year, _tree: self.on_init_thread(_year, _tree))
            self.initThread.start()

    @pyqtSlot()
    def on_download_thread(self, msg):
        if msg == "ok":
            QMessageBox.information(self, "提示", '    文件下载成功！    ')
        elif msg == "fail":
            QMessageBox.information(self, "提示", '    文件下载失败！    ')
        else:
            pass
            # QMessageBox.information(self, "提示", msg)

    @pyqtSlot()
    def on_init_thread(self, year, tree):
        self.tree = tree
        self.root.setText(0, year + "年树状结构")
        # 遍历树状结构
        # self.tree.traverse(self.tree)
        self.tree_dict.clear()
        self.generate_tree_dict()
        self.treeWidget.expandAll()

    def generate_tree_dict(self):
        self.tree_dict.update({
            "一般公共预算收入": self.tree.tree.children[0].children[0],
            "一般公共预算支出": self.tree.tree.children[0].children[1],
            "政府性基金预算收入": self.tree.tree.children[1].children[0],
            "政府性基金预算支出": self.tree.tree.children[1].children[1],
            "国有资本经营预算收入": self.tree.tree.children[2].children[0],
            "国有资本经营预算支出": self.tree.tree.children[2].children[1],
            "社会保险基金预算收入": self.tree.tree.children[3].children[0],
            "社会保险基金预算支出": self.tree.tree.children[3].children[1],
            "支出经济分类科目": self.tree.tree.children[4],
            "一般公共预算收支科目": self.tree.tree.children[0],
            "政府性基金预算收支科目": self.tree.tree.children[1],
            "国有资本经营预算收支科目": self.tree.tree.children[2],
            "社会保险基金预算收支科目": self.tree.tree.children[3]
        })


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())

# -*- coding:utf8 -*-
import json
import os
import sys
from os.path import join

import matplotlib
from matplotlib import pylab
from matplotlib.pyplot import axis

matplotlib.use("Qt5Agg")
import camelot
import pymysql
from PyQt5 import QtGui, QtCore
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QMainWindow, QTreeWidget, QWidget, QTreeWidgetItem, QHBoxLayout, QApplication, QLabel, \
    QLineEdit, QVBoxLayout, QGridLayout, QGroupBox, QPushButton, QMessageBox, QFileDialog, \
    QComboBox, QListView, QTextEdit

from lineedit import LineEdit
from node import MultiTree, TreeNode


class RunThread(QtCore.QThread):
    resSignal = QtCore.pyqtSignal(str, MultiTree)

    def __init__(self, root, tree, year):
        super(RunThread, self).__init__()
        self.root = root
        self.tree = tree
        self.year = year

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
            QMessageBox.information(self, "提示", '    数据库连接失败！    ')
            print(e)

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
            conn.rollback()
            QMessageBox.information(self, "提示", '    数据库查询失败！    ')
        try:
            # 关闭连接释放资源，先关闭游标
            cursor.close()
            # 再关闭连接对象
            conn.close()
        except Exception as e:
            QMessageBox.information(self, "提示", '    关闭数据库失败！    ')

    def run(self):
        print("Access the thread!")
        self.generate_tree_by_year(self.year)
        self.resSignal.emit(self.year, self.tree)  # 任务完成后，发送信号


class ExcelThread(QtCore.QThread):
    resSignal = QtCore.pyqtSignal(str)

    def __init__(self):
        super(RunThread, self).__init__()

    def __del__(self):
        self.wait()

    def run(self):
        print("")


class MainWindow(QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__()
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

        self.page_settings = QPushButton("")

        self.budgetLayout = QGridLayout()
        self.budgetLayout.setContentsMargins(10, 20, 10, 15)
        self.budgetLayout.addWidget(self.pathLab, 0, 0)
        self.budgetLayout.addWidget(self.pathEdit, 0, 1, 1, 3)
        self.budgetLayout.addWidget(self.targetLab, 1, 0)
        self.budgetLayout.addWidget(self.targetEdit, 1, 1, 1, 3)

        # self.budgetLayout.addWidget(self.configEdit, 2, 0, 6, 4)
        self.budgetLayout.addWidget(self.comboBox, 2, 0, 1, 4)
        self.budgetLayout.addWidget(self.subLab, 3, 0)
        self.budgetLayout.addWidget(self.subEdit, 3, 1)
        # self.budgetLayout.addWidget(self.regxLab, 3, 2)
        # self.budgetLayout.addWidget(self.regxEdit, 3, 3)
        # self.budgetLayout.addWidget(self.posLab, 4, 2)
        # self.budgetLayout.addWidget(self.posEdit, 4, 3)

        self.budgetLayout.addWidget(self.codeLab, 3, 2)
        self.budgetLayout.addWidget(self.codeEdit, 3, 3)

        self.budgetLayout.addWidget(self.budgetLab, 4, 0)
        self.budgetLayout.addWidget(self.budgetEdit, 4, 1)
        self.budgetLayout.addWidget(self.actualLab, 4, 2)
        self.budgetLayout.addWidget(self.actualEdit, 4, 3)

        self.budgetLayout.addWidget(self.pageLab, 7, 0)
        self.budgetLayout.addWidget(self.pageEdit, 7, 1)

        self.budgetLayout.addWidget(self.boundLab, 8, 0)
        self.budgetLayout.addWidget(self.boundEdit, 8, 1)
        self.budgetLayout.addWidget(self.currentPageLab, 8, 2)
        self.budgetLayout.addWidget(self.currentPageEdit, 8, 3)
        self.budgetLayout.addWidget(self.checkBtn, 8, 4)

        self.budgetLayout.addWidget(self.initBtn, 10, 0)
        self.budgetLayout.addWidget(self.extractBtn, 10, 1)
        self.budgetLayout.addWidget(self.genBtn, 10, 3)

        # for i in range(11):
        #     self.budgetLayout.setRowMinimumHeight(i, 25)
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
        # 子树对应字典
        self.tree_dict = {}

        self.connectBtn.clicked.connect(self.on_connect_clicked)
        self.pathEdit.clicked.connect(self.on_edit_double_clicked)
        self.checkBtn.clicked.connect(self.on_check_btn_clicked)

        self.targetEdit.clicked.connect(self.on_edit_double_clicked)
        self.extractBtn.clicked.connect(self.on_extract_clicked)

    @pyqtSlot()
    def on_init_btn_clicked(self):
        if self.targetEdit.text().strip() == "" or not os.path.isdir(self.targetEdit.text().strip()):
            print()
        else:
            for file in os.listdir(self.self.targetEdit.text().strip()):
                if os.path.isfile(join(self.targetEdit.text(), file)):
                    os.remove(join(self.targetEdit.text(), file))

    @pyqtSlot()
    def on_check_btn_clicked(self):
        if self.currentPageEdit.text().strip() == "" or self.pathEdit.text().strip() == "":
            QMessageBox.information(self, "提示", '    输入不能为空！    ')
        elif not self.pathEdit.text().strip().endswith(".pdf"):
            QMessageBox.information(self, "提示", '    只有PDF文件需要次操作！    ')
        else:
            pdf = camelot.read_pdf(self.pathEdit.text().strip(), flavor='stream',
                                   pages=self.currentPageEdit.text().strip(), table_areas=['0,700,550,50'])
            if pdf:
                plt = camelot.plot(pdf[0], kind='textedge')
                print(type(plt))
                plt.show()
                axis('tight')
                fig = pylab.gcf()
                fig.canvas.set_window_title("第" + self.currentPageEdit.text().strip() + "页表格解析示意图")

    @pyqtSlot()
    def on_extract_clicked(self):
        self.tree.prepare_search_name(self.tree)
        # name_list = ["一般公共服务支出", "人大事务", "行政运行", "政协事务", "行政运行", "机关服务", "教育支出", "教育管理事务", "行政运行"]
        name_list = ["一般公共服务支出", "援助其他地区支出", "一般公共服务", "国防支出", "公共安全支出", "教育支出"]
        res = []
        mark = ""
        for i, name in enumerate(name_list):
            search_res = self.tree.search_name(name)
            # print(search_res)
            # res.append(search_res)
            if len(search_res) == 1:
                mark = search_res[0]
                print(i, search_res[0])
                res.append((i, mark))
            elif len(search_res) > 1:
                search_res.sort(key=lambda j: len(j))
                for search_item in search_res:
                    if search_item.startswith(mark):
                        print(i, search_item)
                        mark = search_item
                        res.append((i, mark))
                        break
                    elif len(search_item) == len(mark):
                        print(i, search_item)
                        mark = search_item
                        res.append((i, mark))
                        break
        print('_' * 120)
        print(res)

    @pyqtSlot()
    def on_edit_double_clicked(self):
        if self.sender().objectName() == "path":
            filepath = QFileDialog.getOpenFileName(self,
                                                   "请选择文件路径", "/")
            if filepath:
                self.pathEdit.setText(filepath[0].strip())
        else:
            directory = QFileDialog.getExistingDirectory(self, "请选择文件夹路径", "/")
            if directory:
                self.targetEdit.setText(directory.strip())

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
            "支出经济分类": self.tree.tree.children[4],
            "一般公共预算收支": self.tree.tree.children[0],
            "政府性基金预算收支": self.tree.tree.children[1],
            "国有资本经营预算收支": self.tree.tree.children[2],
            "社会保险基金预算收支": self.tree.tree.children[3]
        })


if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())

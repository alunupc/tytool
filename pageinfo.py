from PyQt5 import QtGui, QtCore
from PyQt5.QtWidgets import QDialog, QLabel, QHBoxLayout, QTextEdit


class PageInfo(QDialog):
    def __init__(self, page, info):
        super().__init__()
        self.icon = QtGui.QIcon()
        self.icon.addPixmap(QtGui.QPixmap("images/logo.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.setWindowIcon(self.icon)
        self.setWindowTitle(f"第{1}页面信息".format(page))
        self.setContentsMargins(0, 0, 0, 0)
        self.setFixedSize(600, 400)
        self.mainLayout = QHBoxLayout()
        self.mainLayout.setContentsMargins(0, 0, 0, 0)
        self.infoText = QTextEdit()
        self.infoText.setText(info)
        self.mainLayout.addWidget(self.infoText)
        self.setLayout(self.mainLayout)
        self.infoText.setAlignment(QtCore.Qt.AlignVCenter | QtCore.Qt.AlignHCenter)

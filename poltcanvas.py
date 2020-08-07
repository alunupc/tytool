# -*- coding: utf-8 -*-
import sys

import camelot
import matplotlib
import matplotlib.font_manager as fm
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.pyplot import axis

matplotlib.use("Qt5Agg")  # 声明使用QT5
myfont = fm.FontProperties(fname='C:/Windows/Fonts/msyh.ttc')


# 通过继承FigureCanvas类，使得该类既是一个PyQt5的Qwidget，又是一个matplotlib的FigureCanvas，这是连接pyqt5与matplot lib的关键
class Figure_Canvas(FigureCanvas):

    def __init__(self, parent=None, width=8, height=8, dpi=100):
        fig = Figure(figsize=(width, height),
                     dpi=200)  # 创建一个Figure，注意：该Figure为matplotlib下的figure，不是matplotlib.pyplot下面的figure
        pdf = camelot.read_pdf(r"C:\Users\localhost\Desktop\石家庄市2018年市本级和全市财政总决算报表.pdf", flavor='stream', pages='5')
        if pdf:
            fig = camelot.plot(pdf[0], kind='textedge')
        fig.set_dpi(150)
        axis('tight')
        FigureCanvas.__init__(self, fig)  # 初始化父类
        self.setParent(parent)


class Mytest(QMainWindow):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # 设置窗口标题
        self.setWindowTitle('Demo for matplotlib in GUI')
        self.setFixedSize(1000, 800)

        # ===通过graphicview来显示图形
        self.graphicview = QtWidgets.QGraphicsView()  # 第一步，创建一个QGraphicsView
        self.graphicview.setObjectName("graphicview")

        dr = Figure_Canvas()
        # 第三步，创建一个QGraphicsScene，因为加载的图形（FigureCanvas）不能直接放到graphicview控件中，必须先放到graphicScene，然后再把graphicscene放到graphicview中
        graphicscene = QtWidgets.QGraphicsScene()
        graphicscene.addWidget(dr)  # 第四步，把图形放到QGraphicsScene中，注意：图形是作为一个QWidget放到QGraphicsScene中的
        self.graphicview.setScene(graphicscene)  # 第五步，把QGraphicsScene放入QGraphicsView
        self.graphicview.show()  # 最后，调用show方法呈现图形！Voila!!
        self.setCentralWidget(self.graphicview)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    test = Mytest()
    test.show()
    app.exec_()

# -*- coding: utf-8 -*-
import sys

from PyQt5.QtCore import pyqtSlot
from PyQt5.QtWidgets import QWidget, QMessageBox, \
    QApplication, QLabel, QLineEdit, QGridLayout, QPushButton
import readexcel


class Example(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):

        path = QLabel('excel文件路径:', self)

        self.pathbox = QLineEdit()
        grid = QGridLayout()
        grid.setSpacing(10)
        grid.addWidget(path, 1, 0)
        grid.addWidget(self.pathbox, 1, 1)
        btn = QPushButton('运行程序', self)
        btn.clicked.connect(self.on_click)
        grid.addWidget(btn, 2, 1)
        self.setLayout(grid)
        self.setGeometry(300, 300, 350, 300)
        self.setWindowTitle('小程序')
        self.show()

    @pyqtSlot()
    def on_click(self):
        pathName = self.pathbox.text()
        print(pathName)
        flag = readexcel.read_excel(pathName)
        # if flag == 1:


    def closeEvent(self, event):

        reply = QMessageBox.question(self, 'Message', 'Are you sure to quit?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()



if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Example()
    sys.exit(app.exec_())

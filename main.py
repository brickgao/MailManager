# -*- coding: utf-8 -*-

import sys
from con_db import con_db
from PyQt4 import QtGui, QtCore

class QMainArea(QtGui.QWidget):

    def __init__(self):

        super(QMainArea, self).__init__()
        self.initLayout()

    def initLayout(self):

        grid = QtGui.QGridLayout()
        grid.setSpacing(5)


        self.AddressList = QtGui.QListWidget()

        self.MailList = QtGui.QTreeWidget()
        self.MailList.setHeaderLabels([u'时间', u'收件人',
                                       u'发件人', u'标题', u'预览'])

        grid.addWidget(self.MailList)


        self.setLayout(grid) 

class MainWindow(QtGui.QMainWindow):

    def __init__(self):

        super(MainWindow, self).__init__()
        self.initLayout()

    def initLayout(self):

        self.statusBar()

        menubar = self.menuBar()

        exitAction = QtGui.QAction(u'退出', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip(u'退出 Mail Manager')
        exitAction.triggered.connect(QtGui.qApp.quit)

        outputAction = QtGui.QAction(u'批量导出', self)
        outputAction.setStatusTip(u'批量导出文件')

        inputAction = QtGui.QAction(u'打开', self)
        inputAction.setStatusTip(u'打开文件')
        inputAction.triggered.connect(self.inputfile)
    
        fileMenu = menubar.addMenu(u'文件')
        fileMenu.addAction(inputAction)
        fileMenu.addAction(outputAction)
        fileMenu.addAction(exitAction)

        self.MainArea = QMainArea()

        self.setCentralWidget(self.MainArea)
        
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle(u'Mail Manager')
        self.show()

    def inputfile(self):

        fname = QtGui.QFileDialog.getOpenFileName(self, u'打开文件')
        self.db = con_db(fname)

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
    


if __name__ == '__main__':
    main()

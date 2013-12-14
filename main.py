# -*- coding: utf-8 -*-

import sys
from PyQt4 import QtGui

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
    
        fileMenu = menubar.addMenu(u'文件')
        fileMenu.addAction(inputAction)
        fileMenu.addAction(outputAction)
        fileMenu.addAction(exitAction)
        
        self.setGeometry(100, 100, 800, 600)
        self.setWindowTitle(u'Mail Manager')
        self.show()


def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
    


if __name__ == '__main__':
    main()

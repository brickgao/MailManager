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

        self.MailList = QtGui.QTreeWidget()
        self.MailList.setHeaderLabels([u'时间', u'收件人',
                                       u'发件人', u'标题', u'预览'])
        self.MailList.setColumnWidth(0, 120)
        self.MailList.setColumnWidth(1, 200)
        self.MailList.setColumnWidth(2, 200)
        self.MailList.setColumnWidth(3, 100)

        grid.addWidget(self.MailList)

        self.setLayout(grid) 

    def updateInfo(self, list_db):

        self.MailList.clear() 
        for i in range(len(list_db)):

            MailInfo = QtGui.QTreeWidgetItem()
            receive_date = list_db[i]['receive_date']
            date_s = str(receive_date.year) + '-' + \
                     str(receive_date.month) + '-' + \
                     str(receive_date.day) + ' ' + \
                     str(receive_date.hour) + ':' + \
                     str(receive_date.second)
            MailInfo.setText(0, date_s)
            MailInfo.setText(1, list_db[i]['to_address'])
            MailInfo.setText(2, list_db[i]['from_address'])
            MailInfo.setText(3, list_db[i]['subject'])
            MailInfo.setText(4, list_db[i]['snippet'])
            self.MailList.addTopLevelItem(MailInfo)


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
        sfname = str(fname)
        if sfname == '':
            return
        try:
            self.db = con_db(sfname)
        except Exception, e:
            print e
            self.ErrorAlert(u'打开 db 错误，请选择正确的 db 文件')
        else:
            try:
                self.list_db = self.db.get_mail_list()
            except Exception, e:
                print e
                self.ErrorAlert(u'解析 db 错误，请选择正确的 db 文件')
            else:
                self.MainArea.updateInfo(self.list_db)
            

    def ErrorAlert(self, s):  
        QtGui.QMessageBox.critical(self, u'错误', s)  
        
    
            

def main():
    
    app = QtGui.QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())
    


if __name__ == '__main__':
    main()

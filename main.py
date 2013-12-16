# -*- coding: utf-8 -*-

import sys, os, eml_gen
from con_db import con_db
from PyQt4 import QtGui, QtCore, QtWebKit

class QMailView(QtGui.QDialog):

    def __init__(self, MailMsg, parent = None):
        
        self.MailMsg = MailMsg
        QtGui.QDialog.__init__(self, parent)
        self.initLayout()

    def initLayout(self):
        
        from_address = QtGui.QLabel(u'发信人：')
        from_address_val = QtGui.QLabel(self.MailMsg['from_address'])

        to_address = QtGui.QLabel(u'收信人：')
        to_address_val = QtGui.QLabel(self.MailMsg['to_address'])

        subject = QtGui.QLabel(u'标题：')
        subject_val = QtGui.QLabel(self.MailMsg['subject'])

        receive_date = QtGui.QLabel(u'收信时间：')
        receive_date_val = self.MailMsg['receive_date']
        date_s = QtGui.QLabel(
                    str(receive_date_val.year) + '-' + \
                    str(receive_date_val.month) + '-' + \
                    str(receive_date_val.day) + ' ' + \
                    str(receive_date_val.hour) + ':' + \
                    str(receive_date_val.second))

        ContentView = QtGui.QLabel(u'邮件正文：')
        ContentView_val = QtWebKit.QWebView()
        HtmlCode = unicode(self.MailMsg['body'], 'utf-8')
        ContentView_val.setHtml(HtmlCode, 
                            QtCore.QUrl(''))

        attachmentsList = QtGui.QLabel(u'附件：')
        attachmentsList_val = QtGui.QTreeWidget()
        attachmentsList_val.setHeaderLabels([u'#' ,u'文件名', 
                                             u'文件大小', u'是否存在',
                                             u'文件路径'])
        attachmentsList_val.setColumnWidth(0, 60)
        for i in range(len(self.MailMsg['attachments'])):
            FileInfo = QtGui.QTreeWidgetItem()
            FileInfo.setText(0, str(i))
            FileInfo.setText(1, self.MailMsg['attachments'][i]['name'])
            FileInfo.setText(2, self.MailMsg['attachments'][i]['format_size'])
            FileInfo.setText(3, str(self.MailMsg['attachments'][i]['exist']))
            if self.MailMsg['attachments'][i]['exist'] == True:
                filepath = self.MailMsg['attachments'][i]['path']
                FileInfo.setText(4, filepath)
            else:
                FileInfo.setText(4, u'文件不存在')
            attachmentsList_val.addTopLevelItem(FileInfo)

        MailoutputBtn = QtGui.QPushButton(u'导出当前文件')
        MailoutputBtn.clicked.connect(self.Mailoutput)

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(from_address, 1, 0)
        grid.addWidget(from_address_val, 1, 1)
        grid.addWidget(to_address, 2, 0)
        grid.addWidget(to_address_val, 2, 1)
        grid.addWidget(subject, 3, 0)
        grid.addWidget(subject_val, 3, 1)
        grid.addWidget(receive_date, 4, 0)
        grid.addWidget(date_s, 4, 1)
        grid.addWidget(attachmentsList, 5, 0)
        grid.addWidget(attachmentsList_val, 6, 0, 1, 3)
        grid.addWidget(ContentView, 7, 0)
        grid.addWidget(MailoutputBtn, 7, 2)
        grid.addWidget(ContentView_val, 8, 0, 1, 3)

        self.setLayout(grid)

        self.setWindowTitle(u'查看邮件：' + self.MailMsg['subject'])

        self.show()
    
    def Mailoutput(self):

        fname = QtGui.QFileDialog.getSaveFileName(self, u'保存文件', u'', u'*.eml')
        sfname = str(fname)
        if sfname == '':
            return
        sfname = os.path.abspath(sfname)
        eml_gen.save_to_eml(self.MailMsg, sfname)


class QMainArea(QtGui.QWidget):

    def __init__(self):

        super(QMainArea, self).__init__()
        self.initLayout()

    def initLayout(self):

        grid = QtGui.QGridLayout()
        grid.setSpacing(5)

        self.MailList = QtGui.QTreeWidget()
        self.MailList.setHeaderLabels([u'#' ,u'时间', u'收件人',
                                       u'发件人', u'标题', u'预览'])
        self.MailList.setColumnWidth(0, 60)
        self.MailList.setColumnWidth(1, 120)
        self.MailList.setColumnWidth(2, 200)
        self.MailList.setColumnWidth(3, 200)
        self.MailList.setColumnWidth(4, 100)
        self.MailList.itemDoubleClicked.connect(self.getMailView)

        grid.addWidget(self.MailList)

        self.setLayout(grid) 

    def updateInfo(self, list_db):

        self.MailList.clear() 
        self.list_db = list_db

        for i in range(len(self.list_db)):
            MailInfo = QtGui.QTreeWidgetItem()
            receive_date = self.list_db[i]['receive_date']
            date_s = str(receive_date.year) + '-' + \
                     str(receive_date.month) + '-' + \
                     str(receive_date.day) + ' ' + \
                     str(receive_date.hour) + ':' + \
                     str(receive_date.second)
            MailInfo.setText(0, str(i))
            MailInfo.setText(1, date_s)
            MailInfo.setText(2, self.list_db[i]['to_address'])
            MailInfo.setText(3, self.list_db[i]['from_address'])
            MailInfo.setText(4, self.list_db[i]['subject'])
            MailInfo.setText(5, self.list_db[i]['snippet'])
            self.MailList.addTopLevelItem(MailInfo)

    def getMailView(self):

        MailId = int(self.MailList.currentItem().text(0))
        MailView = QMailView(self.list_db[MailId], parent = self)
        MailView.exec_()
        MailView.destroy()

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

        fname = QtGui.QFileDialog.getOpenFileName(self, u'打开文件', u'', u'*.db')
        sfname = str(fname)
        if sfname == '':
            return
        sfname = os.path.abspath(sfname)
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

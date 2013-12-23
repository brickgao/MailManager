#!/usr/bin/env python2
# -*- coding: utf-8 -*-

import sys, os, eml_gen, shutil, re, con_db
from PyQt4 import QtGui, QtCore, QtWebKit

class QOutputDialog(QtGui.QDialog):

    def __init__(self, list_db, parent = None):

        self.list_db = list_db
        QtGui.QDialog.__init__(self, parent)
        self.initLayout()

    def initLayout(self):

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        self.OutputView = QtGui.QTreeWidget()
        self.OutputView.setHeaderLabels([u'' , u'#', u'时间',
                                         u'收件人', u'发件人', u'标题', u'预览'])
        self.OutputView.setColumnWidth(0, 200)
        self.OutputView.setColumnWidth(1, 60)
        self.OutputView.setColumnWidth(2, 120)
        self.OutputView.setColumnWidth(3, 200)
        self.OutputView.setColumnWidth(4, 200)
        self.OutputView.setColumnWidth(5, 100)

        self.updateInfo()

        grid.addWidget(self.OutputView)

        self.setLayout(grid)

        self.setGeometry(120, 120, 1000, 600)

        self.setWindowTitle(u'批量导出至 *.eml')
        self.show()

    def updateInfo(self):

        for _ in self.list_db:
            userInfo = QtGui.QTreeWidgetItem()
            userInfo.setCheckState(0, QtCore.Qt.Unchecked)
            user = _.dbs.keys()[0]
            userInfo.setText(0, user)
            self.OutputView.addTopLevelItem(userInfo)
            self.OutputView.expandItem(userInfo)
            for __ in _.dbs[user]:
                mailboxInfo = QtGui.QTreeWidgetItem(userInfo)
                mailboxInfo.setCheckState(0, QtCore.Qt.Unchecked)
                mailbox = __
                mailboxInfo.setText(0, mailbox)
                self.OutputView.addTopLevelItem(mailboxInfo)
                self.OutputView.expandItem(mailboxInfo)
                currentList = _.dbs[user][__]
                for i in range(len(currentList)):
                    mailInfo = QtGui.QTreeWidgetItem(mailboxInfo)
                    mailInfo.setCheckState(1, QtCore.Qt.Unchecked)
                    receive_date = currentList[i]['receive_date']
                    date_s = str(receive_date.year) + '-' + \
                             str(receive_date.month) + '-' + \
                             str(receive_date.day) + ' ' + \
                             str(receive_date.hour) + ':' + \
                             str(receive_date.second)
                    mailInfo.setText(1, str(i))
                    mailInfo.setText(2, date_s)
                    mailInfo.setText(3, currentList[i]['to_address'])
                    mailInfo.setText(4, currentList[i]['from_address'])
                    mailInfo.setText(5, currentList[i]['subject'])
                    mailInfo.setText(6, currentList[i]['snippet'])
                    self.OutputView.addTopLevelItem(mailInfo)
            

        

class QMailView(QtGui.QDialog):

    def __init__(self, mailMsg, parent = None):

        self.mailMsg = mailMsg
        QtGui.QDialog.__init__(self, parent)
        self.initLayout()

    def initLayout(self):

        self.from_address = QtGui.QLabel(u'发信人：')
        self.from_address_val = QtGui.QLabel(self.mailMsg['from_address'])

        self.to_address = QtGui.QLabel(u'收信人：')
        self.to_address_val = QtGui.QLabel(self.mailMsg['to_address'])

        self.subject = QtGui.QLabel(u'标题：')
        self.subject_val = QtGui.QLabel(self.mailMsg['subject'])

        self.receive_date = QtGui.QLabel(u'收信时间：')
        self.receive_date_val = self.mailMsg['receive_date']
        self.date_s = QtGui.QLabel(
                    str(self.receive_date_val.year) + '-' + \
                    str(self.receive_date_val.month) + '-' + \
                    str(self.receive_date_val.day) + ' ' + \
                    str(self.receive_date_val.hour) + ':' + \
                    str(self.receive_date_val.second))

        self.ContentView = QtGui.QLabel(u'邮件正文：')
        self.ContentView_val = QtWebKit.QWebView()
        HtmlCode = unicode(self.mailMsg['body'], 'utf-8')
        self.ContentView_val.setHtml(HtmlCode,
                            QtCore.QUrl(''))

        self.attachmentsList = QtGui.QLabel(u'附件：')
        self.attachmentsList_val = QtGui.QTreeWidget()
        self.attachmentsList_val.setHeaderLabels([u'#' ,u'文件名',
                                             u'文件大小', u'是否存在',
                                             u'文件路径'])
        self.attachmentsList_val.setColumnWidth(0, 60)
        self.attachmentsList_val.itemDoubleClicked.connect(self.attachmentsOpen)
        for i in range(len(self.mailMsg['attachments'])):
            fileInfo = QtGui.QTreeWidgetItem()
            fileInfo.setText(0, str(i))
            fileInfo.setText(1, self.mailMsg['attachments'][i]['name'])
            fileInfo.setText(2, self.mailMsg['attachments'][i]['format_size'])
            fileInfo.setText(3, str(self.mailMsg['attachments'][i]['exist']))
            if self.mailMsg['attachments'][i]['exist'] == True:
                filepath = self.mailMsg['attachments'][i]['path']
                fileInfo.setText(4, filepath)
            else:
                fileInfo.setText(4, u'文件不存在')
            self.attachmentsList_val.addTopLevelItem(fileInfo)

        self.mailOutputBtn = QtGui.QPushButton(u'导出当前文件')
        self.mailOutputBtn.clicked.connect(self.mailOutput)

        self.attachmentsOutputBtn = QtGui.QPushButton(u'导出选中附件')
        self.attachmentsOutputBtn.clicked.connect(self.attachmentsOutput)

        grid = QtGui.QGridLayout()
        grid.setSpacing(10)

        grid.addWidget(self.from_address, 1, 0)
        grid.addWidget(self.from_address_val, 1, 1)
        grid.addWidget(self.to_address, 2, 0)
        grid.addWidget(self.to_address_val, 2, 1)
        grid.addWidget(self.subject, 3, 0)
        grid.addWidget(self.subject_val, 3, 1)
        grid.addWidget(self.receive_date, 4, 0)
        grid.addWidget(self.date_s, 4, 1)
        grid.addWidget(self.attachmentsList, 5, 0)
        grid.addWidget(self.attachmentsOutputBtn, 5, 2)
        grid.addWidget(self.attachmentsList_val, 6, 0, 1, 3)
        grid.addWidget(self.ContentView, 7, 0)
        grid.addWidget(self.mailOutputBtn, 7, 2)
        grid.addWidget(self.ContentView_val, 8, 0, 1, 3)

        self.setLayout(grid)

        self.setWindowTitle(u'查看邮件：' + self.mailMsg['subject'])

        self.show()


    def mailOutput(self):

        fname = QtGui.QFileDialog.getSaveFileName(self, u'保存文件', u'', u'*.eml')
        fname = unicode(fname)
        if fname == '':
            return
        fname = os.path.abspath(fname)
        eml_gen.save_to_eml(self.mailMsg, fname)

    def attachmentsOpen(self):

        Id = int(self.attachmentsList_val.currentItem().text(0))
        if not self.mailMsg['attachments'][Id]['exist']:
            return self.errorAlert(u'请选择存在的文件')
        fname = self.mailMsg['attachments'][Id]['path'].encode('gbk')
        os.system(fname)

    def attachmentsOutput(self):

        if not self.attachmentsList_val.currentItem():
            return self.errorAlert(u'请选择文件')
        Id = int(self.attachmentsList_val.currentItem().text(0))
        if not self.mailMsg['attachments'][Id]['exist']:
            return self.errorAlert(u'请选择存在的文件')
        fname = QtGui.QFileDialog.getSaveFileName(self, u'保存附件', self.mailMsg['attachments'][Id]['name'], u'*.*')
        fname = unicode(fname)
        if fname == '':
            return
        shutil.copyfile(self.mailMsg['attachments'][Id]['path'], fname)


    def errorAlert(self, s):

        QtGui.QMessageBox.critical(self, u'错误', s)



class QMainArea(QtGui.QWidget):

    def __init__(self):

        super(QMainArea, self).__init__()
        self.initLayout()


    def initLayout(self):

        grid = QtGui.QGridLayout()
        grid.setSpacing(5)

        self.mailboxList = QtGui.QTreeWidget()
        self.mailboxList.setHeaderLabels([u''])
        self.mailboxList.itemClicked.connect(self.getMailList)

        self.mailList = QtGui.QTreeWidget()
        self.mailList.setHeaderLabels([u'#' ,u'时间', u'收件人',
                                       u'发件人', u'标题', u'预览'])
        self.mailList.setColumnWidth(0, 60)
        self.mailList.setColumnWidth(1, 120)
        self.mailList.setColumnWidth(2, 200)
        self.mailList.setColumnWidth(3, 200)
        self.mailList.setColumnWidth(4, 100)
        self.mailList.itemDoubleClicked.connect(self.getMailView)

        grid.addWidget(self.mailboxList, 1, 0)
        grid.addWidget(self.mailList, 1, 1)

        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 4)

        self.setLayout(grid)


    def updateInfo(self, list_db):

        self.mailList.clear()

        self.mailboxList.clear()
        self.list_db = list_db

        for _ in list_db:
            user = QtGui.QTreeWidgetItem()
            userName = _.dbs.keys()[0]
            user.setText(0, userName)
            self.mailboxList.addTopLevelItem(user)
            self.mailboxList.expandItem(user)
            for __ in _.dbs[userName]:
                mailbox = QtGui.QTreeWidgetItem(user)
                mailboxName = __
                mailbox.setText(0, __)
                self.mailboxList.addTopLevelItem(mailbox)

    def getMailList(self):
            
        self.currentList = []

        if not self.mailboxList.currentItem().parent():

            userName = unicode(self.mailboxList.currentItem().text(0))

            self.mailList.clear()

            for _ in self.list_db:
                if userName in _.dbs:
                    for __ in _.dbs[userName].keys():
                        ___ = _.dbs[userName][__]
                        self.currentList += ___
        else:

            userName = unicode(self.mailboxList.currentItem().parent().text(0))
            mailboxName = unicode(self.mailboxList.currentItem().text(0))

            self.mailList.clear()

            for _ in self.list_db:
                if userName in _.dbs:
                    self.currentList = _.dbs[userName][mailboxName]

        for i in range(len(self.currentList)):
            mailInfo = QtGui.QTreeWidgetItem()
            receive_date = self.currentList[i]['receive_date']
            date_s = str(receive_date.year) + '-' + \
                     str(receive_date.month) + '-' + \
                     str(receive_date.day) + ' ' + \
                     str(receive_date.hour) + ':' + \
                     str(receive_date.second)
            mailInfo.setText(0, str(i))
            mailInfo.setText(1, date_s)
            mailInfo.setText(2, self.currentList[i]['to_address'])
            mailInfo.setText(3, self.currentList[i]['from_address'])
            mailInfo.setText(4, self.currentList[i]['subject'])
            mailInfo.setText(5, self.currentList[i]['snippet'])
            self.mailList.addTopLevelItem(mailInfo)
            


    def getMailView(self):

        mailId = int(self.mailList.currentItem().text(0))
        mailView = QMailView(self.currentList[mailId], parent = self)
        mailView.exec_()
        mailView.destroy()



class MainWindow(QtGui.QMainWindow):

    def __init__(self):
        
        self.list_db = []
        super(MainWindow, self).__init__()
        self.initLayout()


    def initLayout(self):

        self.statusBar()

        menubar = self.menuBar()

        exitAction = QtGui.QAction(u'退出', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip(u'退出 mail Manager')
        exitAction.triggered.connect(QtGui.qApp.quit)

        inputAction = QtGui.QAction(u'打开', self)
        inputAction.setStatusTip(u'打开文件')
        inputAction.triggered.connect(self.inputFile)

        outputAction = QtGui.QAction(u'批量导出', self)
        outputAction.setStatusTip(u'将邮件批量导出至 *.eml')
        outputAction.triggered.connect(self.outputFile)

        fileMenu = menubar.addMenu(u'文件')
        fileMenu.addAction(inputAction)
        fileMenu.addAction(outputAction)
        fileMenu.addAction(exitAction)

        self.mainArea = QMainArea()

        self.setCentralWidget(self.mainArea)

        self.setGeometry(100, 100, 1000, 600)
        self.setWindowTitle(u'Mail Manager')
        self.show()


    def inputFile(self):

        fname = QtGui.QFileDialog.getExistingDirectory(self, u'打开文件夹')
        if fname == '':
            return
        fname = os.path.abspath(fname)
        try:
            alist = os.listdir(fname + '\\databases')
        except Exception, e:
            print e
            self.errorAlert(u'打开 db 错误，请选择正确的文件夹')
        else:
            flist = []
            r = re.compile(r'mailstore')
            for _ in alist:
                if r.match(_):
                    flist.append(fname + '\\databases\\' + _)
            try:
                for _ in flist:
                    __ = con_db.mail_dbs()
                    __.open_db(_)
                    tag = True
                    for ___ in self.list_db:
                        if ___.dbs[___.dbs.keys()[0]] == __.dbs[__.dbs.keys()[0]]:
                            tag = False
                            ___.open_db(_)
                    if tag:     self.list_db.append(__)
            except Exception, e:
                print e
                self.errorAlert(u'解析 db 错误，请选择正确的 db 文件')
            else:
                self.mainArea.updateInfo(self.list_db)

    def outputFile(self):

        dialog = QOutputDialog(self.list_db, parent = self)
        dialog.exec_()
        dialog.destroy()

    def errorAlert(self, s):

        QtGui.QMessageBox.critical(self, u'错误', s)



def main():

    app = QtGui.QApplication(sys.argv)
    ex = MainWindow()
    sys.exit(app.exec_())



if __name__ == '__main__':
    main()

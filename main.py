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
        self.OutputView.setHeaderLabels([u'' , u'#', u'附件', u'时间',
                                         u'收件人', u'发件人', u'标题', u'预览'])
        self.OutputView.setColumnWidth(0, 200)
        self.OutputView.setColumnWidth(1, 60)
        self.OutputView.setColumnWidth(2, 40)
        self.OutputView.setColumnWidth(3, 120)
        self.OutputView.setColumnWidth(4, 200)
        self.OutputView.setColumnWidth(5, 200)
        self.OutputView.setColumnWidth(6, 100)

        self.updateInfo()

        self.OutputView.itemChanged.connect(self.updateCheckbox)
        
        self.mailOutputBtn = QtGui.QPushButton(u'导出当前选择文件')
        self.mailOutputBtn.clicked.connect(self.mailOutput)

        self.selectAll = QtGui.QPushButton(u'全选')
        self.selectAll.clicked.connect(self.selectAllAction)

        self.removeAll = QtGui.QPushButton(u'重置')
        self.removeAll.clicked.connect(self.removeAllAction)

        grid.addWidget(self.OutputView, 0, 0, 1, 5)
        grid.addWidget(self.selectAll, 1, 0, 1, 1)
        grid.addWidget(self.removeAll, 1, 1, 1, 1)
        grid.addWidget(self.mailOutputBtn, 1, 4, 1, 1)
    
        self.setLayout(grid)

        self.setGeometry(50, 50, 1200, 600)

        self.setWindowTitle(u'批量导出至 *.eml')
        self.show()

    def selectAllAction(self):
        
        self.updateAllcheckbox(QtCore.Qt.Checked)

    def removeAllAction(self):

        self.updateAllcheckbox(QtCore.Qt.Unchecked)

    def updateAllcheckbox(self, state):

        root = self.OutputView.invisibleRootItem()
        for i in range(root.childCount()):
            user = root.child(i)
            self.updateFromUser(user, state)
        

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
                    if len(currentList[i]['attachments']):          mailInfo.setText(2, u'*')
                    else:                                           mailInfo.setText(2, u'')
                    mailInfo.setText(3, date_s)
                    mailInfo.setText(4, currentList[i]['to_address'])
                    mailInfo.setText(5, currentList[i]['from_address'])
                    mailInfo.setText(6, currentList[i]['subject'])
                    mailInfo.setText(7, currentList[i]['snippet'])
                    self.OutputView.addTopLevelItem(mailInfo)

    def mailOutput(self):
        
        _ = []

        dirName = QtGui.QFileDialog.getExistingDirectory(self, u'打开文件夹')
        if dirName == '':   return
        dirName = os.path.abspath(dirName)

        root = self.OutputView.invisibleRootItem()
        for i in range(root.childCount()):
            user = root.child(i)
            userName = unicode(user.text(0))
            currentdb = {}
            for __ in self.list_db:
                if userName in __.dbs:
                    currentdb = __.dbs
            for j in range(user.childCount()):
                mailbox = user.child(j)
                mailboxName = unicode(mailbox.text(0))
                for k in range(mailbox.childCount()):
                    mail = mailbox.child(k)
                    mailId = int(mail.text(1))
                    if mail.checkState(1) == QtCore.Qt.Checked:
                        _.append(currentdb[userName][mailboxName][mailId])

        if not len(_):
            return self.errorAlert(u'请选择文件')

        eml_gen.save_all_emails_from_list(_, dirName)

        self.infoAlert(u'导出成功')

    def updateCheckbox(self, item, column):
        
        self.OutputView.blockSignals(True)

        if item.checkState(column) == QtCore.Qt.Checked:
            if item.childCount():
                if item.parent():
                    self.updateFromMailbox(item, QtCore.Qt.Checked)
                else:
                    self.updateFromUser(item, QtCore.Qt.Checked)

        elif item.checkState(column) == QtCore.Qt.Unchecked:
            if item.childCount():
                if item.parent():
                    self.updateFromMailbox(item, QtCore.Qt.Unchecked)
                else:
                    self.updateFromUser(item, QtCore.Qt.Unchecked)

        root = self.OutputView.invisibleRootItem()
        for i in range(root.childCount()):
            user = root.child(i)
            __ = user.childCount() * 2
            for j in range(user.childCount()):
                mailbox = user.child(j)
                _ = mailbox.childCount() * 2
                for k in range(mailbox.childCount()):
                    mail = mailbox.child(k)
                    if mail.checkState(1) == QtCore.Qt.Unchecked:           _ -= 2
                    if mail.checkState(1) == QtCore.Qt.PartiallyChecked:    _ -= 1
                if _ == 0:
                    mailbox.setCheckState(0, QtCore.Qt.Unchecked)
                elif _ == mailbox.childCount() * 2:
                    mailbox.setCheckState(0, QtCore.Qt.Checked)
                else:
                    mailbox.setCheckState(0, QtCore.Qt.PartiallyChecked)
                if mailbox.checkState(0) == QtCore.Qt.Unchecked:           __ -= 2
                if mailbox.checkState(0) == QtCore.Qt.PartiallyChecked:    __ -= 1
            if __ == 0:
                user.setCheckState(0, QtCore.Qt.Unchecked)
            elif __ == user.childCount() * 2:
                user.setCheckState(0, QtCore.Qt.Checked)
            else:
                user.setCheckState(0, QtCore.Qt.PartiallyChecked)

        self.OutputView.blockSignals(False)

    def updateFromUser(self, item, state):

        for i in range(item.childCount()):
            mailbox = item.child(i)
            mailbox.setCheckState(0, state)
            for j in range(mailbox.childCount()):
                mail = mailbox.child(j)
                mail.setCheckState(1, state)

    def updateFromMailbox(self, item, state):

        for i in range(item.childCount()):
            mail = item.child(i)
            mail.setCheckState(1, state)

    def errorAlert(self, s):

        QtGui.QMessageBox.critical(self, u'错误', s)

    def infoAlert(self, s):
        
        QtGui.QMessageBox.information(self, u'提示', s)
        
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
            fileInfo.setCheckState(0, QtCore.Qt.Unchecked)
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
        self.infoAlert(u'导出成功')

    def attachmentsOpen(self):

        Id = int(self.attachmentsList_val.currentItem().text(0))
        if not self.mailMsg['attachments'][Id]['exist']:
            return self.errorAlert(u'请选择存在的文件')
        fname = self.mailMsg['attachments'][Id]['path'].encode('gbk')
        os.popen(fname)

    def attachmentsOutput(self):

        _l = []
        root = self.attachmentsList_val.invisibleRootItem()
        for i in range(root.childCount()):
            attachment = root.child(i)
            if attachment.checkState(0):    _l.append(int(attachment.text(0)))
        if not len(_l):     return self.errorAlert(u'请选择文件')
        for _id in _l:
            if not self.mailMsg['attachments'][_id]['exist']:
                return self.errorAlert(u'请选择存在的文件')
        fname = QtGui.QFileDialog.getExistingDirectory(self, u'保存到文件夹')
        fname = unicode(fname)
        for _id in _l:
            _fname = fname + u'\\' + unicode(self.mailMsg['attachments'][_id]['name'])
            print _fname
            shutil.copyfile(self.mailMsg['attachments'][_id]['path'], _fname)
        self.infoAlert(u'导出成功')
            

    def errorAlert(self, s):

        QtGui.QMessageBox.critical(self, u'错误', s)

    def infoAlert(self, s):
        
        QtGui.QMessageBox.information(self, u'提示', s)



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
        self.mailList.setHeaderLabels([u'#' ,u'附件', u'时间', u'收件人',
                                       u'发件人', u'标题', u'预览'])
        self.mailList.setColumnWidth(0, 60)
        self.mailList.setColumnWidth(1, 40)
        self.mailList.setColumnWidth(2, 120)
        self.mailList.setColumnWidth(3, 200)
        self.mailList.setColumnWidth(4, 200)
        self.mailList.setColumnWidth(5, 100)
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
            if len(self.currentList[i]['attachments']):     mailInfo.setText(1, u'*')
            else:                                           mailInfo.setText(1, u'')
            mailInfo.setText(2, date_s)
            mailInfo.setText(3, self.currentList[i]['to_address'])
            mailInfo.setText(4, self.currentList[i]['from_address'])
            mailInfo.setText(5, self.currentList[i]['subject'])
            mailInfo.setText(6, self.currentList[i]['snippet'])
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
        fname = os.path.abspath(unicode(fname))
        try:
            alist = os.listdir(fname + '\\databases')
        except Exception, e:
            print e
            self.errorAlert(u'打开 db 错误，请选择正确的文件夹')
        else:
            flist = []
            r = re.compile(r'mailstore\.(.*)\.db$')
            for _ in alist:
                if r.match(_):
                    flist.append(fname + '\\databases\\' + _)
            try:
                for _ in flist:
                    __ = con_db.mail_dbs()
                    __.open_db(_)
                    tag = True
                    for ___ in self.list_db:
                        if ___.dbs.keys()[0] == __.dbs.keys()[0]:
                            tag = False
                            ___.open_db(_)
                            break
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

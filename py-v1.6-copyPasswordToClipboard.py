# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication,QMainWindow,QMessageBox
import sys
import pandas as pd
import win32clipboard
import datetime


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(600, 400)

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.formLayout = QtWidgets.QFormLayout(self.centralwidget)
        self.formLayout.setObjectName("formLayout")

        r=0
        # self.label = QtWidgets.QLabel(self.centralwidget)
        # self.label.setObjectName("label")
        # self.formLayout.setWidget(r, QtWidgets.QFormLayout.LabelRole, self.label)
        self.button = QtWidgets.QPushButton(self.centralwidget)
        self.button.setObjectName("button")
        self.formLayout.setWidget(r, QtWidgets.QFormLayout.LabelRole, self.button)

        self.lineEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit.setObjectName("lineEdit")
        self.formLayout.setWidget(r, QtWidgets.QFormLayout.FieldRole, self.lineEdit)

        r+=1
        self.tableView = QtWidgets.QTableView(self.centralwidget)
        self.tableView.setObjectName("tableView")
        self.formLayout.setWidget(r, QtWidgets.QFormLayout.SpanningRole, self.tableView)

        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)


    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "复制密码到剪贴板"))

        # self.label.setText(_translate("Dialog", "搜索:"))
        self.button.setText('🌊')
        self.button.setToolTip('点击刷新')
        self.button.setStyleSheet('QPushButton{border:0px solid;border-radius:12px;font-size:18px;width:30px;height:26px}'
                                  'QPushButton::hover{background-color:#00FFFF}'
                                  'QPushButton::Pressed{background-color:#008000}')
        self.button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.button.clicked.connect(lambda :self.refreshWindow())

        self.lineEdit.setToolTip('输入密码描述，自动搜索，列在下面')

        self.lineEdit.textChanged.connect(lambda : self.tableView_seach(self.lineEdit.text()))

        self.tableView_init()


    def tableView_init(self):
        self.tableView.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)   # 所有列自动拉伸，充满界面
        self.tableView.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)    # 设置只能选中整行
        self.tableView.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)     # 设置只能选一行
        self.tableView.setEditTriggers(QtWidgets.QTableView.NoEditTriggers)     # 不可编辑

        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.tableView.setSizePolicy(sizePolicy)

        self.getDateFrame()
        
        self.tableView_add(self.df.iterrows())

        self.tableView.clicked.connect(lambda : self.setStrToClipboard())

        self.setStyleSheet('QHeaderView::section{background-color:#00FFFF}')

        self.tableView.setToolTip('选中条目，自动复制对应的密码到剪贴板中')


    COLOUMN_DESC=0
    COLOUMN_ID=1
    COLOUMN_PASSWORD=2

    def refreshWindow(self):
        self.getDateFrame()

        if self.lineEdit.text()!='':
            self.lineEdit.clear()
        else:
            self.tableView_add(self.df.iterrows())

        self.tableView.scrollToTop()

    def getDateFrame(self):
        try:
            filePath='密码.xlsx'
            # 不用行、列作index、head,使用列
            self.df=pd.read_excel(filePath,index_col=None,header=None,usecols=[self.COLOUMN_DESC,self.COLOUMN_ID,self.COLOUMN_PASSWORD])
            self.df.dropna(axis='index',how='all',inplace=True)     # 删除全为缺失的行,inplace使原df改变
            self.df.fillna('',inplace=True)         # 用空填充缺失值

            self.list_head=[self.df[self.COLOUMN_DESC][0],self.df[self.COLOUMN_ID][0]]
            self.df=self.df.iloc[1:,:]  # 放弃第一行

            # print(self.df)
        except FileNotFoundError as e:
            QMessageBox.critical(self, '文件不存在', '此文件夹下不存在“' + filePath + '”！')
            sys.exit()


    def tableView_add(self,iterrows):
        # 创建0行1列的标准模型
        self.model=QtGui.QStandardItemModel(0, 1)

        # 设置表头标签
        self.model.setHorizontalHeaderLabels(self.list_head)

        for i,se in iterrows:
            passwordDesc=QtGui.QStandardItem(str(se[self.COLOUMN_DESC]))
            ID=QtGui.QStandardItem(str(se[self.COLOUMN_ID]))
            self.model.appendRow([passwordDesc,ID])    # 添加值

        self.tableView.setModel(self.model)


    def tableView_seach(self,str_input):
        str_search=str_input.strip()
        if str_search=='':
            if str_input=='':
                self.tableView_add(self.df.iterrows())
            return

        # 查询
        df_seach=self.df[self.df[self.COLOUMN_DESC].str.contains(str_search,case=False,regex=False)]
        # print(df_seach)
        self.tableView_add(df_seach.iterrows())


    def setStrToClipboard(self):
        row_select=self.tableView.currentIndex().row()

        str_select=self.model.item(row_select,self.COLOUMN_DESC).text()

        # print(str_select)

        df_select=self.df[self.df[self.COLOUMN_DESC]==str_select]

        str_text=str(df_select[self.COLOUMN_PASSWORD].values[0]).strip()

        # print(str_text)

        win32clipboard.OpenClipboard()

        win32clipboard.EmptyClipboard()

        win32clipboard.SetClipboardText(str_text)

        win32clipboard.CloseClipboard()

        self.showMessageBox()


    def showMessageBox(self):
        info_box = QMessageBox()
        # 因为没使用这种方式 QMessageBox.information(self, '复制', '复制成功', QMessageBox.Yes) 写弹出框，
        # 则主窗口的样式不能应用在QMessageBox中，因此重新写了弹出框的部件样式

        info_box.setStyleSheet('QPushButton{font-weight: bold; background: skyblue; border-radius: 14px;'
                                'width: 64px; height: 28px; font-size: 18px; text-align: center;}'
                                'QLabel{font-weight: bold; font-size: 20px; color: #008000}')

        info_box.setWindowTitle('成功')     # QMessageBox标题
        info_box.setText('成功，复制密码到剪贴板！')     #  QMessageBox的提示文字
        info_box.setStandardButtons(QMessageBox.Ok)      # QMessageBox显示的按钮
        info_box.button(QMessageBox.Ok).animateClick(1000)    # t时间后自动关闭(t单位为毫秒)
        info_box.exec_()    # 如果使用.show(),会导致QMessageBox框一闪而逝


class MyMainForm(QMainWindow,Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainForm, self).__init__(parent)
        try:
            self.setupUi(self)
        except Exception as e:
            QMessageBox.critical(self, '错误', str(e))

            with open('error_copyPasswordToClipboard.txt','a',encoding='utf8') as f:
                f.write(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S ')+str(e) + '\n')
            sys.exit()

if __name__=='__main__':
    app=QApplication(sys.argv)
    myWin=MyMainForm()
    myWin.show()
    sys.exit(app.exec_())

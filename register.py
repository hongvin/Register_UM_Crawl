from bs4 import BeautifulSoup
from urllib.request import Request,urlopen
from urllib.error import HTTPError
from PyQt5 import QtCore, QtGui, QtWidgets, Qt
import sys
import threading
import datetime 
import win32con
import os
import struct
import time
from win32api import *
from win32gui import *

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1236, 900)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 361, 771))
        self.groupBox.setObjectName("groupBox")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBox)
        self.tableWidget.setGeometry(QtCore.QRect(10, 30, 341, 721))
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(True)
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(380, 10, 681, 771))
        self.groupBox_2.setObjectName("groupBox_2")
        self.tableWidget_2 = QtWidgets.QTableWidget(self.groupBox_2)
        self.tableWidget_2.setGeometry(QtCore.QRect(10, 30, 651, 721))
        self.tableWidget_2.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget_2.setObjectName("tableWidget_2")
        self.tableWidget_2.setColumnCount(0)
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.horizontalHeader().setCascadingSectionResizes(True)
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(1070, 310, 91, 16))
        self.label.setObjectName("label")
        self.groupBox_3 = QtWidgets.QGroupBox(Form)
        self.groupBox_3.setGeometry(QtCore.QRect(1070, 10, 141, 80))
        self.groupBox_3.setObjectName("groupBox_3")
        self.pushButton = QtWidgets.QPushButton(self.groupBox_3)
        self.pushButton.setGeometry(QtCore.QRect(10, 50, 111, 28))
        self.pushButton.setObjectName("pushButton")
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit.setGeometry(QtCore.QRect(10, 20, 113, 22))
        self.lineEdit.setObjectName("lineEdit")
        self.groupBox_4 = QtWidgets.QGroupBox(Form)
        self.groupBox_4.setGeometry(QtCore.QRect(1070, 100, 141, 61))
        self.groupBox_4.setObjectName("groupBox_4")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox_4)
        self.lineEdit_2.setGeometry(QtCore.QRect(10, 20, 113, 22))
        self.lineEdit_2.setText("")
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.groupBox_5 = QtWidgets.QGroupBox(Form)
        self.groupBox_5.setGeometry(QtCore.QRect(1070, 170, 141, 51))
        self.groupBox_5.setObjectName("groupBox_5")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox_5)
        self.lineEdit_3.setGeometry(QtCore.QRect(10, 20, 113, 22))
        self.lineEdit_3.setText("")
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.groupBox_6 = QtWidgets.QGroupBox(Form)
        self.groupBox_6.setGeometry(QtCore.QRect(1070, 230, 141, 51))
        self.groupBox_6.setObjectName("groupBox_6")
        self.lineEdit_4 = QtWidgets.QLineEdit(self.groupBox_6)
        self.lineEdit_4.setGeometry(QtCore.QRect(10, 20, 113, 22))
        self.lineEdit_4.setText("")
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(10, 790, 1211, 101))
        self.listWidget.setObjectName("listWidget")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Form"))
        self.groupBox.setTitle(_translate("Form", "Elective"))
        self.tableWidget.setSortingEnabled(True)
        self.groupBox_2.setTitle(_translate("Form", "KoK"))
        self.tableWidget_2.setSortingEnabled(True)
        self.label.setText(_translate("Form", "Koay Hong Vin"))
        self.groupBox_3.setTitle(_translate("Form", "Set Timer"))
        self.pushButton.setText(_translate("Form", "START!!"))
        self.lineEdit.setText(_translate("Form", "5"))
        self.groupBox_4.setTitle(_translate("Form", "Targeted Elective 1"))
        self.groupBox_5.setTitle(_translate("Form", "Targeted Elective 2"))
        self.groupBox_6.setTitle(_translate("Form", "Targeted KoK 1"))



class WindowsBalloonTip:
    def __init__(self):
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }
        # Register the Window class.
        wc = WNDCLASS()
        self.hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        self.classAtom = RegisterClass(wc)

    def ShowWindow(self,title, msg):
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( self.classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, self.hinst, None)
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], "balloontip.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
           hicon = LoadImage(self.hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",msg,200,title))
        # self.show_balloon(title, msg)
        DestroyWindow(self.hwnd)

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        nid = (self.hwnd, 0)
        Shell_NotifyIcon(NIM_DELETE, nid)
        PostQuitMessage(0) # Terminate the app.

w=WindowsBalloonTip()

class App(QtWidgets.QMainWindow,Ui_Form):
    def __init__(self):
        super(self.__class__,self).__init__()
        self.setupUi(self)      
        self.pushButton.clicked.connect(self.startengine)
        
    def startengine(self):
        self.listWidget.scrollToBottom()
        self.lineEdit.setEnabled(False)
        timeout=float(self.lineEdit.text())
        t=threading.Timer(timeout,self.startengine)
        t.daemon=True
        t.start()

        e1=self.lineEdit_2.text()
        e2=self.lineEdit_3.text()
        k1=self.lineEdit_4.text()
        e1b=False
        e2b=False
        k1b=False
        courses=['','','']

        url = 'http://register.um.edu.my/el_kosong_bi.asp' 
        request = Request(url)
        try:
            self.tableWidget.setEnabled(True)
            json = urlopen(request).read().decode()
            soup = BeautifulSoup(json,"html.parser")
            a = soup.find_all('div')
            self.tableWidget.setColumnCount(3)
            self.tableWidget.setRowCount((len(a)-1)/3)
            self.tableWidget.setHorizontalHeaderLabels(["Subject Code","Group","Vacant"])
            j=-1
            k=0
            for i in range(0,len(a)):
                if i%3==0:
                    j+=1
                    k=0
                self.tableWidget.setItem(j,k,QtWidgets.QTableWidgetItem(a[i].text))
                if e1==a[i].text:
                    self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Matched found for targeted elective '+a[i].text)
                    e1b=True
                    courses[0]=(a[i].text)
                if e2==a[i].text:
                    self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Matched found for targeted elective '+a[i].text)
                    e2b=True
                    courses[1]=(a[i].text)
                    
                k+=1
        except HTTPError as e:
            print("Error",e)
            self.tableWidget.setEnabled(False)
            self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Error occured at Elective')
        
        url = 'http://register.um.edu.my/kok_kosong_bi.asp' 
        request = Request(url)
        try:
            self.tableWidget_2.setEnabled(True)
            json = urlopen(request).read().decode()
            soup = BeautifulSoup(json,"html.parser")
            a = soup.find_all('td')
            self.tableWidget_2.setColumnCount(3)
            self.tableWidget_2.setRowCount((len(a)-10)/4)
            self.tableWidget_2.setHorizontalHeaderLabels(["Subject Code","Course Name","Vacant"])
            j=-1
            k=0
            m=1
            for i in range(9,len(a)-1):
                if (i-m)%4==0:
                    j+=1
                    k=0
                    continue
                if a[i].text=='Bil' or a[i].text=='Code' or a[i].text=='Description' or a[i].text=='Vacant':
                    m=2
                    if a[i].text=='Vacant':
                        j-=1
                        self.tableWidget_2.setRowCount(((len(a)-10)/4)-1)
                    continue
                if k1==a[i].text:
                    self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Matched found for targeted KoK '+a[i].text)
                    k1b=True
                    courses[2]=(a[i].text)
                self.tableWidget_2.setItem(j,k,QtWidgets.QTableWidgetItem(a[i].text))
                k+=1
        except HTTPError as e:
            print("Error",e)
            self.tableWidget.setEnabled(False)
            self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Error occured at KoK')
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.listWidget.scrollToBottom()
        self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: List refreshed')
        if e1b==True or e2b==True or k1b==True:
            w.ShowWindow("Matching course found!","Course Code: {0} {1} {2}".format(*courses))
        
            
        

    
def main():  
    app=QtWidgets.QApplication(sys.argv)
    form=App()
    form.show()
    app.exec_()

if __name__ == '__main__':
    main()

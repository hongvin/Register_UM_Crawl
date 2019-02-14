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
import pyttsx3
from win32api import *
from win32gui import *

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(1236, 996)
        self.groupBox = QtWidgets.QGroupBox(Form)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 361, 831))
        self.groupBox.setObjectName("groupBox")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBox)
        self.tableWidget.setGeometry(QtCore.QRect(10, 30, 341, 791))
        self.tableWidget.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.setRowCount(0)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(True)
        self.groupBox_2 = QtWidgets.QGroupBox(Form)
        self.groupBox_2.setGeometry(QtCore.QRect(380, 10, 681, 831))
        self.groupBox_2.setObjectName("groupBox_2")
        self.tableWidget_2 = QtWidgets.QTableWidget(self.groupBox_2)
        self.tableWidget_2.setGeometry(QtCore.QRect(10, 30, 651, 791))
        self.tableWidget_2.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)
        self.tableWidget_2.setObjectName("tableWidget_2")
        self.tableWidget_2.setColumnCount(0)
        self.tableWidget_2.setRowCount(0)
        self.tableWidget_2.horizontalHeader().setCascadingSectionResizes(True)
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
        self.groupBox_7 = QtWidgets.QGroupBox(Form)
        self.groupBox_7.setGeometry(QtCore.QRect(10, 840, 1051, 151))
        self.groupBox_7.setObjectName("groupBox_7")
        self.listWidget = QtWidgets.QListWidget(self.groupBox_7)
        self.listWidget.setGeometry(QtCore.QRect(10, 20, 1021, 121))
        self.listWidget.setObjectName("listWidget")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(1090, 930, 131, 41))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "REGISTER UM"))
        self.groupBox.setTitle(_translate("Form", "Elective"))
        self.tableWidget.setSortingEnabled(True)
        self.groupBox_2.setTitle(_translate("Form", "KoK"))
        self.tableWidget_2.setSortingEnabled(True)
        self.groupBox_3.setTitle(_translate("Form", "Set Timer (sec)"))
        self.pushButton.setText(_translate("Form", "START!!"))
        self.lineEdit.setText(_translate("Form", "5"))
        self.groupBox_4.setTitle(_translate("Form", "Targeted Elective 1"))
        self.groupBox_5.setTitle(_translate("Form", "Targeted Elective 2"))
        self.groupBox_6.setTitle(_translate("Form", "Targeted KoK 1"))
        self.groupBox_7.setTitle(_translate("Form", "Command"))
        self.label.setText(_translate("Form", "A program by\n"
"hongvin"))



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
        iconPathName = os.path.abspath(os.path.join( sys.path[0], "favicon.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
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

engine = pyttsx3.init()
def TTS(text):
    split=" ".join(text)
    engine.say('Found!' + split)
    engine.runAndWait() 

class App(QtWidgets.QMainWindow,Ui_Form):
    def __init__(self):
        super(self.__class__,self).__init__()
        self.setupUi(self)      
        self.pushButton.clicked.connect(self.startengine)
        
    def startengine(self):
        self.listWidget.scrollToBottom()
        self.lineEdit.setEnabled(False)
        timeout=float(self.lineEdit.text())
        
        time_now=time.time()

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
                    TTS(a[i].text)
                if e2==a[i].text:
                    self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Matched found for targeted elective '+a[i].text)
                    e2b=True
                    courses[1]=(a[i].text)
                    TTS(a[i].text)
                    
                k+=1
        except Exception as e:
            print("Error",e)
            self.tableWidget.setEnabled(False)
            self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Error occured at Elective')
            self.lineEdit.setEnabled(True)
        
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
                    TTS(a[i].text)
                self.tableWidget_2.setItem(j,k,QtWidgets.QTableWidgetItem(a[i].text))
                k+=1
        except Exception as e:
            print("Error",e)
            self.tableWidget.setEnabled(False)
            self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: Error occured at KoK')
            self.lineEdit.setEnabled(True)
        self.tableWidget.resizeColumnsToContents()
        self.tableWidget_2.resizeColumnsToContents()
        self.listWidget.scrollToBottom()
        self.listWidget.addItem('['+str(datetime.datetime.now().time())+']: List refreshed')
        if e1b==True or e2b==True or k1b==True:
            w.ShowWindow("Matching course found!","Course Code: {0} {1} {2}".format(*courses))
        
        next_time=time_now+timeout
        t=threading.Timer(next_time-time.time(),self.startengine)
        t.daemon=True
        t.start()
        
            
        

    
def main():  
    app=QtWidgets.QApplication(sys.argv)
    form=App()
    form.show()
    app.exec_()

if __name__ == '__main__':
    main()

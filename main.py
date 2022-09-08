import sys
import datetime
import pprint
import os
import time
import subprocess
import json

from PyQt5 import QtCore, QtGui, QtWidgets, uic
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QInputDialog, QLineEdit, QFileDialog, QMessageBox, \
    QAction,QListView

import common.dataScrape as dataScrape
import common.configManage as configManage

#### ESTABLISH PRE-ENTERED DATA
#
defaultQuotesPath = os.getcwd()
defaultQuotesPathAlert = dataScrape.pathCheck(defaultQuotesPath)
#
# try:
#     companyData =  dataScrape.getCompanyListData()
# except Exception as e:
#     print(f"There was an error retrieving data from the company database google sheet\n{e}")
#
# companies = dataScrape.getCompanyList(companyData["companies"])
# lineItemPresets = dataScrape.getLineItemPresets(companyData)
# descriptionPresets = dataScrape.getServiceDescriptionPresets(companyData)

json_file = "common/dummy_client_info.json"
with open(json_file) as f:
   data = json.load(f)

companyData = data["company data"]

my_company_name = "my company"

companies = [
            "ABC cleaning supplies",
            "John's Hardware",
            "Market Rate Foods",
            "Penn Logistics"
]

lineItemPresets = [
            "General Consulting",
            "Account Audit",
            "Security Audit",
            "Site Inventory"
]

quoteNumber = dataScrape.getLastQuoteNumber(defaultQuotesPath)
defaultTax = 0
defaultDiscount = 0.00
defaultDescription = "Asset Management Services"
defaultQTY = 1.00
defaultUnitPrice = 3000.00
now = datetime.datetime.now()
currentDate = now.strftime("%m/%d/%Y") #now formated to mm/dd/yy

class ComboBox(QtWidgets.QComboBox):
    popupAboutToBeShown = QtCore.pyqtSignal()

    def __init__(self, window, posX,posY):
        super().__init__(window)
        self.posX = posX
        self.posY = posY

    def showPopup(self):
        self.setGeometry(QtCore.QRect(self.posY, self.posX, 480, 41))
        self.popupAboutToBeShown.emit()
        super().showPopup()

    def hidePopup(self):
        super().hidePopup()
        self.setGeometry(QtCore.QRect(self.posY, self.posX, 21, 41))
        self.popupAboutToBeShown.emit()

class Window(QMainWindow):

    def setupUi(self):
        # return gui_config(self,defaultQuotesPathAlert,quoteNumber)
        # MainWindow.setObjectName("MainWindow")
        self.setWindowIcon(QtGui.QIcon('src/images/logo.png'))

        self.resize(916, 750)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        self.setMinimumSize(QtCore.QSize(916, 750))
        self.setMaximumSize(QtCore.QSize(916, 750))

        self.quoteDirLabel = QtWidgets.QLabel(self)
        self.quoteDirLabel.setGeometry(QtCore.QRect(120, 15, 151, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.quoteDirLabel.setFont(font)
        self.quoteDirLabel.setObjectName("quoteDirLabel")

        self.quoteDirPath = QtWidgets.QLineEdit(self)
        self.quoteDirPath.setGeometry(QtCore.QRect(270, 10, 481, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.quoteDirPath.setFont(font)
        self.quoteDirPath.setText("")
        self.quoteDirPath.setAlignment(QtCore.Qt.AlignLeading | QtCore.Qt.AlignLeft | QtCore.Qt.AlignVCenter)
        self.quoteDirPath.setObjectName("quoteDirPath")

        self.quoteDirDialogButton = QtWidgets.QPushButton(self)
        self.quoteDirDialogButton.setGeometry(QtCore.QRect(760, 10, 41, 41))
        self.quoteDirDialogButton.setText("")
        self.quoteDirDialogButton.setObjectName("quoteDirDialogButton")

        self.companyListLabel = QtWidgets.QLabel(self)
        self.companyListLabel.setGeometry(QtCore.QRect(240, 84, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.companyListLabel.setFont(font)
        self.companyListLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.companyListLabel.setObjectName("companyListLabel")

        self.companyList = QtWidgets.QComboBox(self)
        self.companyList.setGeometry(QtCore.QRect(190, 110, 211, 41))
        font = QtGui.QFont()
        font.setPointSize(16)
        font.setKerning(True)
        self.companyList.setFont(font)
        self.companyList.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.companyList.setAutoFillBackground(False)
        self.companyList.setObjectName("companyList")  ############################Company LIST#################+
        self.companyList.addItems(companies)

        self.quoteDirAlertLabel = QtWidgets.QLabel(self)
        self.quoteDirAlertLabel.setGeometry(QtCore.QRect(280, 60, 471, 20))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.quoteDirAlertLabel.setFont(font)
        self.quoteDirAlertLabel.setStyleSheet("color: red;")
        self.quoteDirAlertLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.quoteDirAlertLabel.setObjectName("quoteDirAlertLabel")

        self.dateLabel = QtWidgets.QLabel(self)
        self.dateLabel.setGeometry(QtCore.QRect(460, 84, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.dateLabel.setFont(font)
        self.dateLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.dateLabel.setObjectName("dateLabel")

        self.quoteNumLabel = QtWidgets.QLabel(self)
        self.quoteNumLabel.setGeometry(QtCore.QRect(630, 84, 101, 31))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.quoteNumLabel.setFont(font)
        self.quoteNumLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.quoteNumLabel.setObjectName("quoteNumLabel")

        self.dateEntry = QtWidgets.QLineEdit(self)
        self.dateEntry.setGeometry(QtCore.QRect(450, 110, 121, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.dateEntry.setFont(font)
        self.dateEntry.setAlignment(QtCore.Qt.AlignCenter)
        self.dateEntry.setObjectName("dateEntry")

        self.quoteNumEntry = QtWidgets.QLineEdit(self)
        self.quoteNumEntry.setGeometry(QtCore.QRect(620, 110, 121, 41))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.quoteNumEntry.setFont(font)
        self.quoteNumEntry.setAlignment(QtCore.Qt.AlignCenter)
        self.quoteNumEntry.setObjectName("quoteNumEntry")

        self.headerDivider = QtWidgets.QFrame(self)
        self.headerDivider.setGeometry(QtCore.QRect(-340, 150, 2961, 20))
        self.headerDivider.setFrameShape(QtWidgets.QFrame.HLine)
        self.headerDivider.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.headerDivider.setObjectName("headerDivider")

        self.lineItemTitleLabel = QtWidgets.QLabel(self)
        self.lineItemTitleLabel.setGeometry(QtCore.QRect(220, 180, 141, 31))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.lineItemTitleLabel.setFont(font)
        self.lineItemTitleLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.lineItemTitleLabel.setObjectName("lineItemTitleLabel")

        self.liDiscountEntry = QtWidgets.QLineEdit(self)
        self.liDiscountEntry.setGeometry(QtCore.QRect(552, 200, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDiscountEntry.setFont(font)
        self.liDiscountEntry.setAlignment(QtCore.Qt.AlignCenter)
        self.liDiscountEntry.setObjectName("liDiscountEntry")

        self.liDiscountLabel = QtWidgets.QLabel(self)
        self.liDiscountLabel.setGeometry(QtCore.QRect(570, 160, 91, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDiscountLabel.setFont(font)
        self.liDiscountLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liDiscountLabel.setObjectName("liDiscountLabel")

        self.liTaxRateEntry = QtWidgets.QLineEdit(self)
        self.liTaxRateEntry.setGeometry(QtCore.QRect(712, 200, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTaxRateEntry.setFont(font)
        self.liTaxRateEntry.setAlignment(QtCore.Qt.AlignCenter)
        self.liTaxRateEntry.setObjectName("liTaxRateEntry")

        self.liTaxRateLabel = QtWidgets.QLabel(self)
        self.liTaxRateLabel.setGeometry(QtCore.QRect(720, 160, 111, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTaxRateLabel.setFont(font)
        self.liTaxRateLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liTaxRateLabel.setObjectName("liTaxRateLabel")

        self.liTitleDivider = QtWidgets.QFrame(self)
        self.liTitleDivider.setGeometry(QtCore.QRect(40, 230, 821, 20))
        self.liTitleDivider.setFrameShape(QtWidgets.QFrame.HLine)
        self.liTitleDivider.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.liTitleDivider.setObjectName("liTitleDivider")

        ### first group box#######
        self.groupBox = QtWidgets.QGroupBox(self)
        self.groupBox.setGeometry(QtCore.QRect(40, 286, 831, 61))
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.liDesc1 = QtWidgets.QLineEdit(self.groupBox)
        self.liDesc1.setGeometry(QtCore.QRect(10, 10, 481, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDesc1.setFont(font)
        self.liDesc1.setAlignment(QtCore.Qt.AlignCenter)
        self.liDesc1.setObjectName("liDesc1")
        self.liQTY1 = QtWidgets.QLineEdit(self.groupBox)
        self.liQTY1.setGeometry(QtCore.QRect(520, 10, 61, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liQTY1.setFont(font)
        self.liQTY1.setAlignment(QtCore.Qt.AlignCenter)
        self.liQTY1.setObjectName("liQTY1")
        self.liUP1 = QtWidgets.QLineEdit(self.groupBox)
        self.liUP1.setGeometry(QtCore.QRect(600, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liUP1.setFont(font)
        self.liUP1.setAlignment(QtCore.Qt.AlignCenter)
        self.liUP1.setObjectName("liUP1")
        self.liTOT1 = QtWidgets.QLineEdit(self.groupBox)
        self.liTOT1.setGeometry(QtCore.QRect(720, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTOT1.setFont(font)
        self.liTOT1.setAlignment(QtCore.Qt.AlignCenter)
        self.liTOT1.setReadOnly(True)
        self.liTOT1.setObjectName("liTOT1")
        self.liDescLabel = QtWidgets.QLabel(self)
        self.liDescLabel.setGeometry(QtCore.QRect(220, 250, 141, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.liDescLabel.setFont(font)
        self.liDescLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liDescLabel.setObjectName("liDescLabel")
        self.liQTYLabel = QtWidgets.QLabel(self)
        self.liQTYLabel.setGeometry(QtCore.QRect(560, 250, 51, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.liQTYLabel.setFont(font)
        self.liQTYLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liQTYLabel.setObjectName("liQTYLabel")
        self.liUPLabel = QtWidgets.QLabel(self)
        self.liUPLabel.setGeometry(QtCore.QRect(640, 240, 91, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liUPLabel.setFont(font)
        self.liUPLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liUPLabel.setObjectName("liUPLabel")
        self.liTOTLabel = QtWidgets.QLabel(self)
        self.liTOTLabel.setGeometry(QtCore.QRect(760, 240, 91, 51))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTOTLabel.setFont(font)
        self.liTOTLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.liTOTLabel.setObjectName("liTOTLabel")
        ### END FIRST GROUP BOX

        self.groupBox_2 = QtWidgets.QGroupBox(self)
        self.groupBox_2.setGeometry(QtCore.QRect(40, 346, 831, 61))
        self.groupBox_2.setTitle("")
        self.groupBox_2.setObjectName("groupBox_2")
        self.liDesc2 = QtWidgets.QLineEdit(self.groupBox_2)
        self.liDesc2.setGeometry(QtCore.QRect(10, 10, 481, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDesc2.setFont(font)
        self.liDesc2.setText("")
        self.liDesc2.setAlignment(QtCore.Qt.AlignCenter)
        self.liDesc2.setObjectName("liDesc2")
        self.liQTY2 = QtWidgets.QLineEdit(self.groupBox_2)
        self.liQTY2.setGeometry(QtCore.QRect(520, 10, 61, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liQTY2.setFont(font)
        self.liQTY2.setText("")
        self.liQTY2.setAlignment(QtCore.Qt.AlignCenter)
        self.liQTY2.setObjectName("liQTY2")
        self.liUP2 = QtWidgets.QLineEdit(self.groupBox_2)
        self.liUP2.setGeometry(QtCore.QRect(600, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liUP2.setFont(font)
        self.liUP2.setText("")
        self.liUP2.setAlignment(QtCore.Qt.AlignCenter)
        self.liUP2.setObjectName("liUP2")
        self.liTOT2 = QtWidgets.QLineEdit(self.groupBox_2)
        self.liTOT2.setGeometry(QtCore.QRect(720, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTOT2.setFont(font)
        self.liTOT2.setText("")
        self.liTOT2.setAlignment(QtCore.Qt.AlignCenter)
        self.liTOT2.setReadOnly(True)
        self.liTOT2.setObjectName("liTOT2")
        self.groupBox_3 = QtWidgets.QGroupBox(self)
        self.groupBox_3.setGeometry(QtCore.QRect(40, 406, 831, 61))
        self.groupBox_3.setTitle("")
        self.groupBox_3.setObjectName("groupBox_3")
        self.liDesc3 = QtWidgets.QLineEdit(self.groupBox_3)
        self.liDesc3.setGeometry(QtCore.QRect(10, 10, 481, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDesc3.setFont(font)
        self.liDesc3.setText("")
        self.liDesc3.setAlignment(QtCore.Qt.AlignCenter)
        self.liDesc3.setObjectName("liDesc3")
        self.liQTY3 = QtWidgets.QLineEdit(self.groupBox_3)
        self.liQTY3.setGeometry(QtCore.QRect(520, 10, 61, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liQTY3.setFont(font)
        self.liQTY3.setText("")
        self.liQTY3.setAlignment(QtCore.Qt.AlignCenter)
        self.liQTY3.setObjectName("liQTY3")
        self.liUP3 = QtWidgets.QLineEdit(self.groupBox_3)
        self.liUP3.setGeometry(QtCore.QRect(600, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liUP3.setFont(font)
        self.liUP3.setText("")
        self.liUP3.setAlignment(QtCore.Qt.AlignCenter)
        self.liUP3.setObjectName("liUP3")
        self.liTOT3 = QtWidgets.QLineEdit(self.groupBox_3)
        self.liTOT3.setGeometry(QtCore.QRect(720, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTOT3.setFont(font)
        self.liTOT3.setText("")
        self.liTOT3.setAlignment(QtCore.Qt.AlignCenter)
        self.liTOT3.setReadOnly(True)
        self.liTOT3.setObjectName("liTOT3")
        self.groupBox_4 = QtWidgets.QGroupBox(self)
        self.groupBox_4.setGeometry(QtCore.QRect(40, 466, 831, 61))
        self.groupBox_4.setTitle("")
        self.groupBox_4.setObjectName("groupBox_4")
        self.liDesc4 = QtWidgets.QLineEdit(self.groupBox_4)
        self.liDesc4.setGeometry(QtCore.QRect(10, 10, 481, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liDesc4.setFont(font)
        self.liDesc4.setText("")
        self.liDesc4.setAlignment(QtCore.Qt.AlignCenter)
        self.liDesc4.setObjectName("liDesc4")
        self.liQTY4 = QtWidgets.QLineEdit(self.groupBox_4)
        self.liQTY4.setGeometry(QtCore.QRect(520, 10, 61, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liQTY4.setFont(font)
        self.liQTY4.setText("")
        self.liQTY4.setAlignment(QtCore.Qt.AlignCenter)
        self.liQTY4.setObjectName("liQTY4")
        self.liUP4 = QtWidgets.QLineEdit(self.groupBox_4)
        self.liUP4.setGeometry(QtCore.QRect(600, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liUP4.setFont(font)
        self.liUP4.setText("")
        self.liUP4.setAlignment(QtCore.Qt.AlignCenter)
        self.liUP4.setObjectName("liUP4")
        self.liTOT4 = QtWidgets.QLineEdit(self.groupBox_4)
        self.liTOT4.setGeometry(QtCore.QRect(720, 10, 91, 41))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.liTOT4.setFont(font)
        self.liTOT4.setText("")
        self.liTOT4.setAlignment(QtCore.Qt.AlignCenter)
        self.liTOT4.setReadOnly(True)
        self.liTOT4.setObjectName("liTOT4")
        # self.groupBox_5 = QtWidgets.QGroupBox(self)
        # self.groupBox_5.setGeometry(QtCore.QRect(40, 610, 831, 61))
        # self.groupBox_5.setTitle("")
        # self.groupBox_5.setObjectName("groupBox_5")
        # self.liDesc5 = QtWidgets.QLineEdit(self.groupBox_5)
        # self.liDesc5.setGeometry(QtCore.QRect(10, 10, 481, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liDesc5.setFont(font)
        # self.liDesc5.setText("")
        # self.liDesc5.setAlignment(QtCore.Qt.AlignCenter)
        # self.liDesc5.setObjectName("liDesc5")
        # self.liQTY5 = QtWidgets.QLineEdit(self.groupBox_5)
        # self.liQTY5.setGeometry(QtCore.QRect(520, 10, 61, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liQTY5.setFont(font)
        # self.liQTY5.setText("")
        # self.liQTY5.setAlignment(QtCore.Qt.AlignCenter)
        # self.liQTY5.setObjectName("liQTY5")
        # self.liUP5 = QtWidgets.QLineEdit(self.groupBox_5)
        # self.liUP5.setGeometry(QtCore.QRect(610, 10, 91, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liUP5.setFont(font)
        # self.liUP5.setText("")
        # self.liUP5.setAlignment(QtCore.Qt.AlignCenter)
        # self.liUP5.setObjectName("liUP5")
        # self.liTOT5 = QtWidgets.QLineEdit(self.groupBox_5)
        # self.liTOT5.setGeometry(QtCore.QRect(720, 10, 91, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liTOT5.setFont(font)
        # self.liTOT5.setText("")
        # self.liTOT5.setAlignment(QtCore.Qt.AlignCenter)
        # self.liTOT5.setReadOnly(True)
        # self.liTOT5.setObjectName("liTOT5")
        # self.groupBox_6 = QtWidgets.QGroupBox(self)
        # self.groupBox_6.setGeometry(QtCore.QRect(40, 670, 831, 61))
        # self.groupBox_6.setTitle("")
        # self.groupBox_6.setObjectName("groupBox_6")
        # self.liDesc6 = QtWidgets.QLineEdit(self.groupBox_6)
        # self.liDesc6.setGeometry(QtCore.QRect(10, 10, 481, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liDesc6.setFont(font)
        # self.liDesc6.setText("")
        # self.liDesc6.setAlignment(QtCore.Qt.AlignCenter)
        # self.liDesc6.setObjectName("liDesc6")
        # self.liQTY6 = QtWidgets.QLineEdit(self.groupBox_6)
        # self.liQTY6.setGeometry(QtCore.QRect(520, 10, 61, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liQTY6.setFont(font)
        # self.liQTY6.setText("")
        # self.liQTY6.setAlignment(QtCore.Qt.AlignCenter)
        # self.liQTY6.setObjectName("liQTY6")
        # self.liUP6 = QtWidgets.QLineEdit(self.groupBox_6)
        # self.liUP6.setGeometry(QtCore.QRect(610, 10, 91, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liUP6.setFont(font)
        # self.liUP6.setText("")
        # self.liUP6.setAlignment(QtCore.Qt.AlignCenter)
        # self.liUP6.setObjectName("liUP6")
        # self.liTOT6 = QtWidgets.QLineEdit(self.groupBox_6)
        # self.liTOT6.setGeometry(QtCore.QRect(720, 10, 91, 41))
        # font = QtGui.QFont()
        # font.setPointSize(12)
        # self.liTOT6.setFont(font)
        # self.liTOT6.setText("")
        # self.liTOT6.setAlignment(QtCore.Qt.AlignCenter)
        # self.liTOT6.setReadOnly(True)
        # self.liTOT6.setObjectName("liTOT6")
        font = QtGui.QFont()
        font.setPointSize(13)
        font.setKerning(True)

        combo1Pos = (296, 50)
        self.combo1 = ComboBox(self, combo1Pos[0], combo1Pos[1])
        self.combo1.setGeometry(QtCore.QRect(combo1Pos[1], combo1Pos[0], 21, 41))
        self.combo1.addItems(lineItemPresets)
        self.combo1.setFont(font)

        combo2Pos = (356, 50)
        self.combo2 = ComboBox(self, combo2Pos[0], combo2Pos[1])
        self.combo2.setGeometry(QtCore.QRect(combo2Pos[1], combo2Pos[0], 21, 41))
        self.combo2.addItems(lineItemPresets)
        self.combo2.setFont(font)

        combo3Pos = (416, 50)
        self.combo3 = ComboBox(self, combo3Pos[0], combo3Pos[1])
        self.combo3.setGeometry(QtCore.QRect(combo3Pos[1], combo3Pos[0], 21, 41))
        self.combo3.addItems(lineItemPresets)
        self.combo3.setFont(font)

        combo4Pos = (476, 50)
        self.combo4 = ComboBox(self, combo4Pos[0], combo4Pos[1])
        self.combo4.setGeometry(QtCore.QRect(combo4Pos[1], combo4Pos[0], 21, 41))
        self.combo4.addItems(lineItemPresets)
        self.combo4.setFont(font)
        #
        # combo5Pos = (620, 50)
        # self.combo5 = ComboBox(self, combo5Pos[0], combo5Pos[1])
        # self.combo5.setGeometry(QtCore.QRect(combo5Pos[1], combo5Pos[0], 21, 41))
        # self.combo5.addItems(lineItemPresets)
        # self.combo5.setFont(font)
        #
        # combo6Pos = (680, 50)
        # self.combo6 = ComboBox(self, combo6Pos[0], combo6Pos[1])
        # self.combo6.setGeometry(QtCore.QRect(combo6Pos[1], combo6Pos[0], 21, 41))
        # self.combo6.addItems(lineItemPresets)
        # self.combo6.setFont(font)
        # self.combo6.setObjectName("combo6")

        # self.serviceDesc = QtWidgets.QComboBox(self)
        # self.serviceDesc.setObjectName(u"serviceDesc")
        # self.serviceDesc.setGeometry(QtCore.QRect(40, 740, 501, 61))
        # # self.serviceDesc.addItems(descriptionPresets)
        # self.serviceDesc.setModel(QtCore.QStringListModel(descriptionPresets))
        # # The popup widget is QListView
        # listView = QListView()
        # # Turn On the word wrap
        # listView.setWordWrap(True)
        # # set popup view widget into the combo box
        # self.serviceDesc.setView(listView)

        font = QtGui.QFont()
        font.setPointSize(16)
        self.grandTotalLabel = QtWidgets.QLabel(self)
        self.grandTotalLabel.setObjectName("grandTotalLabel")
        self.grandTotalLabel.setGeometry(QtCore.QRect(580, 540, 141, 41))
        self.grandTotalLabel.setFont(font)
        self.grandTotalLabel.setAlignment(QtCore.Qt.AlignCenter)

        font = QtGui.QFont()
        font.setPointSize(14)
        self.liGrandTotal = QtWidgets.QLineEdit(self)
        self.liGrandTotal.setObjectName("liGrandTotal")
        self.liGrandTotal.setGeometry(QtCore.QRect(750, 540, 111, 41))
        self.liGrandTotal.setFont(font)
        self.liGrandTotal.setAlignment(QtCore.Qt.AlignCenter)
        self.liGrandTotal.setReadOnly(True)

        self.liTitleDivider_3 = QtWidgets.QFrame(self)
        self.liTitleDivider_3.setObjectName("liTitleDivider_3")
        self.liTitleDivider_3.setGeometry(QtCore.QRect(40, 580, 821, 20))
        self.liTitleDivider_3.setFrameShape(QtWidgets.QFrame.HLine)
        self.liTitleDivider_3.setFrameShadow(QtWidgets.QFrame.Sunken)


        # self.liTitleDivider = QtWidgets.QFrame(self)
        # self.liTitleDivider.setGeometry(QtCore.QRect(40, 320, 821, 20))
        # self.liTitleDivider.setFrameShape(QtWidgets.QFrame.HLine)
        # self.liTitleDivider.setFrameShadow(QtWidgets.QFrame.Sunken)
        # self.liTitleDivider.setObjectName("liTitleDivider")
        ##### END LINE ITEM SECTION #######

        ##### START FOOTER SECTION ############
        self.generateButton = QtWidgets.QPushButton(self)
        self.generateButton.setGeometry(QtCore.QRect(290, 610, 341, 51))
        font = QtGui.QFont()
        font.setPointSize(16)
        self.generateButton.setFont(font)
        self.generateButton.setObjectName("generateButton")

        font = QtGui.QFont()
        font.setPointSize(12)

        self.openPDFOnGenerate = QtWidgets.QCheckBox(self)
        self.openPDFOnGenerate.setGeometry(QtCore.QRect(640, 600, 251, 41))
        self.openPDFOnGenerate.setChecked(True)
        self.openPDFOnGenerate.setFont(font)

        self.saveEntriesForNextTime = QtWidgets.QCheckBox(self)
        self.saveEntriesForNextTime.setGeometry(QtCore.QRect(640, 620, 251, 41))
        self.saveEntriesForNextTime.setChecked(False)
        self.saveEntriesForNextTime.setFont(font)

        self.progressLabel = QtWidgets.QLabel(self)
        self.progressLabel.setGeometry(QtCore.QRect(20, 665, 881, 51))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.progressLabel.setFont(font)
        self.progressLabel.setStyleSheet("color: green;")
        self.progressLabel.setAlignment(QtCore.Qt.AlignCenter)
        self.progressLabel.setObjectName("progressLabel")



        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Company Quote Generator"))
        self.quoteDirLabel.setText(_translate("MainWindow", "Quotes Directory"))
        self.quoteDirPath.setText(_translate("MainWindow", defaultQuotesPath))  ####### DEFAULT QUOTE DIRECTORY PATH####
        self.quoteDirAlertLabel.setText(_translate("MainWindow", defaultQuotesPathAlert))
        self.companyListLabel.setText(_translate("MainWindow", "Company"))
        self.dateLabel.setText(_translate("MainWindow", "Date"))
        self.quoteNumLabel.setText(_translate("MainWindow", "Quote#"))
        self.dateEntry.setText(_translate("MainWindow", currentDate))  ######## DATE ENTRY ####
        self.quoteNumEntry.setText(_translate("MainWindow", f"{quoteNumber}"))  ######## QUOTE NUMBER ###
        self.lineItemTitleLabel.setText(_translate("MainWindow", "LINE ITEMS"))
        # self.liDesc1.setText(_translate("MainWindow", defaultDescription))  ######## default lidesc entry ###
        # self.liQTY1.setText(_translate("MainWindow", "{:.2f}".format(defaultQTY)))  ######## default liQTY entry ###
        # self.liUP1.setText(
        #     _translate("MainWindow", "{:.2f}".format(defaultUnitPrice)))  ######## default liUnit Price entry ###
        # self.liTOT1.setText(_translate("MainWindow", "{:.2f}".format(
        #     defaultQTY * defaultUnitPrice)))  ######## default liTOTAL entry ###
        self.liDescLabel.setText(_translate("MainWindow", "Description"))
        self.liQTYLabel.setText(_translate("MainWindow", "QTY"))
        self.liUPLabel.setText(_translate("MainWindow", "Unit Price"))
        self.liTOTLabel.setText(_translate("MainWindow", "TOTAL"))
        self.liDiscountEntry.setText(_translate("MainWindow", "{:.2f}".format(defaultDiscount)))
        self.liDiscountLabel.setText(_translate("MainWindow", "Discount"))
        self.liTaxRateEntry.setText(_translate("MainWindow", f"{defaultTax}"))
        self.liTaxRateLabel.setText(_translate("MainWindow", "Tax Rate %"))


        self.grandTotalLabel.setText(_translate("MainWindow", "Grand Total"))
        self.openPDFOnGenerate.setText(_translate("MainWindow", "Open pdf file on creation"))
        self.saveEntriesForNextTime.setText(_translate("MainWindow", "Save entries for next open"))
        self.generateButton.setText(_translate("MainWindow", "GENERATE QUOTE PDF"))
        self.progressLabel.setText(_translate("MainWindow", ""))

        #### Button click actions
        self.quoteDirDialogButton.clicked.connect(self.dirDialogButton)  ##### Directory change button ###
        self.generateButton.clicked.connect(self.genButtonClicked)  ##### Generate Button change button ###

        #### Alert Label actions
        self.quoteDirPath.textChanged.connect(lambda: self.pathCheck(self.quoteDirPath.text().strip()))

        #CONNECT Description drop down boxes to corresponding line edit
        self.combo1.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc1, self.combo1))
        self.combo2.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc2, self.combo2))
        self.combo3.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc3, self.combo3))
        self.combo4.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc4, self.combo4))
        # self.combo5.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc5, self.combo5))
        # self.combo6.activated.connect(lambda: self.dropDownToLineEdit(self.liDesc6, self.combo6))

        #### SET ALL TOTAL FIELDS TO HAVE GREY BACKROUND TO INDICATE READ ONLY
        self.setLineEditColor(self.liTOT1, "darkgray")
        self.setLineEditColor(self.liTOT2, "darkgray")
        self.setLineEditColor(self.liTOT3, "darkgray")
        self.setLineEditColor(self.liTOT4, "darkgray")
        # self.setLineEditColor(self.liTOT5, "darkgray")
        # self.setLineEditColor(self.liTOT6, "darkgray")
        self.setLineEditColor(self.liGrandTotal, "darkgray")


        #### On line item text being changed, run multiplyQTYbyUnit func
        self.liDiscountEntry.textChanged.connect(self.setGrandTotal)

        self.liTaxRateEntry.textChanged.connect(self.setGrandTotal)

        self.liQTY1.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY1.text(), self.liUP1.text(), self.liTOT1))
        self.liUP1.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY1.text(), self.liUP1.text(), self.liTOT1))

        self.liQTY2.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY2.text(), self.liUP2.text(), self.liTOT2))
        self.liUP2.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY2.text(), self.liUP2.text(), self.liTOT2))

        self.liQTY3.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY3.text(), self.liUP3.text(), self.liTOT3))
        self.liUP3.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY3.text(), self.liUP3.text(), self.liTOT3))

        self.liQTY4.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY4.text(), self.liUP4.text(), self.liTOT4))
        self.liUP4.textChanged.connect(
            lambda: self.multiplyQTYbyUnit(self.liQTY4.text(), self.liUP4.text(), self.liTOT4))

        # self.liQTY5.textChanged.connect(
        #     lambda: self.multiplyQTYbyUnit(self.liQTY5.text(), self.liUP5.text(), self.liTOT5))
        # self.liUP5.textChanged.connect(
        #     lambda: self.multiplyQTYbyUnit(self.liQTY5.text(), self.liUP5.text(), self.liTOT5))
        #
        # self.liQTY6.textChanged.connect(
        #     lambda: self.multiplyQTYbyUnit(self.liQTY6.text(), self.liUP6.text(), self.liTOT6))
        # self.liUP6.textChanged.connect(
        #     lambda: self.multiplyQTYbyUnit(self.liQTY6.text(), self.liUP6.text(), self.liTOT6))

        # connect first line item to service description paragraph
        # self.liDesc1.textChanged.connect(self.changeDescriptionParagraph)

     # when generate button is clicked
    def genButtonClicked(self):
        self.progressLabel.setText("Processing...")
        self.progressLabel.repaint()

        grand_total = self.liGrandTotal.text().strip()
        #get all inputs from GUI
        userInputs = self.parseUserInputs()

        #check for valid inputs
        if grand_total == "ERROR" or grand_total == "":
            self.messageBox("Please input valid inputs for QTY and Unit Price")
            return

        ##check if quote path exists
        if not dataScrape.doesPathExist(userInputs['quotes dir']):
            ### PLACE ERROR MESSAGE HERE
            # self.statusbar.showMessage("ERROR - Directory does not exist")
            self.messageBox(
                "The path entered does not exist. Please create it or change the path to an existing directory")
            return

        ## check if quote name already exist in quote path
        try:
            isNewNumber = dataScrape.hasQuoteNumberBeenUsed(userInputs['quotes dir'], userInputs['quote number'])
        except Exception as e:
            print(e)

        if isNewNumber:
            ### PLACE ERROR MESSAGE HERE
            # self.statusbar.showMessage("ERROR - Quote number already exists")
            self.messageBox("This quote number already exist at the given path. Please change the quote number")
            return

        ## get data for company and bundle with userInputs dictionary
        userInputs['company data'] = dataScrape.getCompanyInfo(userInputs['company'], companyData)

        ## establish quote file name
        quoteFileName = f"S{userInputs['quote number']} {my_company_name} Quote - {userInputs['company data']['Full Name']}"

        ## send all data to spreadsheet template and create new sheet
        try:
            filesCreated = dataScrape.writeToQuoteTemplate(my_company_name,userInputs)
            if not filesCreated:
                self.messageBox("ERROR - Script timed out. xlsx file not created.")
        except Exception as e:
            self.messageBox(f"Something went wrong. Quote files were not created\n{e}")
            return

        # Job Complete Label
        self.progressLabel.setText(f"{quoteFileName} xlsx and pdf files\ncreated @ \"{userInputs['quotes dir']}\"")

        # Auto increment quote number field
        self.quoteNumEntry.setText(f"{dataScrape.getLastQuoteNumber(userInputs['quotes dir'])}")

        # Evaluate open file checkbox
        print(filesCreated)
        self.evalOpenFileCheckbox(filesCreated["pdfFileName"])

    def parseUserInputs(self):
        inputs = {}
        inputs["quotes dir"] = self.quoteDirPath.text().strip()
        inputs["quote number"] = self.quoteNumEntry.text().upper().strip()
        inputs["company"] = self.companyList.currentText()
        inputs["date"] = self.dateEntry.text().strip()
        inputs["discount"] = self.liDiscountEntry.text().strip()
        inputs["tax"] = self.liTaxRateEntry.text().strip()

        lineItemInputs = [
            (self.liDesc1, self.liQTY1, self.liUP1),
            (self.liDesc2, self.liQTY2, self.liUP2),
            (self.liDesc3, self.liQTY3, self.liUP3),
            (self.liDesc4, self.liQTY4, self.liUP4),
            # (self.liDesc5, self.liQTY5, self.liUP5),
            # (self.liDesc6, self.liQTY6, self.liUP6)
        ]

        lineItemText = []

        for entry in lineItemInputs:
            row = {}
            row["description"] = entry[0].text().strip()
            row["quantity"] = entry[1].text().strip()
            row["unit price"] = entry[2].text().strip()
            lineItemText.append(row)

        inputs["line items"] = lineItemText
        return inputs

    def dirDialogButton(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Directory")
        if path:
            if sys.platform == "win32":
                adjustedPath = path.replace("/", "\\")
                self.quoteDirPath.setText(adjustedPath)
            else:
                self.quoteDirPath.setText(path)


    def multiplyQTYbyUnit(self, qty, up, tot):

        qty, up = qty.strip(), up.strip()

        if (not qty) and (not up):  # if both values are blank
            tot.setText("")
            self.setLineEditColor(tot, "darkgray")
        else:

            try:
                qty = float(qty)
                up = float(up)
                tot.setText("{:.2f}".format(qty * up))
                self.setLineEditColor(tot, "darkgray")

            except:
                tot.setText("ERROR")
                self.setLineEditColor(tot, "red")
        self.setGrandTotal() #Change grand total value on every edit

    def pathCheck(self, directory):
        if not dataScrape.doesPathExist(directory):
            self.quoteDirAlertLabel.setText("THIS PATH DOES NOT EXIST")

        elif not dataScrape.anyQuotesAtThisPath(directory):
            self.quoteDirAlertLabel.setText("THERE ARE NO QUOTES SAVED AT THIS PATH")
        else:
            self.quoteNumEntry.setText(str(dataScrape.getLastQuoteNumber(directory)))
            self.quoteDirAlertLabel.setText("")

    def setLineEditColor(self, input, color):
        input.setStyleSheet("QLineEdit"
                            "{"
                            f"background : {color};"
                            "}")

    def dropDownToLineEdit(self, lineEditObj, comboBoxObj):
        lineEditObj.setText(comboBoxObj.currentText())

    def setGrandTotal(self):
        try:
            discount = 0 if self.liDiscountEntry.text().strip() == "" else float(self.liDiscountEntry.text())
            taxRate = 0 if self.liTaxRateEntry.text().strip() == "" else float(self.liTaxRateEntry.text())
            tot1 = 0 if self.liTOT1.text().strip() == "" else float(self.liTOT1.text())
            tot2 = 0 if self.liTOT2.text().strip() == "" else float(self.liTOT2.text())
            tot3 = 0 if self.liTOT3.text().strip() == "" else float(self.liTOT3.text())
            tot4 = 0 if self.liTOT4.text().strip() == "" else float(self.liTOT4.text())
            # tot5 = 0 if self.liTOT5.text().strip() == "" else float(self.liTOT5.text())
            # tot6 = 0 if self.liTOT6.text().strip() == "" else float(self.liTOT6.text())
            taxConverted = taxRate / 100
            subTotal = (tot1 + tot2 + tot3 + tot4)# + tot5 + tot6)
            subLessDis = (subTotal - discount)
            taxTotal = float("{:.2f}".format(subLessDis * taxConverted))
            grandTotal = subLessDis + taxTotal
            self.liGrandTotal.setText(str(grandTotal))
            self.setLineEditColor(self.liGrandTotal, "darkgray")

        except Exception as e:
            self.liGrandTotal.setText("ERROR")
            self.setLineEditColor(self.liGrandTotal, "red")

    def changeDescriptionParagraph(self):
        if self.liDesc1.text().strip() in lineItemPresets:
            descriptionIndex = lineItemPresets.index(self.liDesc1.text().strip())
            self.serviceDesc.setCurrentIndex(descriptionIndex)
            print("its here")

    def messageBox(self, message):
        from PyQt5.QtGui import QIcon, QPixmap  # for loading local images
        import random
        cwd = os.getcwd()

        try:
            msg = QMessageBox()
            msg.setWindowIcon(QtGui.QIcon('src/images/logo.png'))
            msg.setWindowTitle("ALERT")
            msg.setText(message)
            self.progressLabel.setText("")

            x = msg.exec()
        except Exception as e:
            print(e)

    def evalOpenFileCheckbox(self,file):
        if self.openPDFOnGenerate.isChecked():
            if sys.platform == "win32": # if running app on windows
                # subprocess.run(file, shell=True) #holds up file generation module til after pdf is closed
                os.startfile(file)
            else:
                subprocess.run(file)

    def evalSaveInputsCheckbox(self):
        cwd = os.getcwd()
        if self.saveEntriesForNextTime.isChecked():
            print("writing config file")
            configManage.onClose(self.parseUserInputs())
        else:
            if dataScrape.doesPathExist(os.path.join(cwd,"config.ini")):
                print("config file exists, deleting now")
                dataScrape.deleteFile("config.ini")
            print("no config file")

    def closeEvent(self, event):
        self.evalSaveInputsCheckbox()
        # if jpype.isJVMStarted():
        #     # print("shutting down JVM")
        #     jpype.shutdownJVM()



def main():
    app = QApplication(sys.argv)
    QtGui.QFontDatabase.addApplicationFont('src/fonts/Gilroy-ExtraBold.otf')
    QtGui.QFontDatabase.addApplicationFont('src/fonts/Gilroy-Bold.ttf')
    QtGui.QFontDatabase.addApplicationFont('src/fonts/Gilroy-Light.ttf')
    qss_file = open('common/styles.qss').read()
    app.setStyleSheet(qss_file)
    win = Window()
    win.setupUi()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
# -*- coding: utf-8 -*-

from os import getcwd

# Form implementation generated from reading ui file 'Result.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!
from PyQt5 import QtCore, QtGui, QtWidgets, QtWebEngineWidgets
from PyQt5.QtCore import QUrl

program_directory = getcwd()[:getcwd().index('Design')]
image_directory = program_directory + 'Design Image\\'

save_token = 0

width = 1500
height = 950
contents_height = 260


class Ui_Form(QtWidgets.QMainWindow):  # 원래 object
    def setupUi(self):
        self.setObjectName("Form")
        self.resize(width, height)
        self.setMinimumSize(QtCore.QSize(width, height))
        self.setMaximumSize(QtCore.QSize(width, height))
        self.setStyleSheet("background-color: rgb(255, 255, 255);")
        self.UpperLine = QtWidgets.QLabel(self)
        self.UpperLine.setGeometry(QtCore.QRect(0, 0, width, 15))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.UpperLine.sizePolicy().hasHeightForWidth())
        self.UpperLine.setSizePolicy(sizePolicy)
        self.UpperLine.setStyleSheet("background-color: rgb(0, 0, 77);")
        self.UpperLine.setText("")
        self.UpperLine.setObjectName("UpperLine")
        self.Title = QtWidgets.QLabel(self)
        title_width = 1200
        title_height = 75
        self.Title.setGeometry(QtCore.QRect(int((width - title_width) / 2), 120, title_width, title_height))
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(36)
        font.setBold(True)
        font.setWeight(75)
        self.Title.setFont(font)
        self.Title.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.Title.setAlignment(QtCore.Qt.AlignCenter)
        self.Title.setObjectName("Title")
        self.DistinctLine = QtWidgets.QFrame(self)
        distinctline_width = width - 100
        self.DistinctLine.setGeometry(QtCore.QRect(int((width - distinctline_width) / 2), 250, distinctline_width, 20))
        self.DistinctLine.setLineWidth(3)
        self.DistinctLine.setFrameShape(QtWidgets.QFrame.HLine)
        self.DistinctLine.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.DistinctLine.setObjectName("DistinctLine")
        self.CompletedSemester = QtWidgets.QLabel(self)
        cpltd_semester_width = 500
        self.CompletedSemester.setGeometry(
            QtCore.QRect(int((width - cpltd_semester_width) / 2), 190, cpltd_semester_width, 31))
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(14)
        self.CompletedSemester.setFont(font)
        self.CompletedSemester.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.CompletedSemester.setAlignment(QtCore.Qt.AlignCenter)
        self.CompletedSemester.setObjectName("CompletedSemester")
        self.scrollArea = QtWidgets.QScrollArea(self)
        self.scrollArea.setGeometry(QtCore.QRect(0, contents_height, width, height - contents_height))
        self.scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.scrollArea.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOff)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")
        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 882, 741))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")
        self.gridLayout = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout.setContentsMargins(50, 0, 30, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.CreditAndMission = QtWebEngineWidgets.QWebEngineView(self.scrollAreaWidgetContents)
        self.CreditAndMission.setMinimumSize(QtCore.QSize(int((width - 130) / 2), 600))
        self.CreditAndMission.setMaximumSize(QtCore.QSize(int((width - 130) / 2), 600))
        self.CreditAndMission.setUrl(QtCore.QUrl("about:blank"))
        self.CreditAndMission.setObjectName("CreditAndMission")
        self.horizontalLayout.addWidget(self.CreditAndMission)
        spacerItem = QtWidgets.QSpacerItem(48, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.CreditGraph = QtWebEngineWidgets.QWebEngineView(self.scrollAreaWidgetContents)
        self.CreditGraph.setMinimumSize(QtCore.QSize(int((width - 130) / 2), 600))
        self.CreditGraph.setMaximumSize(QtCore.QSize(int((width - 130) / 2), 600))
        self.CreditGraph.setUrl(QtCore.QUrl("about:blank"))
        self.CreditGraph.setObjectName("CreditGraph")
        self.horizontalLayout.addWidget(self.CreditGraph)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.Line1 = QtWidgets.QFrame(self.scrollAreaWidgetContents)
        self.Line1.setFrameShape(QtWidgets.QFrame.HLine)
        self.Line1.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.Line1.setObjectName("Line1")
        self.verticalLayout.addWidget(self.Line1)
        self.CompletedMission = QtWebEngineWidgets.QWebEngineView(self.scrollAreaWidgetContents)
        self.CompletedMission.setMinimumSize(QtCore.QSize(width - 200, 500))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        self.CompletedMission.setSizePolicy(sizePolicy)
        self.CompletedMission.setUrl(QtCore.QUrl("about:blank"))
        self.CompletedMission.setObjectName("CompletedMission")
        self.verticalLayout.addWidget(self.CompletedMission)
        self.Line2 = QtWidgets.QFrame(self.scrollAreaWidgetContents)
        self.Line2.setFrameShape(QtWidgets.QFrame.HLine)
        self.Line2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.Line2.setObjectName("Line2")
        self.verticalLayout.addWidget(self.Line2)
        self.IncompletedMission = QtWebEngineWidgets.QWebEngineView(self.scrollAreaWidgetContents)
        self.IncompletedMission.setMinimumSize(QtCore.QSize(width - 200, 500))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Fixed)
        self.IncompletedMission.setSizePolicy(sizePolicy)
        self.IncompletedMission.setUrl(QtCore.QUrl("about:blank"))
        self.IncompletedMission.setObjectName("IncompletedMission")
        self.verticalLayout.addWidget(self.IncompletedMission)
        self.gridLayout.addLayout(self.verticalLayout, 0, 0, 1, 1)
        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.widget = QtWidgets.QWidget(self)
        upper_profile_width = width - 90
        self.widget.setGeometry(QtCore.QRect(int((width - upper_profile_width) / 2), 30, upper_profile_width, 80))
        self.widget.setObjectName("widget")
        self.UpperProfile = QtWidgets.QHBoxLayout(self.widget)
        self.UpperProfile.setContentsMargins(0, 0, 0, 0)
        self.UpperProfile.setObjectName("UpperProfile")
        self.Logo = QtWidgets.QLabel(self.widget)
        self.Logo.setMinimumSize(QtCore.QSize(175, 40))
        self.Logo.setMaximumSize(QtCore.QSize(175, 40))
        self.Logo.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.Logo.setText("")
        self.Logo.setPixmap(QtGui.QPixmap(image_directory + "Logo2.png"))
        self.Logo.setScaledContents(True)
        self.Logo.setObjectName("Logo")
        self.UpperProfile.addWidget(self.Logo)
        spacerItem1 = QtWidgets.QSpacerItem(28, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.UpperProfile.addItem(spacerItem1)
        self.NameMajor = QtWidgets.QLabel(self.widget)
        self.NameMajor.setMinimumSize(QtCore.QSize(650, 60))
        self.NameMajor.setMaximumSize(QtCore.QSize(650, 60))
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(14)
        font.setBold(False)
        font.setWeight(50)
        self.NameMajor.setFont(font)
        self.NameMajor.setStyleSheet("background-color: rgba(255, 255, 255, 0);")
        self.NameMajor.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignTrailing | QtCore.Qt.AlignVCenter)
        self.NameMajor.setObjectName("NameMajor")
        self.UpperProfile.addWidget(self.NameMajor)
        self.retranslateUi()

    def retranslateUi(self):
        from pandas import read_excel
        _translate = QtCore.QCoreApplication.translate
        info_file = program_directory + '\\edited\\info\\personalinfo2.xlsx'
        df_info = read_excel(info_file, header=None)
        my_semester = str(df_info.values[1][1])
        my_name = str(df_info.values[2][1])
        my_major = str(df_info.values[3][1])
        self.setWindowTitle(_translate("Form", "학점분석표"))
        self.Title.setText(_translate("Form", "학점분석표"))
        self.CompletedSemester.setText(_translate("Form", "현재 이수 진행(완료) 학기: " + my_semester + "학기"))
        self.NameMajor.setText(_translate("Form", "성명:   " + my_name + "\n전공:   " + my_major))
        info_directory = program_directory + 'edited\\info\\'
        local_url = QUrl.fromLocalFile(info_directory + "1. Credit_Mission_Graph.html")
        self.CreditAndMission.load(local_url)
        local_url = QUrl.fromLocalFile(info_directory + "2. Major_and_Entire_Grade.html")
        self.CreditGraph.load(local_url)
        local_url = QUrl.fromLocalFile(info_directory + "3. Completed_Mission.html")
        self.CompletedMission.load(local_url)
        local_url = QUrl.fromLocalFile(info_directory + "4. Uncompleted_Mission.html")
        self.IncompletedMission.load(local_url)

    def closeEvent(Form, event):
        close = QtWidgets.QMessageBox.question(Form,
                                               "QUIT",
                                               "분석표에 나온 표와 그래프를 저장하시겠습니까?\n"
                                               "(해당 표와 그래프는 edited/info 폴더에 html 파일 형태로 저장됩니다.)",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Cancel)
        global save_token
        if close == QtWidgets.QMessageBox.Yes:
            save_token = 1
            event.accept()
        elif close == QtWidgets.QMessageBox.No:
            save_token = 0
            event.accept()
        else:
            event.ignore()


def executeResult():
    if __name__ != "__main__":
        import sys
        import subprocess
        for char in getcwd():  # 경로에 한글이 있을 시 그래프가 있는 폴더만 보여주고 종료
            if ord('ㄱ') <= ord(char) <= ord('힣'):
                subprocess.Popen('explorer "' + program_directory + 'edited\\info"')
                return 1
        app2 = QtWidgets.QApplication(sys.argv)
        ui = Ui_Form()
        ui.setupUi()
        ui.setWindowIcon(QtGui.QIcon(image_directory + 'Icon.png'))
        ui.show()
        app2.exec()
        return save_token
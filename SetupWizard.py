from os import getcwd
from PyQt5 import QtCore, QtWidgets, QtGui

program_directory = getcwd()[:getcwd().index('Design')]
image_directory = program_directory + '\\Design Image\\'

cancel_token = 0  # 중간에 취소하는 등의 행동을 취했을 경우 값이 바뀜.(값이 바뀌면 이후 프로그램의 실행을 중단)


class MagicWizard(QtWidgets.QWizard):  # 프로그램 본체
    def __init__(self, parent=None):
        super(MagicWizard, self).__init__(parent)
        self.setWindowTitle("GIST Credit Analysis Program")
        self.setWindowIcon(QtGui.QIcon(image_directory + 'Icon.png'))
        self.setWizardStyle(QtWidgets.QWizard.ModernStyle)
        self.setWindowFlags(QtCore.Qt.WindowCloseButtonHint)
        self.addPage(Page1(self))
        self.addPage(Page3(self))  # Page2(파일 확인 페이지)는 폐기 처리(Page3에 파일 확인 기능 추가)
        self.addPage(Page4(self))
        self.setFixedSize(1300, 900)
        self.setPixmap(QtWidgets.QWizard.WatermarkPixmap, QtGui.QPixmap(image_directory + '설치 이미지4.png'))
        self.setButtonText(QtWidgets.QWizard.BackButton, '< 뒤로')
        self.setButtonText(QtWidgets.QWizard.NextButton, '다음 >')
        self.setButtonText(QtWidgets.QWizard.CommitButton, '다음 >')
        self.setButtonText(QtWidgets.QWizard.FinishButton, '완료')
        self.setButtonText(QtWidgets.QWizard.CancelButton, '취소')

    def PushCancelButton(self):
        global cancel_token
        cancel_token = 1

    def closeEvent(self, event):
        close = QtWidgets.QMessageBox.question(self, "QUIT", "GIST 학점 분석 프로그램을 종료하시겠습니까?",
                                               QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            event.accept()
            quit(2)
        else:
            event.ignore()


class Page1(QtWidgets.QWizardPage):  # 프로그램 소개
    def __init__(self, parent=None):
        super(Page1, self).__init__(parent)
        self.TitleLabel = QtWidgets.QLabel()
        self.ExplainLabel1 = QtWidgets.QLabel()
        self.ExplainLabel2 = QtWidgets.QLabel()
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.TitleLabel)

        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(27)
        font.setBold(True)
        self.TitleLabel.setFont(font)

        layout.addWidget(self.ExplainLabel1)
        layout.addWidget(self.ExplainLabel2)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(50)
        font.setKerning(True)
        self.ExplainLabel1.setFont(font)
        self.ExplainLabel2.setFont(font)
        self.ExplainLabel1.setWordWrap(True)
        self.ExplainLabel2.setWordWrap(True)

        self.setLayout(layout)
        self.setCommitPage(True)

    def initializePage(self):
        self.TitleLabel.setText("\nGIST 학점 분석 프로그램\n실행을 시작합니다.")
        self.ExplainLabel1.setText("\n\n해당 프로그램은 총 이수학점과 평점, 지금까지 달성했거나 앞으로 달성해야 할 졸업 요건 등을 분석하고자"
                                   "만들어진 프로그램입니다.")
        self.ExplainLabel2.setText("\n대학에서 공식으로 만들어진 프로그램이 아니므로 학점 계산 과정에서 오차가 발생할 수 있습니다. "
                                   "따라서 해당 분석표에 나온 자료는 단순 참고용으로만 사용하시기 바랍니다.")


class Page3(QtWidgets.QWizardPage):
    import ForPage3
    necfilestring, step_count, nextbuttonon = ForPage3.CheckNecFile(None)
    if nextbuttonon:
        download_info = 1
    else:
        download_info = 0
    downloaded_file_status = 3

    def __init__(self, parent=None):
        super(Page3, self).__init__(parent)
        # 텍스트 레이블 관련 명령
        # 제목 레이블 작성 관련 명령
        self.TitleLabel = QtWidgets.QLabel()
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.TitleLabel)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(27)
        font.setBold(True)
        self.TitleLabel.setFont(font)
        self.setLayout(layout)

        # 설명 내용 관련 레이블 작성
        self.ExplainLabel1 = QtWidgets.QLabel()
        layout.addWidget(self.ExplainLabel1)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(50)
        font.setKerning(True)
        self.ExplainLabel1.setFont(font)
        self.ExplainLabel1.setWordWrap(True)
        self.setLayout(layout)

        # 필요한 파일 내용 관련 레이블 작성
        self.NeedFileLabel = QtWidgets.QLabel()
        layout.addWidget(self.NeedFileLabel)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(15)
        font.setBold(True)
        font.setWeight(100)
        font.setKerning(True)
        self.NeedFileLabel.setFont(font)
        self.setLayout(layout)

        # 경고 레이블 작성
        self.WarningLabel = QtWidgets.QLabel()
        layout.addWidget(self.WarningLabel)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(50)
        font.setKerning(True)
        self.WarningLabel.setFont(font)
        self.WarningLabel.setWordWrap(True)
        self.setLayout(layout)

        # 단계별 설치 버튼 관련 레이블 작성
        self.ButtonStep1 = QtWidgets.QPushButton()
        self.ButtonStep2 = QtWidgets.QPushButton()
        self.ButtonStep3 = QtWidgets.QPushButton()
        button_layout = QtWidgets.QHBoxLayout()
        button_layout.addWidget(self.ButtonStep1)
        button_layout.addWidget(self.ButtonStep2)
        button_layout.addWidget(self.ButtonStep3)
        layout.addLayout(button_layout)
        self.ButtonStep1.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.ButtonStep2.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.ButtonStep3.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        self.ButtonStep1.setMinimumSize(QtCore.QSize(150, 60))
        self.ButtonStep2.setMinimumSize(QtCore.QSize(150, 60))
        self.ButtonStep3.setMinimumSize(QtCore.QSize(150, 60))
        self.ButtonStep1.setText('1. ZEUS 로그인')
        self.ButtonStep2.setText('2. 성적 및 학생정보 수집')
        self.ButtonStep3.setText('3. 교양 과목 DB 업데이트')
        font = QtGui.QFont()
        font.setFamily("나눔스퀘어 Bold")
        font.setPointSize(12)
        self.ButtonStep1.setFont(font)
        self.ButtonStep2.setFont(font)
        self.ButtonStep3.setFont(font)

        self.ButtonStep1.clicked.connect(self.Step1)
        self.ButtonStep2.clicked.connect(self.Step2)
        self.ButtonStep3.clicked.connect(self.Step3)

        self.ButtonStep2.setEnabled(False)
        self.ButtonStep3.setEnabled(False)
        self.setLayout(layout)

        # 상태창 관련 명령
        self.TextBrowser = QtWidgets.QTextBrowser()
        layout.addWidget(self.TextBrowser)
        self.TextBrowser.viewport().setProperty("cursor", QtGui.QCursor(QtCore.Qt.IBeamCursor))
        self.TextBrowser.setMouseTracking(False)
        self.TextBrowser.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.TextBrowser.setStyleSheet("background-color: rgb(20, 20, 20);")
        palette = QtGui.QPalette()
        palette.setColor(QtGui.QPalette.Text, QtCore.Qt.green)
        self.TextBrowser.setPalette(palette)
        font = QtGui.QFont()
        font.setFamily("Consolas")
        font.setPointSize(10)
        self.TextBrowser.setFont(font)
        self.TextBrowser.setWordWrapMode(QtGui.QTextOption.WordWrap)
        self.TextBrowser.setReadOnly(True)
        self.TextBrowser.setObjectName("Status")
        self.setLayout(layout)

        # 다음 버튼을 확인 버튼으로 바꿈(다음 페이지에서 뒤로 버튼 사용 불가)
        self.setCommitPage(True)

        # 상태 창에 현재 DB 날짜 출력
        from openpyxl import load_workbook
        self.cprint("=" * 35)
        try:
            load_wb = load_workbook(program_directory + 'edited/elective_list/DBinfo.xlsx', data_only=True)
            load_ws = load_wb['Sheet']
            DB_updated_date = load_ws['B1'].value
            DB_applied_year = load_ws['B2'].value
            self.cprint("최신 DB 업데이트 날짜: " + DB_updated_date)
            self.cprint("교양 과목 DB 적용 연도: 2014 ~ " + str(DB_applied_year)+"년")
        except:
            self.cprint("현재 폴더 내 데이터베이스 없음.")
        self.cprint("=" * 35)
        self.cprint('Chromedriver 프로그램을 설치하지 않은 경우 '
                    'Chrome 버전 확인 후 버전에 맞는 Chromedriver.exe 파일을 내려받은 다음, '
                    '프로그램 폴더(~/Design File)에 저장하십시오.\n'
                    '(해당 프로그램을 처음 이용할 경우 Chromedriver 파일을 내려받아야 합니다.)\n\n'
                    '1. Chrome 버전 확인: 아래 주소(Chrome 업데이트 항목)에서 \'Chrome 업데이트에 관한 기타 정보\' 부분의 '
                    '\'업데이트 및 현재 브라우저 버전 확인\' 섹션 참조.\n'
                    'https://support.google.com/chrome/answer/95414?co=GENIE.Platform%3DDesktop&hl=ko\n\n'
                    '2. Chromedriver 설치: https://chromedriver.chromium.org/downloads\n')



    def Step1(self):
        import ForPage3
        self.ButtonStep1.setEnabled(False)
        temp_parameter = -2
        try:
            temp_parameter, Page3.key_id, Page3.key_pw = ForPage3.PushStep1(self)  # Step1 버튼 눌렀을 때 환산되어 얻는 값
            print(temp_parameter)
        except:
            self.cprint('예상하지 못한 오류가 발생했습니다. 다시 시도해보시고, 계속해서 문제 발생 시 프로그램 관리지에게 연락하십시오.\n'
                        '혹은 Chromedriver가 설치되지 않았거나 현재 Chrome 버전과 호환되지 않을 수 있습니다.\n'
                        'Chrome 버전 확인 후 버전에 맞는 Chromedriver를 내려받은 다음 '
                        '프로그램 폴더(~/Design File)에 저장하십시오.\n'
                        'Chromedriver 설치: https://chromedriver.chromium.org/downloads')
            self.ButtonStep1.setEnabled(True)

        if temp_parameter != 100:  # 로그인 성공 아닐 시
            if temp_parameter == -1:
                self.cprint(
                    'Chromedriver가 설치되지 않았거나 현재 Chrome 버전과 호환되지 않습니다.\n'
                    'Chrome 버전 확인 후 버전에 맞는 Chromedriver를 내려받은 다음 '
                    '프로그램 폴더(~/Design File)에 저장하십시오.'
                    '\nChromedriver 설치: https://chromedriver.chromium.org/downloads')
            elif temp_parameter == 50:
                self.cprint('ZEUS 로그인에 실패했습니다. 다시 한 번 시도하십시오.')
            else:
                self.cprint('알 수 없는 오류입니다. 다시 한 번 시도하시고, 이후 실패 시 관리자에게 연락하십시오.')
            self.ButtonStep1.setEnabled(True)
        else:
            self.cprint('ZEUS 로그인에 성공했습니다. 다음 단계를 진행하십시오.')
            self.ButtonStep2.setEnabled(True)

    key_id = 0
    key_pw = 0
    def Step2(self):
        import ForPage3
        self.ButtonStep2.setEnabled(False)
        try:
            status_parameter = ForPage3.PushStep2(self, parameter_key_id=Page3.key_id, parameter_key_pw=Page3.key_pw)
            if status_parameter != 0:
                if status_parameter == 1:
                    self.cprint("학생의 이름과 전공 복사에 실패했습니다. 계속해서 문제 발생 시"
                                " 프로그램 관리지에게 연락하십시오.")
                elif status_parameter == 2:
                    self.cprint("Download 폴더에서 원본 성적표를 찾을 수 없습니다. 계속해서 문제 발생 시"
                                " 프로그램 관리지에게 연락하십시오.")
                self.ButtonStep2.setEnabled(True)
            else:
                if self.downloaded_file_status == 2:  # 3단계를 진행할 필요 없을 시
                    self.download_info = 1
                    self.completeChanged.emit()
                self.ButtonStep3.setEnabled(True)
        except:
            self.cprint('예상하지 못한 오류가 발생했습니다. 다시 시도해보시고, 계속해서 문제 발생 시 프로그램 관리지에게 연락하십시오.')
            self.ButtonStep2.setEnabled(True)

    def Step3(self):
        import ForPage3
        self.ButtonStep3.setEnabled(False)
        try:
            download_status = ForPage3.PushStep3(self)
            if download_status == 0:
                self.download_info = 1
                self.completeChanged.emit()
            else:
                self.cprint('오류가 발생했습니다. 다시 시도해보시고, '
                            '계속해서 문제 발생 시 프로그램 관리자에게 연락하십시오.')
        except:
            self.cprint('예상하지 못한 오류가 발생했습니다. 다시 시도해보시고, 계속해서 문제 발생 시 프로그램 관리지에게 연락하십시오.')
            self.ButtonStep3.setEnabled(True)

    def cprint(self, str_text):
        self.TextBrowser.append(str_text)
        self.TextBrowser.repaint()

    def isComplete(self):
        return self.download_info == 1

    def initializePage(self):
        self.TitleLabel.setText("\n파일 다운로드를 시작합니다.")
        self.ExplainLabel1.setText("\n\n아래 과정을 통해 필요한 파일을 내려받습니다.\n현재 필요한 파일은 다음과 같습니다.")
        import ForPage3
        necfilestring, step_count, nextbuttonon = ForPage3.CheckNecFile(self)
        self.NeedFileLabel.setText(necfilestring)
        self.downloaded_file_status = step_count
        self.WarningLabel.setText("※ 경고:  모든 단계에서 브라우저가 실행하는 동안에는 키보드와 마우스를 조작하지 마십시오."
                                  "\n            프로그램이 정상적으로 작동되지 않습니다.")


class Page4(QtWidgets.QWizardPage):  # 프로그램 소개
    def __init__(self, parent=None):
        super(Page4, self).__init__(parent)
        self.TitleLabel = QtWidgets.QLabel()
        self.ExplainLabel1 = QtWidgets.QLabel()
        layout = QtWidgets.QVBoxLayout()
        layout.addWidget(self.TitleLabel)

        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(27)
        font.setBold(True)
        self.TitleLabel.setFont(font)

        layout.addWidget(self.ExplainLabel1)
        font = QtGui.QFont()
        font.setFamily('맑은 고딕')
        font.setPointSize(13)
        font.setBold(True)
        font.setWeight(50)
        font.setKerning(True)
        self.ExplainLabel1.setFont(font)
        self.ExplainLabel1.setWordWrap(True)

        self.setLayout(layout)
        # 다음 버튼을 확인 버튼으로 바꿈(다음 페이지에서 뒤로 버튼 사용 불가)
        self.setCommitPage(True)

    def initializePage(self):
        self.TitleLabel.setText("\n파일 다운로드를 완료했습니다.")
        self.ExplainLabel1.setText("\n\n학점 분석에 필요한 파일을 모두 내려받았습니다. "
                                   "\n아래 \"완료\"버튼을 눌러 학점 분석표를 보시길 바랍니다.")


def executeWizard():
    if __name__ != "__main__":
        import sys
        app = QtWidgets.QApplication(sys.argv)
        wizard = MagicWizard()
        wizard.button(QtWidgets.QWizard.CancelButton).clicked.connect(MagicWizard.PushCancelButton)
        wizard.show()
        app.exec()
        return cancel_token



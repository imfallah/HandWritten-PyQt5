import sys
import os
import re
import subprocess
import platform

from document_parser import DocumentParser

from PyQt5 import QtCore, QtGui, QtWidgets
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from fbs_runtime.application_context.PyQt5 import ApplicationContext
import joblib

def read_stylesheet(path_to_sheet):
    with open(path_to_sheet, 'r') as f:
        stylesheet = f.read()
    return stylesheet

class Ui_MainWindow(QtCore.QObject):

    def setupUi(self, MainWindow, AppContext):
        ### Styles for application UI elements

        ## Stylesheets
        self.stylesheet_select_unselected = read_stylesheet(AppContext.get_resource('btn_select_unselected.qss'))
        self.stylesheet_select_selected = read_stylesheet(AppContext.get_resource('btn_select_selected.qss'))
        self.stylesheet_write_inactive = read_stylesheet(AppContext.get_resource('btn_write_inactive.qss'))
        self.stylesheet_write_active = read_stylesheet(AppContext.get_resource('btn_write_active.qss'))
        self.stylesheet_progressbar_busy = read_stylesheet(AppContext.get_resource('progressbar_busy.qss'))
        self.stylesheet_progressbar_finshed = read_stylesheet(AppContext.get_resource('progressbar_finished.qss'))

        ## Fonts
        self.font_asleep = QtGui.QFont('Roboto', 12)
        self.font_awake = QtGui.QFont('Roboto', 12)
        self.font_awake.setBold(True)
        font_select = QtGui.QFont('Roboto', 11)
        font_select.setBold(True)

        ### UI Elements
        path_animated_logo = AppContext.get_resource('DSC_logo_animated.gif')
        path_logo_small = AppContext.get_resource('handwriter_logo_small.png')
        path_logo = AppContext.get_resource('handwriter_logo.png')
        self.path_hashes = AppContext.get_resource('hashes.pickle')

        self.MainWindow = MainWindow
        self.MainWindow.setObjectName("HandWriter")
        self.MainWindow.setStyleSheet("QMainWindow {background: 'white'}")
        self.MainWindow.setFixedSize(800, 600)
        self.MainWindow.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)  # Disable window maximize button
        self.centralwidget = QtWidgets.QWidget(self.MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.MainWindow.setCentralWidget(self.centralwidget)
        self.MainWindow.setWindowIcon(QtGui.QIcon(path_logo_small))

        # Animated DSC Logo
        self.logo_label = QtWidgets.QLabel(self.centralwidget)
        self.logo_label.resize(200, 80)
        self.logo_label.move(320, 45)
        self.logo_movie = MovieBox(path_animated_logo).resized_movie(180)
        self.logo_movie.setSpeed(350)
        self.logo_movie.frameChanged.connect(self.check_stopping_frame)
        self.logo_label.setMovie(self.logo_movie)
        self.logo_movie.start()

        # Application name logo
        app_logo = QtWidgets.QLabel(self.centralwidget)
        app_logo.setPixmap(QtGui.QPixmap(path_logo).scaled(540, 100, QtCore.Qt.KeepAspectRatio, transformMode=QtCore.Qt.SmoothTransformation))
        app_logo.setFixedSize(540, 100)
        app_logo.setObjectName("app_logo")
        app_logo.setGeometry(185, 210, 300, 150)

        # Select Document Button
        self.btn_select_document = QtWidgets.QPushButton('Select Document', self.centralwidget)
        self.btn_select_document.setStyleSheet(self.stylesheet_select_unselected)
        self.btn_select_document.setEnabled(True)
        self.btn_select_document.setFixedSize(200, 50)
        self.btn_select_document.setFont(font_select)
        self.btn_select_document.setShortcut('Ctrl+O')
        self.btn_select_document.setGeometry(175, 445, 150, 50)

        # Write Button
        self.btn_write = QtWidgets.QPushButton('Write', self.centralwidget)
        self.btn_write.setEnabled(False)
        self.btn_write.setFixedSize(200, 50)
        self.btn_write.setFont(self.font_asleep)
        self.btn_write.setStyleSheet(self.stylesheet_write_inactive)
        self.btn_write.setShortcut('Ctrl+E')
        self.btn_write.setGeometry(435, 445, 150, 50)

        # Progress Bar
        self.progress = QtWidgets.QProgressBar(self.MainWindow)
        self.progress.setStyleSheet(self.stylesheet_progressbar_busy)
        self.progress.setGeometry(0, 590, 800, 10)

        # Add Menu button (Animated)
        self.btn_menu = QtWidgets.QPushButton('Converted Item', self.centralwidget)
        self.btn_menu.setFixedSize(200, 50)
        self.btn_menu.setGeometry(300, 500, 200, 50)
        self.btn_menu.setStyleSheet("background-color: #3E3E3E; color: white; font-size: 16px;")
        self.btn_menu.clicked.connect(self.show_converted_items)

        # Animation for the menu button (move effect)
        self.animation = QPropertyAnimation(self.btn_menu, b"geometry")
        self.animation.setDuration(1000)
        self.animation.setStartValue(QtCore.QRect(300, 500, 200, 50))
        self.animation.setEndValue(QtCore.QRect(300, 450, 200, 50))

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self.MainWindow)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.MainWindow.setWindowTitle(_translate("MainWindow", "HandWriter"))
        self.btn_select_document.clicked.connect(self.open_document)
        self.btn_write.clicked.connect(self.parse_document)
        self.MainWindow.show()

    def open_document(self):
        self.doc_path = QtWidgets.QFileDialog.getOpenFileName(self.MainWindow, 'Open Document', filter='*.docx')
        self.doc_path = self.doc_path[0]
        if self.doc_path == '':
            self.sleep_btn_write()
            self.unselect_btn_select()
            return
        try:
            self.document = Document(self.doc_path)
            self.selected_btn_select()
        except PackageNotFoundError:
            self.sleep_btn_write()
            return
        self.pdf_path = re.sub('docx', 'pdf', self.doc_path)
        self.wake_btn_write()

    def show_converted_items(self):
        # Show the animation when button is clicked
        self.animation.start()

        # Create a new window to display converted items
        self.converted_window = QtWidgets.QWidget()
        self.converted_window.setWindowTitle("Converted Items")
        self.converted_window.setFixedSize(400, 300)
        self.converted_window.setStyleSheet("background-color: white;")

        # Create a list widget to display the converted items
        self.list_widget = QtWidgets.QListWidget(self.converted_window)
        self.list_widget.setGeometry(50, 50, 300, 200)
        self.list_widget.setStyleSheet("border: 2px solid #2E2E2E; border-radius: 10px;")

        # Populate the list with converted items
        converted_items = ["Item 1", "Item 2", "Item 3"]  # List of converted items
        for item in converted_items:
            self.list_widget.addItem(item)

        self.converted_window.show()

    def parse_document(self):
        self.progress.setRange(0, 0)
        self.progress.setStyleSheet(self.stylesheet_progressbar_busy)
        self.start_parsing()

    def start_parsing(self):
        self.thread = ParserThread(self.doc_path, self.document, self.path_hashes)
        self.thread.change_value.connect(self.popup_success)
        self.thread.key_exception.connect(self.popup_keyerror)
        self.thread.start()

    def stop_progressbar(self):
        self.thread.requestInterruption()
        self.sleep_btn_write()
        self.progress.setRange(0, 1)
        self.progress.setStyleSheet(self.stylesheet_progressbar_finshed)
        self.progress.setTextVisible(False)
        self.progress.setValue(1)
        self.unselect_btn_select()

    def popup_success(self):
        self.stop_progressbar()
        success_popup = QtWidgets.QMessageBox(self.centralwidget)
        success_popup.setIcon(QtWidgets.QMessageBox.NoIcon)
        success_popup.setWindowTitle('Success: File Written')
        success_popup.setText('The file was successfully written to ' + self.pdf_path)
        btn_open_folder = QtWidgets.QPushButton('Open Containing Folder')
        btn_open_folder.clicked.connect(self.open_containing_folder)
        success_popup.addButton(btn_open_folder, QtWidgets.QMessageBox.AcceptRole)
        success_popup.setStandardButtons(QtWidgets.QMessageBox.Ok)
        success_popup.show()

    def open_containing_folder(self):
        if platform.system() == 'Windows':
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)
            if match:
                pdf_path = match.groups()[0]
                os.startfile(pdf_path)
        elif platform.system() == 'Darwin':
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)
            if match:
                pdf_path = match.groups()[0]
                subprocess.Popen(['open', pdf_path])
        else:
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)
            if match:
                pdf_path = match.groups()[0]
                subprocess.Popen(['xdg-open', pdf_path])

    def popup_keyerror(self):
        self.stop_progressbar()
        keyerror_popup = QtWidgets.QMessageBox(self.centralwidget)
        keyerror_popup.setIcon(QtWidgets.QMessageBox.Critical)
        keyerror_popup.setWindowTitle('KeyError Exception')
        keyerror_popup.setText('The following exception was thrown during parsing:\n\n\n\n[KeyError]')
        keyerror_popup.setStandardButtons(QtWidgets.QMessageBox.Ok)
        keyerror_popup.show()

    def wake_btn_write(self):
        self.btn_write.setEnabled(True)
        self.btn_write.setStyleSheet(self.stylesheet_write_active)
        self.btn_write.setFont(self.font_awake)

    def sleep_btn_write(self):
        self.btn_write.setEnabled(False)
        self.btn_write.setStyleSheet(self.stylesheet_write_inactive)
        self.btn_write.setFont(self.font_asleep)

    def selected_btn_select(self):
        self.btn_select_document.setStyleSheet(self.stylesheet_select_selected)

    def unselect_btn_select(self):
        self.btn_select_document.setStyleSheet(self.stylesheet_select_unselected)

class ParserThread(QtCore.QThread):
    change_value = QtCore.pyqtSignal()
    key_exception = QtCore.pyqtSignal()

    def __init__(self, doc_path, document, path_hashes):
        super(ParserThread, self).__init__()
        self.doc_path = doc_path
        self.document = document
        self.hashes = joblib.load(path_hashes)

    def run(self):
        for para in self.document.paragraphs:
            for run in para.runs:
                if self.isInterruptionRequested():
                    return
                try:
                    paragraph_hash = self.hashes[run.text]
                    run.text = paragraph_hash
                except KeyError:
                    self.key_exception.emit()
                    return
        self.document.save(self.doc_path.replace('docx', 'pdf'))
        self.change_value.emit()

class MovieBox(QtGui.QMovie):
    def resized_movie(self, width):
        self.jumpToFrame(0)  # شروع از فریم اول
        movie_size = self.currentImage().size()  # اندازه فعلی تصویر
        movie_aspect = movie_size.width() / movie_size.height()  # نسبت ابعاد

        # تغییر اندازه QMovie
        self.setScaledSize(QtCore.QSize(width, int(width / movie_aspect)))
        return self  # برگرداندن خود شیء QMovie

# سایر کلاس‌ها و توابع کد شما

if __name__ == '__main__':
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow, appctxt)
    MainWindow.show()
    exit_code = appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()
    sys.exit(exit_code)

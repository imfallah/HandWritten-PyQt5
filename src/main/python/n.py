import sys
import os
import re
import subprocess
import platform
from document_parser import DocumentParser
from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from PyQt5 import QtCore, QtGui, QtWidgets
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
        # 'Select Document' button          
        self.stylesheet_select_unselected = read_stylesheet(AppContext.get_resource('btn_select_unselected.qss'))
        self.stylesheet_select_selected = read_stylesheet(AppContext.get_resource('btn_select_selected.qss'))
        #========================================================================================
        # 'Write' button
        self.stylesheet_write_inactive = read_stylesheet(AppContext.get_resource('btn_write_inactive.qss'))
        self.stylesheet_write_active = read_stylesheet(AppContext.get_resource('btn_write_active.qss'))
        #========================================================================================
        # Progressbar
        self.stylesheet_progressbar_busy = read_stylesheet(AppContext.get_resource('progressbar_busy.qss'))
        self.stylesheet_progressbar_finshed = read_stylesheet(AppContext.get_resource('progressbar_finished.qss'))
        #========================================================================================
        ## Fonts   
        # Write inactive
        self.font_asleep = QtGui.QFont('Roboto', 12)
        # Write active
        self.font_awake = QtGui.QFont('Roboto', 12)
        self.font_awake.setBold(True)
        # Select
        font_select = QtGui.QFont('Roboto', 11)
        font_select.setBold(True)
        #========================================================================================
        ### UI Elements
        path_logo = AppContext.get_resource('handlogo.png')
        path_handwritten_image = AppContext.get_resource('E:/project/HandWriter/src/main/resources/base/handlogo.png')  # مسیر تصویر جدید

        self.path_hashes = AppContext.get_resource('hashes.pickle')
        self.MainWindow = MainWindow
        self.MainWindow.setObjectName("HandWriter")
        self.MainWindow.setStyleSheet("QMainWindow {background: 'white'}")
        self.MainWindow.setFixedSize(800, 600)
        self.MainWindow.setWindowFlags(QtCore.Qt.WindowCloseButtonHint | QtCore.Qt.WindowMinimizeButtonHint)    # Disable window maximize button
        self.centralwidget = QtWidgets.QWidget(self.MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.MainWindow.setCentralWidget(self.centralwidget)


        self.handwritten_image_label = QtWidgets.QLabel(self.centralwidget)
        self.handwritten_image_label.setGeometry(QtCore.QRect(320, 180, 250, 40))   
        self.handwritten_image_label.setAlignment(QtCore.Qt.AlignCenter)
        self.handwritten_image_label.setScaledContents(True)

        pixmap = QtGui.QPixmap(path_handwritten_image)
        self.handwritten_image_label.setPixmap(pixmap.scaled(200, 80, QtCore.Qt.KeepAspectRatio, transformMode=QtCore.Qt.SmoothTransformation))

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

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self.MainWindow)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.MainWindow.setWindowTitle(_translate("MainWindow", "HandWriter"))
        self.btn_select_document.clicked.connect(self.open_document)
        self.btn_write.clicked.connect(self.parse_document)

        self.MainWindow.show()


    def open_document(self):
        self.doc_path = QtWidgets.QFileDialog.getOpenFileName(self.MainWindow, 'Open Document', filter = '*.docx')
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
        
    # Parse document on a thread separate from main UI thread
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

    def popup_keyerror(self, foreign_char):
        self.stop_progressbar()
        error_popup = QtWidgets.QMessageBox(self.centralwidget)
        error_popup.setIcon(QtWidgets.QMessageBox.Critical)
        error_popup.setWindowTitle('Error: Unable to write character')
        error_popup.setText('The character ' + foreign_char + ' has not been fed into this version of HandWriter. Raise an issue on the official GDGVIT repo')
        error_popup.setStandardButtons(QtWidgets.QMessageBox.Ok)
        error_popup.show()

    def open_containing_folder(self):
        print(f"PDF Path: {self.pdf_path}")
    
        if platform.system() == 'Windows':
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)  # تغییر الگو
            if match:
                pdf_path = match.groups()[0]
                os.startfile(pdf_path)
            else:
                print("No match found for Windows path.")
    
        elif platform.system() == 'Darwin':
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)
            if match:
                pdf_path = match.groups()[0]
                subprocess.Popen(['open', pdf_path])
            else:
                print("No match found for Darwin path.")
            
        else:
            match = re.search(r'^(.+)[/\\]([^/\\]+)$', self.pdf_path)
            if match:
                pdf_path = match.groups()[0]
                subprocess.Popen(['xdg-open', pdf_path])
            else:
                print("No match found for Linux path.")

    def wake_btn_write(self):
        self.btn_write.setFont(self.font_awake)
        self.btn_write.setStyleSheet(self.stylesheet_write_active)
        self.btn_write.setEnabled(True)

    def sleep_btn_write(self):
        self.btn_write.setFont(self.font_asleep)
        self.btn_write.setStyleSheet(self.stylesheet_write_inactive)
        self.btn_write.setEnabled(False)

    def selected_btn_select(self):
        self.btn_select_document.setStyleSheet(self.stylesheet_select_selected)

    def unselect_btn_select(self):
        self.btn_select_document.setStyleSheet(self.stylesheet_select_unselected)

class ParserThread(QtCore.QThread):
    def __init__(self, doc_path, document, path_hashes):
        super(ParserThread, self).__init__()
        self.doc_path = doc_path
        self.document = document
        self.HASHES = path_hashes  # اطمینان حاصل کنید که HASHES مقداردهی اولیه می‌شود

    change_value = QtCore.pyqtSignal()
    key_exception = QtCore.pyqtSignal(str)
    
    def run(self):
        CHARS_PER_LINE = 54
        LINES_PER_PAGE = 30
        # باز کردن فایل HASHES
        with open(self.HASHES, 'rb') as f:
            hashes = joblib.load(f)
        
        document_parser = DocumentParser(hashes, CHARS_PER_LINE, LINES_PER_PAGE)
        pdf_path = re.sub('docx', 'pdf', self.doc_path)
        
        try:
            document_parser.parse_document(self.document, pdf_path)
            self.change_value.emit()
        except KeyError as e:
            self.key_exception.emit(str(e)[1])
            
class MovieBox():
    def __init__(self, movie_path):
        self.movie = QtGui.QMovie(movie_path)

    def resized_movie(self, width):
        self.movie.jumpToFrame(0)
        movie_size = self.movie.currentImage().size()
        movie_aspect = movie_size.width() / movie_size.height()

        self.movie.setScaledSize(QtCore.QSize(width, int(width / movie_aspect)))
        return self.movie

if __name__ == '__main__':
    appctxt = ApplicationContext()       # 1. Instantiate ApplicationContext
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow, appctxt)      # اطمینان از اینکه MainWindow به درستی به setupUi منتقل شده است
    MainWindow.show()
    exit_code = appctxt.app.exec_()      # 2. Invoke appctxt.app.exec_()
    sys.exit(exit_code)

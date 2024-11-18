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
        self.MainWindow = MainWindow  

        # Load stylesheets
        self.stylesheet_select_unselected = read_stylesheet(AppContext.get_resource('btn_select_unselected.qss'))
        self.stylesheet_select_selected = read_stylesheet(AppContext.get_resource('btn_select_selected.qss'))
        self.stylesheet_write_inactive = read_stylesheet(AppContext.get_resource('btn_write_inactive.qss'))
        self.stylesheet_write_active = read_stylesheet(AppContext.get_resource('btn_write_active.qss'))
        self.stylesheet_progressbar_busy = read_stylesheet(AppContext.get_resource('progressbar_busy.qss'))
        self.stylesheet_progressbar_finished = read_stylesheet(AppContext.get_resource('progressbar_finished.qss'))

        # UI Elements
        path_logo_small = AppContext.get_resource('handwriter_logo_small.png')
        self.MainWindow.setObjectName("HandWriter")
        self.MainWindow.setStyleSheet("QMainWindow {background: 'white'}")
        self.MainWindow.setFixedSize(800, 600)

        # Make window frameless and round corners
        self.MainWindow.setWindowFlags(QtCore.Qt.Window | QtCore.Qt.FramelessWindowHint)
        self.MainWindow.setAttribute(QtCore.Qt.WA_TranslucentBackground)

        self.centralwidget = QtWidgets.QWidget(self.MainWindow)
        self.MainWindow.setCentralWidget(self.centralwidget)
        self.MainWindow.setWindowIcon(QtGui.QIcon(path_logo_small))

        # Title Bar
        self.title_bar = QtWidgets.QWidget(self.MainWindow)
        self.title_bar.setGeometry(0, 0, 800, 40)
        self.title_bar.setStyleSheet("background-color: #4CAF50; border: 2px; border-radius: 10px;")

        # Title Label
        self.title_label = QtWidgets.QLabel("HandWriter", self.title_bar)
        self.title_label.setGeometry(10, 10, 200, 20)
        self.title_label.setStyleSheet("color: white; font-size: 16px; font-weight: bold;")

        # Close Button
        self.btn_close = QtWidgets.QPushButton("X", self.title_bar)
        self.btn_close.setGeometry(760, 5, 30, 30)
        self.btn_close.setStyleSheet("background-color: red; color: white; border: none;")
        self.btn_close.clicked.connect(self.MainWindow.close)

        # Minimize Button
        self.btn_minimize = QtWidgets.QPushButton("_", self.title_bar)
        self.btn_minimize.setGeometry(720, 5, 30, 30)
        self.btn_minimize.setStyleSheet("background-color: yellow; color: black; border: none;")
        self.btn_minimize.clicked.connect(self.MainWindow.showMinimized)

        # Handwritten Label
        self.handwritten_label = QtWidgets.QLabel("handwritten", self.centralwidget)
        self.handwritten_label.setGeometry(QtCore.QRect(320, 45, 200, 80))   
        self.handwritten_label.setAlignment(QtCore.Qt.AlignCenter)
        font_handwritten = QtGui.QFont('Roboto', 16)
        font_handwritten.setBold(True)
        self.handwritten_label.setFont(font_handwritten)

        # Select Document Button
        self.btn_select_document = QtWidgets.QPushButton('Select Document', self.centralwidget)
        self.btn_select_document.setStyleSheet(self.stylesheet_select_unselected)
        self.btn_select_document.setEnabled(True)
        self.btn_select_document.setFixedSize(200, 50)
        self.btn_select_document.setFont(QtGui.QFont('Roboto', 11, QtGui.QFont.Bold))
        self.btn_select_document.setShortcut('Ctrl+O')
        self.btn_select_document.setGeometry(175, 445, 150, 50)

        # Write Button
        self.btn_write = QtWidgets.QPushButton('Write', self.centralwidget)
        self.btn_write.setEnabled(False)
        self.btn_write.setFixedSize(200, 50)
        self.btn_write.setFont(QtGui.QFont('Roboto', 12))
        self.btn_write.setStyleSheet(self.stylesheet_write_inactive)
        self.btn_write.setShortcut('Ctrl+E')
        self.btn_write.setGeometry(435, 445, 150, 50)

        # Progress Bar
        self.progress = QtWidgets.QProgressBar(self.MainWindow)
        self.progress.setStyleSheet(self.stylesheet_progressbar_busy)
        self.progress.setGeometry(0, 590, 800, 10)

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self.MainWindow)

        # Enable mouse tracking for the title bar
        self.title_bar.mousePressEvent = self.mousePressEvent
        self.title_bar.mouseMoveEvent = self.mouseMoveEvent

        # Initialize variables for dragging
        self.is_dragging = False
        self.drag_position = None

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.MainWindow.setWindowTitle(_translate("MainWindow", "HandWriter"))
        self.btn_select_document.clicked.connect(self.open_document)
        self.btn_write.clicked.connect(self.parse_document)

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.is_dragging = True
            self.drag_position = event.globalPos() - self.MainWindow.frameGeometry().topLeft()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.is_dragging:
            self.MainWindow.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.is_dragging = False

    def open_document(self):
        self.doc_path, _ = QtWidgets.QFileDialog.getOpenFileName(self.MainWindow, 'Open Document', filter='*.docx')
        if not self.doc_path:
            self.sleep_btn_write()
            self.unselect_btn_select()
            return
        try:
            self.document = Document(self.doc_path)
            self.selected_btn_select()
        except PackageNotFoundError:
            self.sleep_btn_write()
            return
        self.pdf_path = re.sub('docx$', 'pdf', self.doc_path)
        self.wake_btn_write()

    # ... (The rest of your methods remain unchanged)

if __name__ == "__main__":
    app_context = ApplicationContext()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow, app_context)
    MainWindow.show()
    sys.exit(app.exec_())

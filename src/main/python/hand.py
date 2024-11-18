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
        ### Styles for appliction UI elements
        ## Stylesheets
        # 'Select Document' button          
        self.stylesheet_select_unselected = read_stylesheet(AppContext.get_resource('btn_select_unselected.qss'))
        self.stylesheet_select_selected = read_stylesheet(AppContext.get_resource('btn_select_selected.qss'))
        # 'Write' button
        self.stylesheet_write_inactive = read_stylesheet(AppContext.get_resource('btn_write_inactive.qss'))
        self.stylesheet_write_active = read_stylesheet(AppContext.get_resource('btn_write_active.qss'))
        # Progressbar
        self.stylesheet_progressbar_busy = read_stylesheet(AppContext.get_resource('progressbar_busy.qss'))
        self.stylesheet_progressbar_finshed = read_stylesheet(AppContext.get_resource('progressbar_finished.qss'))

        ## Fonts
        # Write inactive
        self.font_asleep = QtGui.QFont('Roboto', 12)
        # Write active
        self.font_awake = QtGui.QFont('Roboto', 12)
        self.font_awake.setBold(True)
        # Select
        font_select = QtGui.QFont('Roboto', 11)
        font_select.setBold(True)

        ### UI Elements

        # Main Window

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

        # Close Button
        self.btn_close = QtWidgets.QPushButton('Exit', self.centralwidget)
        self.btn_close.setFixedSize(100, 40)
        self.btn_close.setFont(self.font_asleep)
        self.btn_close.setGeometry(650, 10, 100, 40)
        self.btn_close.clicked.connect(self.close_app)

        # Minimize Button
        self.btn_minimize = QtWidgets.QPushButton('Minimize', self.centralwidget)
        self.btn_minimize.setFixedSize(100, 40)
        self.btn_minimize.setFont(self.font_asleep)
        self.btn_minimize.setGeometry(530, 10, 100, 40)
        self.btn_minimize.clicked.connect(self.minimize_app)

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self.MainWindow)

    def retranslateUi(self):
        _translate = QtCore.QCoreApplication.translate
        self.MainWindow.setWindowTitle(_translate("MainWindow", "HandWriter"))
        self.btn_select_document.clicked.connect(self.open_document)
        self.btn_write.clicked.connect(self.parse_document)

        self.MainWindow.show()

    def close_app(self):
        QtCore.QCoreApplication.instance().quit()

    def minimize_app(self):
        self.MainWindow.showMinimized()

    # Rest of the methods as they were...

from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5 import QtWidgets

import digikey
from digikey.v3.productinformation import KeywordSearchRequest
from digikey.v3.productinformation.models import Filters
from digikey.v3.productinformation.models import ParametricFilter

import pcbnew
import os
import sys
import csv
import xlsxwriter
import json
import re
import datetime
from dateutil.tz import tzutc

# import logging

#logger = #logging.get#logger(__name__)
# with open("~/uBitFabKit.log", "w") as f:
#     pass
#logging.basicConfig(filename = "~/uBitFabKit.log", level = #logging.INFO)


############################################################################################################
#                                                                                                          #
#                                              GUIs                                                        #
#                                                                                                          #
############################################################################################################

class MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(453, 468)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 30, 431, 41))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label.setMinimumSize(QtCore.QSize(90, 0))
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.line_file = QtWidgets.QLineEdit(self.horizontalLayoutWidget)
        self.line_file.setReadOnly(True)
        self.line_file.setObjectName("line_file")
        self.horizontalLayout.addWidget(self.line_file)
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(10, 80, 431, 41))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.horizontalLayoutWidget_2)
        self.label_2.setMinimumSize(QtCore.QSize(90, 0))
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.line_board = QtWidgets.QLineEdit(self.horizontalLayoutWidget_2)
        self.line_board.setObjectName("line_board")
        self.horizontalLayout_2.addWidget(self.line_board)
        self.button_create_bom = QtWidgets.QPushButton(self.centralwidget)
        self.button_create_bom.setGeometry(QtCore.QRect(120, 390, 89, 25))
        self.button_create_bom.setObjectName("button_create_bom")
        self.button_load_file = QtWidgets.QPushButton(self.centralwidget)
        self.button_load_file.setGeometry(QtCore.QRect(160, 140, 111, 25))
        self.button_load_file.setObjectName("button_load_file")
        self.label_error = QtWidgets.QLabel(self.centralwidget)
        self.label_error.setGeometry(QtCore.QRect(10, 350, 431, 20))
        self.label_error.setText("")
        self.label_error.setAlignment(QtCore.Qt.AlignCenter)
        self.label_error.setObjectName("label_error")
        self.button_close = QtWidgets.QPushButton(self.centralwidget)
        self.button_close.setGeometry(QtCore.QRect(230, 390, 89, 25))
        self.button_close.setObjectName("button_close")
        self.horizontalLayoutWidget_3 = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget_3.setGeometry(QtCore.QRect(10, 230, 431, 41))
        self.horizontalLayoutWidget_3.setObjectName("horizontalLayoutWidget_3")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_3)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.horizontalLayoutWidget_3)
        self.label_3.setMinimumSize(QtCore.QSize(90, 0))
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.line_capacitor_manufacturer = QtWidgets.QLineEdit(self.horizontalLayoutWidget_3)
        self.line_capacitor_manufacturer.setObjectName("line_capacitor_manufacturer")
        self.horizontalLayout_3.addWidget(self.line_capacitor_manufacturer)
        self.horizontalLayoutWidget_4 = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget_4.setGeometry(QtCore.QRect(10, 280, 431, 41))
        self.horizontalLayoutWidget_4.setObjectName("horizontalLayoutWidget_4")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_4)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.horizontalLayoutWidget_4)
        self.label_4.setMinimumSize(QtCore.QSize(90, 0))
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4)
        self.line_resistor_manufacturer = QtWidgets.QLineEdit(self.horizontalLayoutWidget_4)
        self.line_resistor_manufacturer.setObjectName("line_resistor_manufacturer")
        self.horizontalLayout_4.addWidget(self.line_resistor_manufacturer)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(140, 200, 161, 17))
        self.label_5.setObjectName("label_5")
        self.radio_button_grouped = QtWidgets.QRadioButton(self.centralwidget)
        self.radio_button_grouped.setGeometry(QtCore.QRect(110, 330, 112, 23))
        self.radio_button_grouped.setChecked(True)
        self.radio_button_grouped.setObjectName("radio_button_grouped")
        self.radio_button_split = QtWidgets.QRadioButton(self.centralwidget)
        self.radio_button_split.setGeometry(QtCore.QRect(230, 330, 112, 23))
        self.radio_button_split.setObjectName("radio_button_split")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 453, 22))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.label.setText(_translate("MainWindow", "File"))
        self.label_2.setText(_translate("MainWindow", "Board name"))
        self.button_create_bom.setText(_translate("MainWindow", "Create &BOM"))
        self.button_load_file.setText(_translate("MainWindow", "&Select folder"))
        self.button_close.setText(_translate("MainWindow", "&Close"))
        self.label_3.setText(_translate("MainWindow", "Capacitor"))
        self.label_4.setText(_translate("MainWindow", "Resistor"))
        self.label_5.setText(_translate("MainWindow", "Default manufacturers"))
        self.radio_button_grouped.setText(_translate("MainWindow", "Grouped"))
        self.radio_button_split.setText(_translate("MainWindow", "Split"))



class partWindowDialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(783, 516)
        self.horizontalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(20, 20, 351, 51))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label.setMinimumSize(QtCore.QSize(40, 0))
        self.label.setObjectName("label")
        self.horizontalLayout.addWidget(self.label)
        self.line_row = QtWidgets.QLineEdit(self.horizontalLayoutWidget)
        self.line_row.setReadOnly(True)
        self.line_row.setObjectName("line_row")
        self.horizontalLayout.addWidget(self.line_row)
        self.horizontalLayoutWidget_2 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_2.setGeometry(QtCore.QRect(20, 70, 351, 51))
        self.horizontalLayoutWidget_2.setObjectName("horizontalLayoutWidget_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_2)
        self.horizontalLayout_2.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.horizontalLayoutWidget_2)
        self.label_2.setMinimumSize(QtCore.QSize(40, 0))
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.line_part = QtWidgets.QLineEdit(self.horizontalLayoutWidget_2)
        self.line_part.setReadOnly(True)
        self.line_part.setObjectName("line_part")
        self.horizontalLayout_2.addWidget(self.line_part)
        self.horizontalLayoutWidget_3 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_3.setGeometry(QtCore.QRect(20, 170, 351, 51))
        self.horizontalLayoutWidget_3.setObjectName("horizontalLayoutWidget_3")
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_3)
        self.horizontalLayout_3.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.label_3 = QtWidgets.QLabel(self.horizontalLayoutWidget_3)
        self.label_3.setMinimumSize(QtCore.QSize(115, 0))
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_3.addWidget(self.label_3)
        self.line_manufacturer = QtWidgets.QLineEdit(self.horizontalLayoutWidget_3)
        self.line_manufacturer.setReadOnly(True)
        self.line_manufacturer.setObjectName("line_manufacturer")
        self.horizontalLayout_3.addWidget(self.line_manufacturer)
        self.horizontalLayoutWidget_4 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_4.setGeometry(QtCore.QRect(20, 220, 351, 51))
        self.horizontalLayoutWidget_4.setObjectName("horizontalLayoutWidget_4")
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_4)
        self.horizontalLayout_4.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.label_4 = QtWidgets.QLabel(self.horizontalLayoutWidget_4)
        self.label_4.setMinimumSize(QtCore.QSize(115, 0))
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_4.addWidget(self.label_4)
        self.line_manufacturer_id = QtWidgets.QLineEdit(self.horizontalLayoutWidget_4)
        self.line_manufacturer_id.setReadOnly(True)
        self.line_manufacturer_id.setObjectName("line_manufacturer_id")
        self.horizontalLayout_4.addWidget(self.line_manufacturer_id)
        self.horizontalLayoutWidget_5 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_5.setGeometry(QtCore.QRect(20, 270, 351, 51))
        self.horizontalLayoutWidget_5.setObjectName("horizontalLayoutWidget_5")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_5)
        self.horizontalLayout_5.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_5 = QtWidgets.QLabel(self.horizontalLayoutWidget_5)
        self.label_5.setMinimumSize(QtCore.QSize(115, 0))
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_5.addWidget(self.label_5)
        self.line_distributor = QtWidgets.QLineEdit(self.horizontalLayoutWidget_5)
        self.line_distributor.setReadOnly(True)
        self.line_distributor.setObjectName("line_distributor")
        self.horizontalLayout_5.addWidget(self.line_distributor)
        self.horizontalLayoutWidget_6 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_6.setGeometry(QtCore.QRect(20, 320, 351, 51))
        self.horizontalLayoutWidget_6.setObjectName("horizontalLayoutWidget_6")
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_6)
        self.horizontalLayout_6.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_6 = QtWidgets.QLabel(self.horizontalLayoutWidget_6)
        self.label_6.setMinimumSize(QtCore.QSize(115, 0))
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_6.addWidget(self.label_6)
        self.line_distributor_id = QtWidgets.QLineEdit(self.horizontalLayoutWidget_6)
        self.line_distributor_id.setReadOnly(True)
        self.line_distributor_id.setObjectName("line_distributor_id")
        self.horizontalLayout_6.addWidget(self.line_distributor_id)
        self.horizontalLayoutWidget_7 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_7.setGeometry(QtCore.QRect(20, 370, 351, 51))
        self.horizontalLayoutWidget_7.setObjectName("horizontalLayoutWidget_7")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_7)
        self.horizontalLayout_7.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.label_7 = QtWidgets.QLabel(self.horizontalLayoutWidget_7)
        self.label_7.setMinimumSize(QtCore.QSize(115, 0))
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_7.addWidget(self.label_7)
        self.line_description = QtWidgets.QLineEdit(self.horizontalLayoutWidget_7)
        self.line_description.setReadOnly(True)
        self.line_description.setObjectName("line_description")
        self.horizontalLayout_7.addWidget(self.line_description)
        self.horizontalLayoutWidget_8 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_8.setGeometry(QtCore.QRect(410, 320, 351, 51))
        self.horizontalLayoutWidget_8.setObjectName("horizontalLayoutWidget_8")
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_8)
        self.horizontalLayout_8.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.label_8 = QtWidgets.QLabel(self.horizontalLayoutWidget_8)
        self.label_8.setMinimumSize(QtCore.QSize(125, 0))
        self.label_8.setObjectName("label_8")
        self.horizontalLayout_8.addWidget(self.label_8)
        self.line_total_price_1 = QtWidgets.QLineEdit(self.horizontalLayoutWidget_8)
        self.line_total_price_1.setReadOnly(True)
        self.line_total_price_1.setObjectName("line_total_price_1")
        self.horizontalLayout_8.addWidget(self.line_total_price_1)
        self.horizontalLayoutWidget_9 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_9.setGeometry(QtCore.QRect(410, 370, 351, 51))
        self.horizontalLayoutWidget_9.setObjectName("horizontalLayoutWidget_9")
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_9)
        self.horizontalLayout_9.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.label_9 = QtWidgets.QLabel(self.horizontalLayoutWidget_9)
        self.label_9.setMinimumSize(QtCore.QSize(125, 0))
        self.label_9.setObjectName("label_9")
        self.horizontalLayout_9.addWidget(self.label_9)
        self.line_total_price_1k = QtWidgets.QLineEdit(self.horizontalLayoutWidget_9)
        self.line_total_price_1k.setReadOnly(True)
        self.line_total_price_1k.setObjectName("line_total_price_1k")
        self.horizontalLayout_9.addWidget(self.line_total_price_1k)
        self.horizontalLayoutWidget_10 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_10.setGeometry(QtCore.QRect(410, 220, 351, 51))
        self.horizontalLayoutWidget_10.setObjectName("horizontalLayoutWidget_10")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_10)
        self.horizontalLayout_10.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label_10 = QtWidgets.QLabel(self.horizontalLayoutWidget_10)
        self.label_10.setMinimumSize(QtCore.QSize(125, 0))
        self.label_10.setObjectName("label_10")
        self.horizontalLayout_10.addWidget(self.label_10)
        self.line_price_1 = QtWidgets.QLineEdit(self.horizontalLayoutWidget_10)
        self.line_price_1.setReadOnly(True)
        self.line_price_1.setObjectName("line_price_1")
        self.horizontalLayout_10.addWidget(self.line_price_1)
        self.horizontalLayoutWidget_11 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_11.setGeometry(QtCore.QRect(410, 270, 351, 51))
        self.horizontalLayoutWidget_11.setObjectName("horizontalLayoutWidget_11")
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_11)
        self.horizontalLayout_11.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.label_11 = QtWidgets.QLabel(self.horizontalLayoutWidget_11)
        self.label_11.setMinimumSize(QtCore.QSize(125, 0))
        self.label_11.setObjectName("label_11")
        self.horizontalLayout_11.addWidget(self.label_11)
        self.line_price_1k = QtWidgets.QLineEdit(self.horizontalLayoutWidget_11)
        self.line_price_1k.setReadOnly(True)
        self.line_price_1k.setObjectName("line_price_1k")
        self.horizontalLayout_11.addWidget(self.line_price_1k)
        self.horizontalLayoutWidget_12 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_12.setGeometry(QtCore.QRect(410, 170, 351, 51))
        self.horizontalLayoutWidget_12.setObjectName("horizontalLayoutWidget_12")
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_12)
        self.horizontalLayout_12.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.label_12 = QtWidgets.QLabel(self.horizontalLayoutWidget_12)
        self.label_12.setMinimumSize(QtCore.QSize(125, 0))
        self.label_12.setObjectName("label_12")
        self.horizontalLayout_12.addWidget(self.label_12)
        self.line_type = QtWidgets.QLineEdit(self.horizontalLayoutWidget_12)
        self.line_type.setReadOnly(True)
        self.line_type.setObjectName("line_type")
        self.horizontalLayout_12.addWidget(self.line_type)
        self.button_choose = QtWidgets.QPushButton(Dialog)
        self.button_choose.setGeometry(QtCore.QRect(670, 470, 89, 25))
        self.button_choose.setObjectName("button_choose")
        self.button_skip = QtWidgets.QPushButton(Dialog)
        self.button_skip.setGeometry(QtCore.QRect(570, 470, 89, 25))
        self.button_skip.setObjectName("button_skip")
        self.horizontalLayoutWidget_13 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_13.setGeometry(QtCore.QRect(410, 20, 351, 51))
        self.horizontalLayoutWidget_13.setObjectName("horizontalLayoutWidget_13")
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_13)
        self.horizontalLayout_13.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.label_13 = QtWidgets.QLabel(self.horizontalLayoutWidget_13)
        self.label_13.setMaximumSize(QtCore.QSize(100, 16777215))
        self.label_13.setObjectName("label_13")
        self.horizontalLayout_13.addWidget(self.label_13)
        self.combobox_parts = QtWidgets.QComboBox(self.horizontalLayoutWidget_13)
        self.combobox_parts.setObjectName("combobox_parts")
        self.horizontalLayout_13.addWidget(self.combobox_parts)
        self.horizontalLayoutWidget_14 = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget_14.setGeometry(QtCore.QRect(410, 70, 351, 51))
        self.horizontalLayoutWidget_14.setObjectName("horizontalLayoutWidget_14")
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget_14)
        self.horizontalLayout_14.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.button_previous = QtWidgets.QPushButton(self.horizontalLayoutWidget_14)
        self.button_previous.setObjectName("button_previous")
        self.horizontalLayout_14.addWidget(self.button_previous)
        self.button_next = QtWidgets.QPushButton(self.horizontalLayoutWidget_14)
        self.button_next.setObjectName("button_next")
        self.horizontalLayout_14.addWidget(self.button_next)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label.setText(_translate("Dialog", "Row"))
        self.label_2.setText(_translate("Dialog", "Part"))
        self.label_3.setText(_translate("Dialog", "Manufacturer"))
        self.label_4.setText(_translate("Dialog", "Manufacturer ID"))
        self.label_5.setText(_translate("Dialog", "Distributor"))
        self.label_6.setText(_translate("Dialog", "Distributor ID"))
        self.label_7.setText(_translate("Dialog", "Description"))
        self.label_8.setText(_translate("Dialog", "Total Price(@1)"))
        self.label_9.setText(_translate("Dialog", "Total Price(@1k)"))
        self.label_10.setText(_translate("Dialog", "Price(@1)"))
        self.label_11.setText(_translate("Dialog", "Price(@1k)"))
        self.label_12.setText(_translate("Dialog", "Type"))
        self.button_choose.setText(_translate("Dialog", "Choose"))
        self.button_skip.setText(_translate("Dialog", "Skip"))
        self.label_13.setText(_translate("Dialog", "Parts"))
        self.button_previous.setText(_translate("Dialog", "Previous"))
        self.button_next.setText(_translate("Dialog", "Next"))



class notFoundDialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(400, 217)
        self.label_component = QtWidgets.QLabel(Dialog)
        self.label_component.setGeometry(QtCore.QRect(120, 20, 171, 16))
        self.label_component.setAlignment(QtCore.Qt.AlignCenter)
        self.label_component.setObjectName("label_component")
        self.label_2 = QtWidgets.QLabel(Dialog)
        self.label_2.setGeometry(QtCore.QRect(80, 110, 231, 17))
        self.label_2.setAlignment(QtCore.Qt.AlignCenter)
        self.label_2.setObjectName("label_2")
        self.line_component = QtWidgets.QLineEdit(Dialog)
        self.line_component.setGeometry(QtCore.QRect(82, 130, 231, 25))
        self.line_component.setAlignment(QtCore.Qt.AlignCenter)
        self.line_component.setObjectName("line_component")
        self.label_row = QtWidgets.QLabel(Dialog)
        self.label_row.setGeometry(QtCore.QRect(0, 40, 401, 16))
        self.label_row.setAlignment(QtCore.Qt.AlignCenter)
        self.label_row.setObjectName("label_row")
        self.label_value = QtWidgets.QLabel(Dialog)
        self.label_value.setGeometry(QtCore.QRect(0, 60, 401, 16))
        self.label_value.setAlignment(QtCore.Qt.AlignCenter)
        self.label_value.setObjectName("label_value")
        self.button_skip = QtWidgets.QPushButton(Dialog)
        self.button_skip.setGeometry(QtCore.QRect(100, 170, 89, 25))
        self.button_skip.setObjectName("button_skip")
        self.button_search = QtWidgets.QPushButton(Dialog)
        self.button_search.setGeometry(QtCore.QRect(210, 170, 89, 25))
        self.button_search.setObjectName("button_search")
        self.label_reference = QtWidgets.QLabel(Dialog)
        self.label_reference.setGeometry(QtCore.QRect(0, 80, 401, 16))
        self.label_reference.setAlignment(QtCore.Qt.AlignCenter)
        self.label_reference.setObjectName("label_reference")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_component.setText(_translate("Dialog", "Component not found."))
        self.label_2.setText(_translate("Dialog", "Do you want to alter the search?"))
        self.label_row.setText(_translate("Dialog", "Row: "))
        self.label_value.setText(_translate("Dialog", "Value: "))
        self.button_skip.setText(_translate("Dialog", "Skip"))
        self.button_search.setText(_translate("Dialog", "Search"))
        self.label_reference.setText(_translate("Dialog", "Reference: "))



############################################################################################################
#                                                                                                          #
#                                              Code                                                        #
#                                                                                                          #
############################################################################################################


class HomeWindow(QtWidgets.QMainWindow, MainWindow):
    def __init__(self, parent=None):
        super(HomeWindow, self).__init__(parent)
        self.resize(self.size().width(),self.size().height())
        self.setupUi(self)

        self.choose_part = PartWindow(self)
        self.not_found = NotFoundWindow(self)
        self.capacitor = Capacitor(self)
        self.resistor = Resistor(self)
        self.find_capacitor = self.capacitor.findCapacitor
        self.find_resistor = self.resistor.findResistor

        self.button_create_bom.clicked.connect(self.create_bom)
        self.button_load_file.clicked.connect(self.load_file)
        self.button_close.clicked.connect(self.close)

        self.file = ""
        self.board_name = "board"
        self.line_count = 6
        self.grouped = True
        self.part = ["", "", "", "", "", "", "", "www.digikey.com", "", "", "", "", "", ""]

        self.loop = QtCore.QEventLoop()
        self.choose_part.chosen.connect(self.loop.quit)
        self.capacitor.chosen.connect(self.loop.quit)
        self.resistor.chosen.connect(self.loop.quit)

        self.saved_part = {}
        # workbook = xlsxwriter.Workbook('output.xlsx')
        self.workbook = ""
        self.worksheet= ""
        self.cell_format= ""
        self.cell_header= ""
        self.cell_warning= ""

        self.package_table = {"1210" : "(3225 Metric)","1206":"(3216 Metric)","0603":"(1608 Metric)", "0805":"(2012 Metric)","0402":"(1005 Metric)","0201":"(0603 Metric)"}

        #Capacitor settings
        #Capacitor values should end with F
        #Examples: 10nF,0.1uF,100nF

        #Default settings for voltage rating of capacitors. The name of the column and the default voltage rating if the Voltage-rating field is empty.
        self.cap_v_rated_loc = 0
        self.cap_v_rated_name = "Voltage"
        self.def_cap_v_rated = 10
        #Default capacitor manufacturer, if it cant be found the fallback function will try to find a suitable capacitor from a different manufacturer
        self.def_cap_man = "Samsung Electro-Mechanics"

        #Resistance settings
        self.res_tol_loc = 0
        self.res_tol_name = "TOLERANCE"
        self.def_res_tol = 1

        self.def_res_man = "Yageo"

        #Diode settings
        self.def_diode_man = "Kingbright"

        #Location of the package column
        #Could be searched for
        self.package_loc = 4
        #Manufacturer-ID field
        self.mfm_id_name = "MFN-ID"
        self.mfm_id_loc = 0

        self.value_name = "VALUE"
        self.value_loc = 3
        #Manufacturer-name field, overrides the default capacitor and resistor manufacturer names
        self.mfm_name= "MANUFACTURER_NAME"
        self.mfm_loc = 0

        os.environ['DIGIKEY_CLIENT_ID'] = 'fvot4k1zYy8wXtpt3GSyIG0PDJcyODKG'
        os.environ['DIGIKEY_CLIENT_SECRET'] = 'd7K5BveNdGCqFRsM'
        os.environ['DIGIKEY_CLIENT_SANDBOX'] = 'False'
        os.environ['DIGIKEY_STORAGE_PATH'] = './'

    def load_file(self):
        self.line_file.clear()
        folder = QFileDialog.getExistingDirectory(self, 'Select folder where you want to save the BOM')
        if folder:
            self.line_file.setText(folder)

    def create_bom(self):
        self.button_create_bom.setEnabled(False)

        if self.line_file.text() != "":
            self.file = self.line_file.text()
        if self.line_board.text() != "":
            self.board_name = self.line_board.text()
        if self.line_capacitor_manufacturer.text() != "":
            self.def_cap_man = self.line_capacitor_manufacturer.text()
        if self.line_resistor_manufacturer.text() != "":
            self.def_res_man = self.line_resistor_manufacturer.text()
        if not self.radio_button_grouped.isChecked():
            self.grouped = False
        
        self.init_worksheet()
        counter = 1
        endline = 999

        try:
            grouped_components, individual_components = self.get_components()

            self.bom_header()
            self.bom_first_row()
            self.label_error.setText(f"Processing {len(grouped_components)} components.")

            for key, component in grouped_components.items():
                row = [counter, component['quantity'], ','.join(component['references']), component['value'], component['footprint'], "", "", component['DNP'], "", component['voltage']]
                self.label_error.setText(f"Processing component number {counter}.")
                counter += 1
                
                if self.line_count > 0 and self.line_count < endline:
                    self.process_row(row)                            
                if self.line_count == endline:
                    break

            for component in individual_components:
                row = [counter, component['quantity'], ','.join(component['references']), component['value'], "", "", "", component['DNP'], "", component['voltage']]
                counter += 1

                if self.line_count > 0 and self.line_count < endline:
                    self.process_row(row)                            
                if self.line_count == endline:
                    break

            self.bom_sum(self.line_count)
            self.label_error.setText(f"Processed {self.line_count - 6} components.")
            self.button_create_bom.setEnabled(True)
        except Exception as e:
            self.bom_sum(self.line_count)
            #logger.info("Error processing csv line ["+str(self.line_count)+"] can't continue")
            #logger.info(e)
            self.workbook.close()
            self.button_create_bom.setEnabled(True)

    def get_components(self):
        board = pcbnew.GetBoard()
        components = []

        for footprint in board.GetFootprints():
            reference = footprint.GetReference()
            if reference == "REF**":
                continue
            value = footprint.GetValue()
            fp = footprint.GetFPID().GetLibItemName()
            if (footprint.GetAttributes() & pcbnew.FP_DNP) != 0:
                dnp = "DNP"
            else:
                dnp = ""
            voltage = ""
        
            components.append({
            "reference": reference,
            "value": value,
            "DNP": dnp,
            "voltage": voltage,
            "footprint": str(fp)
            })

        grouped_components = {}
        individual_components = []

        for component in components:
            if component["value"] == "~":
                individual_components.append({
                    "references": [component["reference"]],
                    "value": component["value"],
                    "DNP": component["DNP"],
                    "quantity": 1,
                    "voltage": component["voltage"],
                    "footprint": component["footprint"]
                })
            else:
                key = (component["value"], component["DNP"])
                if key not in grouped_components:
                    grouped_components[key] = {
                        "references": [],
                        "value": component["value"],
                        "DNP": component["DNP"],
                        "quantity": 0,
                        "voltage": component["voltage"],
                        "footprint": component["footprint"]
                    }
                grouped_components[key]["references"].append(component["reference"])
                grouped_components[key]["quantity"] += 1

        return grouped_components, individual_components

    def process_row(self, row):
        #logger.info("Processing row: " + str(row))
        if self.grouped:
            self.bom_item(self.line_count, row, self.line_count - 5, row[2])
            if self.mfm_id_loc != 0 and row[self.mfm_id_loc] != "":
                part = row[self.mfm_id_loc]
                self.find_part_digikey(row, self.line_count)
            else:
                reference = row[2].split(',')[0]
                if reference.startswith("R", 0, 1) and not (reference[1:].isupper() or reference[1:].islower()):
                    #logger.info("ROW " + str(self.line_count + 4) + ": RESISTOR")
                    screen = self.find_resistor(row[3], self.line_count, row, reference)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")
                elif reference.startswith("C", 0, 1) and not (reference[1:].isupper() or reference[1:].islower()):
                    #logger.info("ROW " + str(self.line_count + 4) + ": CAPACITOR")
                    screen = self.find_capacitor(row[3], self.line_count, row, reference)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")
                else:
                    #logger.info("ROW " + str(self.line_count + 4) + ": COMPONENT")
                    screen = self.find_part_digikey(row, self.line_count)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")
            self.line_count += 1
        else:
            self.part = ["", "", "", "", "", "", "", "www.digikey.com", "", "", "", "", "", ""]
            references = row[2].split(',')
            part = references[0]
            self.bom_item(self.line_count, row, self.line_count - 5, part)
            if self.mfm_id_loc != 0 and row[self.mfm_id_loc] != "":
                self.find_part_digikey(row, self.line_count)
            else:
                if part.startswith("R", 0, 1) and not (part[1:].isupper() or part[1:].islower()):
                    #logger.info("ROW " + str(self.line_count + 4) + ": RESISTOR")
                    screen = self.find_resistor(row[3], self.line_count, row, part)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")
                elif part.startswith("C", 0, 1) and not (part[1:].isupper() or part[1:].islower()):
                    #logger.info("ROW " + str(self.line_count + 4) + ": CAPACITOR")
                    screen = self.find_capacitor(row[3], self.line_count, row, part)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")
                else:
                    #logger.info("ROW " + str(self.line_count + 4) + ": COMPONENT")
                    screen = self.find_part_digikey(row, self.line_count)
                    if screen:
                        #logger.info("loop")
                        self.loop.exec()
                    #logger.info("loop done")

            self.line_count += 1
            for i in range(1, len(references)):
                part = references[i].strip()
                self.bom_item(self.line_count, row, self.line_count - 5, part)
                self.copy_last_row(self.line_count)
                self.line_count += 1

    def bom_item(self, line_cnt, row, item_num, ref_design):
        self.worksheet.write(line_cnt, 0, item_num)                 # Item No.
        self.worksheet.write(line_cnt, 1, ref_design)               # Ref.Design
        if self.grouped:
            self.worksheet.write(line_cnt, 2, row[1])               # Quantity
        else:
            self.worksheet.write(line_cnt, 2, 1)
        self.worksheet.write(line_cnt, 3, row[7])                   # DNP
        self.worksheet.write(line_cnt, 4, row[3])                   # Value
        self.worksheet.write_url(line_cnt, 7, "www.digikey.com")    # Distributor

    def bom_first_row(self):
        headers = ["Item", "Ref.Design", "Quantity", "DNP", "Value", "Manufacturer", "Manufacturer ID", "Distributor", "Distributor ID",
                   "Description", "Type", "Price(@1)", "Price(@1k)", "Total Price(@1)", "Total Price(@1k)"]
        for i in range(0, len(headers)):
            self.worksheet.write(5, i, headers[i], self.cell_header)
            
    def bom_header(self):
        for i in range(0,4):
            for j in range(1, 7):
                self.worksheet.write(i,j,"",self.cell_header)
        self.worksheet.write(0, 1, "MicroBitDesign", self.cell_header)
        self.worksheet.write(0, 6, datetime.datetime.today().strftime('%Y.%m.%d.'), self.cell_header)
        self.worksheet.write(1, 1, self.board_name, self.cell_header)
        self.worksheet.write(2, 1, "Bill Of Materials", self.cell_header)

    def init_settings(self, row):
        try:
            self.cap_v_rated_loc = row.index(self.cap_v_rated_name) + 1
        except Exception as e:
            pass
            #logger.info(e)
        try:
            self.res_tol_loc= row.index(self.res_tol_name)
        except Exception as e:
            pass
            #logger.info(e)
        try:
            self.mfm_loc = row.index(self.mfm_name)
        except Exception as e:
            pass
            #logger.info(e)
        try:
            self.mfm_id_loc = row.index(self.mfm_id_name)
        except Exception as e:
            pass
            #logger.info(e)

    def bom_sum(self, line_cnt):
        self.worksheet.write(line_cnt + 2, 12, "SUM")
        self.worksheet.write_formula(line_cnt + 2, 13, '= SUM(L7:L' + str(line_cnt) + ')', self.cell_1_sum, '')
        self.worksheet.write_formula(line_cnt + 2, 14, '= SUM(M7:M' + str(line_cnt) + ')', self.cell_1k_sum, '')

        for i in range(6, line_cnt):
            self.worksheet.write_formula(i, 13, '= L' + str(i + 1) + '*C' + str(i + 1), value = '')
            self.worksheet.write_formula(i, 14, '= M' + str(i + 1) + '*C' + str(i + 1), value = '')

    def find_part_digikey(self, row, line_count):
        part = row[self.value_loc]
        reference = row[2].split(',')[0]
        loop = QtCore.QEventLoop()
        self.not_found.chosen.connect(loop.quit)
        if part in self.saved_part:
            #logger.info("Using existing part " + part)
            self.enter_values_to_csv(line_count, self.saved_part[part])
            return False
        else:
            try:
                search_request = KeywordSearchRequest(keywords = part, record_count = 10)
                result = digikey.keyword_search(body = search_request)
                #logger.info("Component search: ", part)
                #logger.info("Results: ", result.products_count)
                #logger.info("Products: ", len(result.products))
                if len(result.products) > 1 and result.products_count > 0:
                    self.choose_part.showWindow(result.products, line_count, reference)
                    #logger.info("SCREEN")
                    return True
                elif result.products_count == 1:
                    newl = str(result.products)
                    newl.replace('\'','\"')
                    js = json.loads(json.dumps(eval(newl), default = str))
                    self.enter_values_to_csv(line_count, js[0])
                    #logger.info("NO SCREEN, ONLY ONE PART")
                    return False
                else:
                    self.not_found.showWindow(line_count, row)
                    loop.exec()
                    #logger.info("NO RESULTS")

            except Exception as e:
                #logger.info("Error processing line ["+str(line_count)+"]")
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("Process Exception in line {}".format(line), e)
                for i in range(5,15):
                    self.worksheet.write(line_count,i,"",self.cell_format)
                return False

    def enter_values_to_csv(self, line_count, js):
        found_o = False
        found_k = False

        # #logger.info("ENTERING VALUES TO CSV:\n", js)

        try:
            # #logger.info("TRYING TO FIND MANUFACTURER")
            self.part[5] = js['manufacturer']['value']
            self.worksheet.write(line_count, 5, self.part[5])  # Manufacturer
            # #logger.info("MANUFACTURER: ", js['manufacturer']['value'])
        except KeyError:
            #logger.info("KeyError: 'manufacturer' or 'value' key not found")
            self.part[5] = ""
            self.worksheet.write(line_count, 5, self.part[5])  # Handle missing key gracefully
        
        try:
            self.part[6] = js['manufacturer_part_number']
            self.worksheet.write(line_count, 6, self.part[6])  # Manufacturer ID
        except KeyError:
            #logger.info("KeyError: 'manufacturer_part_number' key not found")
            self.part[6] = ""
            self.worksheet.write(line_count, 5, self.part[6])  # Handle missing key gracefully
        
        try:
            self.part[8] = js['digi_key_part_number']
            self.worksheet.write(line_count, 8, self.part[8])  # Distributor ID
        except KeyError:
            #logger.info("KeyError: 'digi_key_part_number' key not found")
            self.part[8] = ""
            self.worksheet.write(line_count, 8, self.part[8])  # Handle missing key gracefully
        
        try:
            self.part[9] = js['product_description']
            self.worksheet.write(line_count, 9, self.part[9])  # Description
        except KeyError:
            #logger.info("KeyError: 'product_description' key not found")
            self.part[9] = ""
            self.worksheet.write(line_count, 9, self.part[9])  # Handle missing key gracefully

        one_price = -1
        try:
            for param in js['parameters']:
                if param['parameter'] == "Package / Case":
                    self.part[10] = param['value']
                    self.worksheet.write(line_count, 10, self.part[10])  # Type
        except (KeyError, TypeError) as e:
            #logger.info(f"Error accessing 'parameters': {e}")
            self.part[10] = ""
            self.worksheet.write(line_count, 10, self.part[10])  # Handle missing key gracefully

        try:
            for pricing in js['standard_pricing']:
                if pricing['break_quantity'] == 1:
                    one_price = pricing['unit_price']
                    self.part[11] = one_price
                    self.worksheet.write(line_count, 11, one_price)  # Price(@1)
                    found_o = True

                if pricing['break_quantity'] == 1000:
                    one_k_price = pricing['unit_price']
                    self.part[12] = one_k_price
                    self.worksheet.write(line_count, 12, one_k_price)  # Price(@1k)
                    found_k = True
        except (KeyError, TypeError) as e:
            pass
            #logger.info(f"Error accessing 'standard_pricing': {e}")

        if one_price == -1:
            try:
                one_price = js['unit_price']
                self.part[11] = one_price
                self.worksheet.write(line_count, 11, one_price)  # Price(@1)
            except KeyError:
                #logger.info("KeyError: 'unit_price' key not found")
                one_price = 0  # Default value if key is missing
                self.part[11] = 0

        if not found_k:
            try:
                last_price = js['standard_pricing'][-1]['unit_price']
                last_quant = js['standard_pricing'][-1]['break_quantity']
                diff = one_price - last_price
                rise = diff / float(last_quant)
                last_price = rise * 1000
                last_price = one_price - last_price

                if last_price <= 0:
                    last_price = js['standard_pricing'][-1]['unit_price']

                self.part[12] = last_price
                self.worksheet.write(line_count, 12, last_price)
                found_k = True
            except (KeyError, IndexError, TypeError) as e:
                #logger.info(f"Error calculating last_price: {e}")
                self.part[12] = ""
                self.worksheet.write(line_count, 12, "", self.cell_format)

        if not found_o:
            self.part[11] = ""
            self.worksheet.write(line_count, 11, "", self.cell_format)
        
        if not found_k:
            self.part[12] = ""
            self.worksheet.write(line_count, 12, "", self.cell_format)

    def copy_last_row(self, line_count):
        if self.part[8] != "":
            for i in range(5, 13):
                self.worksheet.write(line_count, i, self.part[i])
        else:
            self.enter_empty_row_to_csv(line_count)

    def enter_empty_row_to_csv(self, line_count):
        for i in range(5, 15):
            self.worksheet.write(line_count, i, "", self.cell_format)

    def init_worksheet(self):
        date = datetime.datetime.today().strftime("%Y-%m-%d-%H-%M-%S")
        if self.file == "":
            name = "~/" + self.board_name + "_" + date + "_BOM.xls"
        else:
            name = self.file + "/" + self.board_name + "_" + date + "_BOM.xls"

        self.workbook = xlsxwriter.Workbook(name)
        self.worksheet = self.workbook.add_worksheet()

        widths = [0.3, 1.0, 0.3, 0.3, 0.4, 0.5, 0.8, 0.6, 1.0, 0.4, 0.4, 0.4, 0.4, 0.4]
        for i in range(len(widths)):
            self.worksheet.set_column(i, i, widths[i] * 25.4)

        self.cell_format = self.workbook.add_format()
        self.cell_format.set_bg_color("#d61c1c")

        self.cell_header = self.workbook.add_format()
        self.cell_header.set_bg_color("#595959")
        self.cell_header.set_font_color("white")
        self.cell_header.set_bold()

        self.cell_warning = self.workbook.add_format()
        self.cell_warning.set_bg_color("#d6631c")

        self.cell_1_sum = self.workbook.add_format({"num_format": "0.00"})
        self.cell_1_sum.set_bg_color("#ffff66")

        self.cell_1k_sum = self.workbook.add_format({"num_format": "0.00"})
        self.cell_1k_sum.set_bg_color("#00cc33")

    def closeEvent(self, event):
        try:
            self.workbook.close()
        except Exception as e:
            pass


class PartWindow(QtWidgets.QDialog, partWindowDialog):
    chosen = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super(PartWindow, self).__init__(parent)
        self.resize(self.size().width(),self.size().height())
        self.setupUi(self)

        self.parent = parent
        self.data = []
        self.line_count = 0
        self.part = ""
        self.current_index = 0

        self.combobox_parts.currentTextChanged.connect(self.changed_part)
        self.button_skip.clicked.connect(self.skip_part)
        self.button_choose.clicked.connect(self.choose_part)
        self.button_next.clicked.connect(self.next_part)
        self.button_previous.clicked.connect(self.previous_part)

    def showWindow(self, data, line_count, part):
        self.data = data
        self.line_count = line_count
        self.part = part
        self.show()
        self.fill_combobox()

    def fill_combobox(self):
        self.combobox_parts.clear()
        for i in range(len(self.data)):
            newl = str(self.data[i])
            newl.replace('\'','\"')
            js = json.loads(json.dumps(eval(newl), default = str))
            self.combobox_parts.addItem(js['digi_key_part_number'])

        self.fill_data()

    def fill_data(self):
        self.line_row.setText(str(self.line_count - 5))
        self.line_part.setText(str(self.part))

        if self.data != None:
            newl = str(self.data[self.current_index])
            newl.replace('\'','\"')
            js = json.loads(json.dumps(eval(newl), default=str))

            one_price = -1
            found_o = False
            found_k = False

            try:
                self.line_manufacturer.setText(js['manufacturer']['value'])
            except Exception as e:
                self.line_manufacturer.setText("Unknown")

            try:
                self.line_manufacturer_id.setText(js['manufacturer_part_number'])
            except Exception as e:
                self.line_manufacturer_id.setText("Unknown")
                
            self.line_distributor.setText("www.digikey.com")

            try:
                self.line_distributor_id.setText(js['digi_key_part_number'])
            except Exception as e:
                self.line_distributor_id.setText("Unknown")

            try:
                self.line_description.setText(js['product_description'])
            except Exception as e:
                self.line_description.setText("Unknown")

            try:
                self.line_type.setText(js['parameters'][0]['value'])
            except Exception as e:
                self.line_type.setText("Unknown")

            for n in range(0, len(js['standard_pricing'])):
                if js['standard_pricing'][n]['break_quantity'] == 1:
                    try:
                        one_price = js['standard_pricing'][n]['unit_price']
                        self.line_price_1.setText(str(one_price))
                        found_o = True
                    except Exception as e:
                        self.line_price_1.setText("Unknown")

                if js['standard_pricing'][n]['break_quantity'] == 1000:
                    try:
                        one_k_price = js['standard_pricing'][n]['unit_price']
                        self.line_price_1k.setText(str(one_k_price))
                        found_k = True
                    except Exception as e:
                        self.line_price_1k.setText("Unknown")

            if found_o == False:
                try:
                    one_price = js['unit_price']
                    self.line_price_1.setText(str(one_price))
                except Exception as e:
                    self.line_price_1.setText("Unknown")
            
            if found_k == False:
                try:
                    last_price = js['standard_pricing'][len(js['standard_pricing']) - 1]['unit_price']
                    last_quant = js['standard_pricing'][len(js['standard_pricing']) - 1]['break_quantity']
                    diff = one_price - last_price
                    rise = diff / float(last_quant)
                    last_price = rise * 1000
                    last_price = one_price - last_price

                    if last_price <= 0:
                        last_price = js['standard_pricing'][len(js['standard_pricing']) - 1]['unit_price']

                    self.line_price_1k.setText(str(last_price))
                    found_k = True
                except Exception as e:
                    self.line_price_1k.setText("Unknown")


    def changed_part(self):
        self.current_index = self.combobox_parts.currentIndex()
        self.fill_data()

    def skip_part(self):
        self.parent.enter_empty_row_to_csv(self.line_count)
        self.close()

    def choose_part(self):
        if self.data != None:
            self.current_index = self.combobox_parts.currentIndex()
            newl = str(self.data[self.current_index])
            newl.replace('\'','\"')
            js = json.loads(json.dumps(eval(newl), default = str))
            self.parent.enter_values_to_csv(self.line_count, js)
        else:
            self.skip_part()

        self.close()

    def next_part(self):
        self.current_index = (self.current_index + 1) % len(self.data)
        self.combobox_parts.setCurrentIndex(self.current_index)
        self.fill_data()

    def previous_part(self):
        self.current_index = (self.current_index - 1) % len(self.data)
        self.combobox_parts.setCurrentIndex(self.current_index)
        self.fill_data()

    def closeEvent(self, event):
        self.chosen.emit()

class NotFoundWindow(QtWidgets.QDialog, notFoundDialog):
    chosen = QtCore.pyqtSignal()

    def __init__(self, parent=None):
        super(NotFoundWindow, self).__init__(parent)
        self.resize(self.size().width(),self.size().height())
        self.setupUi(self)

        self.parent = parent
        self.line_count = 0
        self.reference = ""
        self.value = ""
        self.row = []

        self.button_skip.clicked.connect(self.skip_part)
        self.button_search.clicked.connect(self.search_part)

    def showWindow(self, line_count, row):
        self.line_count = line_count
        self.reference = row[2]
        self.value = row[3]
        self.row = row
        self.show()
        self.fill_data()

    def fill_data(self):
        self.label_row.setText("Row: " + str(self.line_count - 5))
        self.label_value.setText("Value: " + str(self.value))
        self.label_reference.setText("Reference: " + str(self.reference))
        self.line_component.setText("")

    def skip_part(self):
        self.parent.enter_empty_row_to_csv(self.line_count)
        self.close()

    def search_part(self):
        if self.line_component.text() != "":
            part = self.line_component.text()
            self.row[3] = part
            self.parent.find_part_digikey(self.row, self.line_count)
        else:
            self.parent.enter_empty_row_to_csv(self.line_count)

        self.close()

    def closeEvent(self, event):
        self.chosen.emit()


class Capacitor(QtCore.QObject):
    chosen = QtCore.pyqtSignal()

    def __init__(self, parent = None):
        super(Capacitor, self).__init__(parent)
        self.parent = parent
        self.pf_table = { "u" :1000000, "p":1, "n":1000}
        self.saved_caps = {}

    def findCapacitor(self, capacitor_value, line_count, row, part_name):
        capacitor_value = self.sanitizeCapValue(capacitor_value,line_count)
        new_cap_value = self.capacitanceConversion(capacitor_value)
        saved_name = self.mapNameCap(capacitor_value,row)
        if (saved_name) in self.saved_caps:
            #logger.info("-> CAP: Using existing capacitor "+saved_name)
            self.parent.enter_values_to_csv(line_count, self.saved_caps[saved_name])
        else:
            #logger.info("Try to find capacitor")
            try:
                search_request = self.getSearchQuery(row, new_cap_value)
                result = digikey.keyword_search(body = search_request)
                if result is not None:
                    pass
                    #logger.info("result is not none")
                    #logger.info("-> CAP: results = "+ str(result.products_count))
                    #logger.info("-> CAP: products len ="+ str(len(result.products)))
                if len(result.products) > 1 and result.products_count > 0:
                    self.parent.choose_part.showWindow(result.products, line_count, part_name)
                    #logger.info("SCREEN")
                    return True
                elif result.products_count == 1:
                    self.parent.enter_values_to_csv(line_count, result.products)
                    #logger.info("NO SCREEN, ONLY ONE PART")
                    return False
                else:
                    return self.fallbackCapacitor(row, capacitor_value, line_count, part_name)

            except Exception as e:
                #logger.info("-> CAP: error processing csv line ["+str(line_count)+"]")
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> CAP: process Exception in line {}".format(line), e)
                for i in range(5,15):
                    self.parent.worksheet.write(line_count,i,"",self.parent.cell_format)
                # self.chosen.emit()
                #logger.info("emit")
            
            return False

    def filterCapacitance(self, js,capacitor_value,row):
        # self.parent.cap_v_rated_loc = row.index(self.parent.cap_v_rated)
        voltage_rated = self.parent.def_cap_v_rated
        if(self.parent.cap_v_rated_loc != 0 and row[self.parent.cap_v_rated_loc] != ""):
            row_v_rated = row[self.parent.cap_v_rated_loc]
            if(row_v_rated.endswith("V")):
                row_v_rated = row_v_rated[:len(row_v_rated)-1]
            voltage_rated = float(row_v_rated)
        found_res = False
        found_rated = False
        for n in range(0,len(js['parameters'])):
            if(js['parameters'][n]['parameter']=="Capacitance"):
                query_res = js['parameters'][n]['value']
                # #logger.info("Resistor_value before changes = "+ resistor_value)
                query_res = query_res.replace(" ","")
                #logger.info("-> CAP: capacitance query = "+query_res)
                if "u" in capacitor_value:
                    u_index = capacitor_value.index("u")

                    capacitor_value = capacitor_value[:u_index] + "µ" + capacitor_value[u_index+1:len(capacitor_value)]
                #logger.info("-> CAP: capacitance internal = "+capacitor_value)

                # #logger.info("Fn-result = "+fn_res)
                if(capacitor_value == query_res):
                    # #logger.info("Cap match")
                    found_res = True
            if(js['parameters'][n]['parameter']=="Voltage - Rated"):
                rated = js['parameters'][n]['value']
                rated = rated[:len(rated)-1]
                #logger.info("-> CAP: voltage-rated = "+ rated)
                #logger.info("-> CAP: voltage-rated row = "+ str(voltage_rated))
                v_rat =float(rated)
                if(v_rat >= voltage_rated):
                    found_rated = True

        return found_res and found_rated

    def sanitizeCapValue(self, capacitor_value : str,line_count):
        if capacitor_value.startswith("C"):
            return "0F"
        if capacitor_value.endswith("F") == False:
            capacitor_value = capacitor_value +"F"
            self.parent.worksheet.write(line_count,4,capacitor_value)
        #logger.info("-> CAP: sanitizeCapValue = "+capacitor_value)
        return capacitor_value

    def getCapacitanceFilter(self, capacitor_value):
            if "u" in str(capacitor_value):
                u_index = str(capacitor_value).index("u")
                capacitor_value = str(capacitor_value)[:u_index] + "µ" + str(capacitor_value)[u_index+1:len(str(capacitor_value))]
            # fn_res = "u"+capacitor_value
            fn_res = capacitor_value
            #logger.info("-> CAP: fn_res = "+ str(fn_res))
            taxo_ids = []
            man_ids = []
            parametric_filter = ParametricFilter(2049,str(fn_res))
            par_filter = [parametric_filter]
            res_filter = Filters(taxo_ids,man_ids,par_filter)
            return res_filter

    def getSearchQuery(self, row, capacitor_value):
            cap_man = self.parent.def_cap_man
            if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
                cap_man = row[self.parent.mfm_loc]
            package = row[self.parent.package_loc]
            m = re.search(r"[0-9]", package)
            package = package[m.start():]
            cap_search = cap_man+" CAP CER "+capacitor_value
            #logger.info("Package: ", package)
            try:
                package = package + " " + self.parent.package_table[package]
            except Exception as e:
                #logger.info(cap_search)
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> CAP: process Exception in line {}".format(line), e)
                raise Exception(e)
            cap_search = cap_man + " CAP CER " + capacitor_value + " " + package

            # CAP CER 0.1UF 16V X7R 0603
            cap_filter = self.getCapacitanceFilter(capacitor_value)
            # search_request = KeywordSearchRequest(keywords=cap_search, record_count=20)
            #logger.info("Capacitor search query: ", cap_search)
            search_request = KeywordSearchRequest(keywords=cap_search, record_count=20,filters=cap_filter)
            return search_request

    def fallbackCapacitor(self, row, capacitor_value, line_count, part_name):
        #logger.info("-> CAP: searching for a fallback capacitor")
        package = row[self.parent.package_loc]
        m = re.search(r"[0-9]",package)
        package = package[m.start():]
        capacitor_search=" CAP CER "+str(capacitor_value)
        try:
            package = str(package) + " " + self.parent.package_table[package]
        except Exception as e:
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            #logger.info("-> CAP: process Exception in line {}".format(line), e)
            #logger.info(capacitor_search)
            raise Exception(e)
        capacitor_search=" CAP CER "+str(capacitor_value)+" "+package
        cap_filter = self.getCapacitanceFilter(capacitor_value)
        search_request = KeywordSearchRequest(keywords=capacitor_search, record_count=20,filters=cap_filter)
        result = digikey.keyword_search(body=search_request)
        #logger.info("Fallback capacitor search: ", search_request)
        #logger.info("-> CAP: results = "+ str(result.products_count))
        #logger.info("-> CAP: products len ="+ str(len(result.products)))
        if len(result.products) > 1 and result.products_count > 0:
            self.parent.choose_part.showWindow(result.products, line_count, part_name)
            #logger.info("SCREEN - FALLBACK CAPACITORS")
            return True
        elif result.products_count == 1:
            self.parent.enter_values_to_csv(line_count, result.products)
            #logger.info("NO SCREEN - ONLY ONE FALLBACK CAPACITOR")
            return False
        else:
            self.parent.enter_empty_row_to_csv(line_count)
            #logger.info("NO FALLBACK CAPACITORS")
            return False

    def capacitanceConversion(self, capacitor_value):
        newcap = capacitor_value
        if "nF" in capacitor_value:
            newcap = capacitor_value[:len(capacitor_value)-2]
            newcap = float(newcap)
            if newcap < 100:
                newcap = newcap * 1000
                newcap = str(int(newcap)) + "pF"
            else:
                newcap = newcap /1000
                newcap = str((newcap)) + "uF"
        #logger.info("-> CAP: new capacitance value: " + newcap)
        return newcap

    def mapNameCap(self, capacitor_value,row):
        cap_man = self.parent.def_cap_man
        if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
            cap_man = row[self.parent.mfm_loc]
        voltage_rated = self.parent.def_cap_v_rated
        if(self.parent.cap_v_rated_loc != 0 and row[self.parent.cap_v_rated_loc] != ""):
            voltage_rated = row[self.parent.cap_v_rated_loc]
        saved_name = capacitor_value+" "+row[self.parent.package_loc]+" "+str(voltage_rated)+" "+str(cap_man)
        return saved_name
    

class Resistor(QtCore.QObject):
    chosen = QtCore.pyqtSignal()

    def __init__(self, parent = None):
        super(Resistor, self).__init__(parent)
        self.parent = parent
        
        self.saved_res = {}

    def findResistor(self, resistor_value, line_count, row, part_name):
        resistor_value = resistor_value.upper()
        resistor_value = self.sanitizeResValue(resistor_value)
        saved_name = self.mapNameRes(resistor_value,row)

        if saved_name in self.saved_res:
            #logger.info("-> RES: Using existing resistor "+saved_name)
            self.parent.enter_values_to_csv(line_count,self.saved_res[saved_name])
        else:
            #logger.info("Try to find resistor")
            try:
                search_request = self.getSearchQuery(row,resistor_value)
                result = digikey.keyword_search(body=search_request)
                #logger.info("-> RES: results = "+ str(result.products_count))
                #logger.info("-> RES: products len ="+ str(len(result.products)))
                if len(result.products) > 1 and result.products_count > 0:
                    self.parent.choose_part.showWindow(result.products, line_count, part_name)
                    return True
                elif result.products_count == 1:
                    self.parent.enter_values_to_csv(line_count, result.products)
                    return False
                else:
                    return self.fallbackResistor(row, resistor_value, line_count, part_name)

            except Exception as e:
                #logger.info("-> RES: error processing csv line ["+str(line_count)+"]")
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> RES: process Exception in line {}".format(line), e)
                for i in range(5,15):
                    self.parent.worksheet.write(line_count,i,"",self.parent.cell_format)
                self.chosen.emit()

            return False

    def filterResistance(self, js,resistor_value,row,tolerance_check):
        resistor_tolerance = self.parent.def_res_tol
        if(self.parent.res_tol_loc != 0 and row[self.parent.res_tol_loc] != ""):
            row_tol = row[self.parent.res_tol_loc]
            row_tol = row_tol[:len(row_tol)-1]
            resistor_tolerance = float(row_tol)
        found_res = False
        found_tol = False
        for n in range(0,len(js['parameters'])):
            if(js['parameters'][n]['parameter']=="Resistance"):
                query_res = js['parameters'][n]['value']
                # #logger.info("Resistor_value before changes = "+ resistor_value)
                #logger.info("-> RES: resistance = "+query_res)

                ohm_msg = "Ohms"
                if(resistor_value.endswith("K")):
                    resistor_value = resistor_value[:len(resistor_value)-1]
                    ohm_msg = "k"+ohm_msg
                if(resistor_value.endswith("M")):
                    resistor_value = resistor_value[:len(resistor_value)-1]
                    ohm_msg = "M"+ohm_msg
                fn_res = resistor_value + " "+ohm_msg
                # #logger.info("Fn-result = "+fn_res)
                if(fn_res == query_res):
                    # #logger.info("Resistance match")
                    found_res = True
            if tolerance_check == True:
                if(js['parameters'][n]['parameter']=="Tolerance"):
                    query_res = js['parameters'][n]['value']
                    if( query_res == "Jumper" and int(resistor_value) == 0):
                        found_tol = True
                    else:
                        query_res = query_res[1:len(query_res)-1]
                        #logger.info("-> RES: tolerance = "+query_res)
                        res_tol_query = float(query_res)
                        if(res_tol_query == resistor_tolerance):
                            found_tol = True
            else:
                found_tol = True

        return found_res and found_tol

    def sanitizeResValue(self, resistor_value):
        # val = resistor_value[0:len(resistor_value)-1]
        # 3.3k 3k3 30k 30
        val = resistor_value
        if(not(resistor_value.endswith("K"))):
            if "K" in val:
                point_ind = val.index("K")
                val = val[:point_ind] + "." + val[point_ind+1:]
                val = val+"K"
        # #logger.info("sanitizeResValue = "+val)
        return val

    def getResistanceFilter(self, resistor_value):
            ohm_msg = "Ohms"
            if(resistor_value.endswith("K")):
                resistor_value = resistor_value[:len(resistor_value)-1]
                ohm_msg = "k"+ohm_msg
            # fn_res = "u"+resistor_value + " "+ohm_msg
            fn_res = resistor_value + " "+ohm_msg

            taxo_ids = []
            man_ids = []
            parametric_filter = ParametricFilter(2085,fn_res)
            par_filter = [parametric_filter]
            res_filter = Filters(taxo_ids,man_ids,par_filter)
            return res_filter

    def getSearchQuery(self, row,resistor_value):
            res_man = self.parent.def_res_man
            if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
                res_man = row[self.parent.mfm_loc]
            package = row[self.parent.package_loc]
            m = re.search(r"[0-9]",package)
            package = package[m.start():]
            # resistor_search = res_man+" RES SMD "+resistor_value+" OHM 1/10W "
            resistor_search = res_man+" RES SMD "+resistor_value+" OHM "
            try:
                package = package + " " + self.parent.package_table[package]
            except Exception as e:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> RES: process Exception in line {}".format(line), e)
                #logger.info(resistor_search)
                raise Exception(e)
            resistor_search = res_man+" RES SMD "+resistor_value+" OHM  " + package
            #logger.info(resistor_search)
            res_filter = self.getResistanceFilter(resistor_value)
            search_request = KeywordSearchRequest(keywords=resistor_search, record_count=20, filters=res_filter)
            return search_request

    def fallbackResistor(self, row, resistor_value, line_count, part_name):
        #logger.info("-> RES: searching for a fallback resistor")
        package = row[self.parent.package_loc]
        m = re.search(r"[0-9]",package)
        package = package[m.start():]
        resistor_search ="RES SMD "+resistor_value+" OHM 1/10W "
        try:
            package = package + " " + self.parent.package_table[package]
        except Exception as e:
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            #logger.info("-> RES: process Exception in line {}".format(line), e)
            #logger.info(resistor_search)
            raise Exception(e)
        resistor_search ="RES SMD "+resistor_value+" OHM " + package
        res_filter = self.getResistanceFilter(resistor_value)
        search_request = KeywordSearchRequest(keywords=resistor_search, record_count=20, filters=res_filter)
        result = digikey.keyword_search(body=search_request)
        #logger.info("Fallback resistor search: ", search_request)
        #logger.info("-> RES: results = "+ str(result.products_count))
        #logger.info("-> RES: products len ="+ str(len(result.products)))

        if len(result.products) > 1 and result.products_count > 0:
            self.parent.choose_part.showWindow(result.products, line_count, part_name)
            #logger.info("SCREEN - FALLBACK RESISTORS")
            return True
        elif result.products_count == 1:
            self.parent.enter_values_to_csv(line_count, result.products)
            #logger.info("NO SCREEN - ONLY ONE FALLBACK RESISTOR")
            return False
        else:
            self.parent.enter_empty_row_to_csv(line_count)
            #logger.info("NO FALLBACK RESISTORS")
            return False

    def mapNameRes(self, resistor_value,row):
        res_man = self.parent.def_res_man
        if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
            res_man = row[self.parent.mfm_loc]
        resistor_tolerance = self.parent.def_res_tol
        if(self.parent.res_tol_loc != 0 and row[self.parent.res_tol_loc] != ""):
            resistor_tolerance = row[self.parent.res_tol_loc]
        saved_name = (resistor_value+" "+row[self.parent.package_loc]+" "+str(resistor_tolerance)+" "+res_man)
        return saved_name


class Diode():
    def __init__(self, parent = None):
        self.saved_diodes = {}

    def findDiode(self, diode_color, line_count, row):
        saved_name = self.mapNameDiode(diode_color,row)
        if saved_name in self.saved_diodes:
            #logger.info("-> LED: using existing diode "+saved_name)
            self.parent.enter_values_to_csv(line_count, self.saved_diodes[saved_name])
        else:
            try:
                search_request = self.getSearchQuery(row,diode_color)
                result = digikey.keyword_search(body=search_request)
                found_in_search = False
                found_item = False
                #logger.info("-> LED: results = "+ str(result.products_count))
                #logger.info("-> LED: products len ="+ str(len(result.products)))
                for n in range(0,len(result.products)):
                    # #logger.info(result.products[n])
                    newl = str(result.products[n])
                    newl.replace('\'','\"')
                    found_item = False
                    js = json.loads(json.dumps(eval(newl)))
                    packaging_value =js['packaging']['value']
                    if(packaging_value == 'Cut Tape (CT)' or packaging_value == 'Bulk'):
                        found_item = self.filterDiode(js,diode_color)
                        if(found_item == False):
                            continue
                        found_in_search = True
                        self.parent.enter_values_to_csv(line_count,js)
                        #logger.info("-> LED: saved diode color "+saved_name)
                        self.saved_diodes[saved_name] = js
                        break
                if(result.products_count == 0 or found_in_search==False):
                    found_in_search = self.fallbackDiode(row,diode_color,line_count)
                    if(found_in_search == False):
                        for i in range(4,12):
                            self.parent.worksheet.write(line_count,i,"",self.parent.cell_format)

            except Exception as e:
                #logger.info("-> LED: error processing csv line ["+str(line_count)+"]")
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> LED: process Exception in line {}".format(line), e)
                for i in range(4,12):
                    self.parent.worksheet.write(line_count,i,"",self.parent.cell_format)

    def getSearchQuery(self, row,diode_color):
            diode_man = self.parent.def_diode_man
            if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
                diode_man = row[self.parent.mfm_loc]
            package = row[self.parent.package_loc]
            m = re.search(r"[0-9]",package)
            # #logger.info(m.start())
            package = package[m.start():]
            diode_search = diode_man+" LED "+diode_color+" CLEAR CHIP SMD"
            # LED GREEN CLEAR CHIP SMD
            try:
                package = package + " " + self.parent.package_table[package]
            except Exception as e:
                trace_back = sys.exc_info()[2]
                line = trace_back.tb_lineno
                #logger.info("-> LED: process Exception in line {}".format(line), e)
                #logger.info(diode_search)
                raise Exception(e)
            diode_search = diode_man+" LED "+diode_color+" CLEAR CHIP SMD " + package
            #logger.info(diode_search)
            search_request = KeywordSearchRequest(keywords=diode_search, record_count=20)
            return search_request


    def filterDiode(self, js,diode_color):
        found_color = False
        for n in range(0,len(js['parameters'])):
            if(js['parameters'][n]['parameter']=="Color"):
                query_res = js['parameters'][n]['value']
                # #logger.info("Resistor_value before changes = "+ resistor_value)
                query_res = query_res.upper()
                #logger.info("-> LED: color = "+query_res)

                # #logger.info("Fn-result = "+fn_res)
                if(diode_color == query_res):
                    # #logger.info("Resistance match")
                    found_color = True

        return found_color


    def fallbackDiode(self, row,diode_color,line_count):
        #logger.info("-> LED: searching for a fallback diode")
        package = row[self.parent.package_loc]
        m = re.search(r"[0-9]",package)
        # #logger.info(m.start())
        package = package[m.start():]
        capacitor_search="LED "+diode_color
        try:
            package = package + " " + self.parent.package_table[package]
        except Exception as e:
            trace_back = sys.exc_info()[2]
            line = trace_back.tb_lineno
            #logger.info("-> LED: process Exception in line {}".format(line), e)
            #logger.info(capacitor_search)
            raise Exception(e)
        capacitor_search="LED "+diode_color + " " +package
        search_request = KeywordSearchRequest(keywords=capacitor_search, record_count=20)
        result = digikey.keyword_search(body=search_request)
        found_in_search = False
        found_item = False
        #logger.info("-> LED: results = "+ str(result.products_count))
        #logger.info("-> LED: products len ="+ str(len(result.products)))
        for n in range(0,len(result.products)):
            # #logger.info(result.products[n])
            newl = str(result.products[n])
            newl.replace('\'','\"')
            found_item = False
            js = json.loads(json.dumps(eval(newl)))
            if(js['packaging']['value'] == 'Cut Tape (CT)' or js['packaging']['value'] == 'Bulk'):
                found_item = self.filterDiode(js,diode_color)
                if(found_item == False):
                    continue
                found_in_search = True
                self.parent.enter_values_to_csv(line_count,js)
                break
        return found_in_search

    def mapNameDiode(self, diode_color,row):
        diode_man = self.parent.def_diode_man
        if(row[self.parent.mfm_loc]!="" and self.parent.mfm_loc !=0):
            diode_man = row[self.parent.mfm_loc]
        saved_name = diode_color+" "+row[self.parent.package_loc]+" "+str(diode_man)
        return saved_name
import os
import sys
import requests
import pandas as pd
import traceback
from PIL import Image
import photoshop.api as ph
from photoshop import Session
from psd_tools import PSDImage
from PyQt5.QtWidgets import QMainWindow, QFileDialog
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QCoreApplication
from UliPlot.XLSX import auto_adjust_xlsx_column_width
# from PyQt5 import uic
# import numpy as np
# import pyperclip
# import tkinter as tk
# from tkinter import filedialog
# from openpyxl import load_workbook
# import xlsxwriter
# import pyautogui
# import shutil

################################################################################################################################
################################################################################################################################

################################################################################################################################
# Folders for exported data
################################################################################################################################
desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
export_folder = os.path.join(desktop, "!BulkSI\\Export SI")
export_folder_excel = os.path.join(export_folder, "_Excel Files_")
export_folder_exportfiles = os.path.join(export_folder, "_Export Files_")
errors_list = []

################################################################################################################################
# UI
################################################################################################################################
class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1000, 800)
        MainWindow.setStyleSheet("background: #222222;\n"
"")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setEnabled(True)
        self.centralwidget.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.centralwidget.setAutoFillBackground(False)
        self.centralwidget.setObjectName("centralwidget")




        # IMAGES
        self.instruction1 = QtWidgets.QLabel(self.centralwidget)
        self.instruction1.setGeometry(QtCore.QRect(40, 40, 1000, 50))
        self.instruction1.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.instruction1.setObjectName("instruction1")
        self.id1 = QtWidgets.QLabel(self.centralwidget)
        self.id1.setGeometry(QtCore.QRect(40, 110, 30, 30))
        self.id1.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.id1.setObjectName("id1")
        self.img_label = QtWidgets.QTextEdit(self.centralwidget)
        self.img_label.setGeometry(QtCore.QRect(70, 110, 330, 50))
        self.img_label.setStyleSheet("QTextEdit {\n"
                                 "color: #FFFFFF !important;\n"
                                 "font: 7pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "background: #494949;\n"
                                 "placeholder: 'asdasda';\n"
                                 "padding: 0px;\n"
                                 "border: 4px solid #494949 !important;\n"
                                 "display: inline-block;\n"
                                 "}")
        self.img_label.setObjectName("img_label")


        self.img_toolButton = QtWidgets.QPushButton(self.centralwidget)
        self.img_toolButton.setGeometry(QtCore.QRect(400, 110, 50, 50))
        self.img_toolButton.setStyleSheet("QPushButton {\n"
                                      "color: #FFFFFF !important;\n"
                                      "font: 12pt \'Century Gothic\';\n"
                                      "text-transform: uppercase;\n"
                                      "text-decoration: none;\n"
                                      "background: #A9742A;\n"
                                      "padding: 0px;\n"
                                      "border: 4px solid #494949 !important;\n"
                                      "display: inline-block;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "QPushButton:hover {\n"
                                      "color: #ffffff !important;\n"
                                      "background: #f6b93b;\n"
                                      "border-color: #f6b93b !important;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}")
        self.img_toolButton.setObjectName("img_toolButton")


        self.id2 = QtWidgets.QLabel(self.centralwidget)
        self.id2.setGeometry(QtCore.QRect(40, 180, 30, 30))
        self.id2.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.id2.setObjectName("id2")
        self.img_pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.img_pushButton.setGeometry(QtCore.QRect(70, 180, 380, 41))
        self.img_pushButton.setStyleSheet("QPushButton {\n"
                                      "color: #FFFFFF !important;\n"
                                      "font: 12pt \'Century Gothic\';                                   \n"
                                      "text-transform: uppercase;\n"
                                      "text-decoration: none;\n"
                                      "background: #A9742A;\n"
                                      "padding: 0px;\n"
                                      "border: 4px solid #494949 !important;\n"
                                      "display: inline-block;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "QPushButton:hover {\n"
                                      "color: #ffffff !important;\n"
                                      "background: #f6b93b;\n"
                                      "border-color: #f6b93b !important;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}")
        self.img_pushButton.setObjectName("img_pushButton")


        self.id3 = QtWidgets.QLabel(self.centralwidget)
        self.id3.setGeometry(QtCore.QRect(40, 250, 30, 30))
        self.id3.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.id3.setObjectName("id3")
        self.img_pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.img_pushButton_2.setGeometry(QtCore.QRect(70, 250, 380, 41))
        self.img_pushButton_2.setStyleSheet("QPushButton {\n"
                                        "color: #FFFFFF !important;\n"
                                        "font: 12pt \'Century Gothic\';                                   \n"
                                        "text-transform: uppercase;\n"
                                        "text-decoration: none;\n"
                                        "background: #A9742A;\n"
                                        "padding: 0px;\n"
                                        "border: 4px solid #494949 !important;\n"
                                        "display: inline-block;\n"
                                        "transition: all 0.4s ease 0s;\n"
                                        "}\n"
                                        "QPushButton:hover {\n"
                                        "color: #ffffff !important;\n"
                                        "background: #f6b93b;\n"
                                        "border-color: #f6b93b !important;\n"
                                        "transition: all 0.4s ease 0s;\n"
                                        "}")
        self.img_pushButton_2.setObjectName("img_pushButton_2")

        self.instruction2 = QtWidgets.QLabel(self.centralwidget)
        self.instruction2.setGeometry(QtCore.QRect(40, 320, 1000, 50))
        self.instruction2.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.instruction2.setObjectName("instruction2")


        self.id4 = QtWidgets.QLabel(self.centralwidget)
        self.id4.setGeometry(QtCore.QRect(40, 390, 30, 30))
        self.id4.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.id4.setObjectName("id4")
        self.label = QtWidgets.QTextEdit(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(70, 390, 330, 50))
        self.label.setStyleSheet("QTextEdit {\n"
                                 "color: #FFFFFF !important;\n"
                                 "font: 7pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "background: #494949;\n"
                                 "placeholder: 'asdasda';\n"
                                 "padding: 0px;\n"
                                 "border: 4px solid #494949 !important;\n"
                                 "display: inline-block;\n"
                                 "}")
        self.label.setObjectName("label")

        self.toolButton = QtWidgets.QPushButton(self.centralwidget)
        self.toolButton.setGeometry(QtCore.QRect(400, 390, 50, 50))
        self.toolButton.setStyleSheet("QPushButton {\n"
                                      "color: #FFFFFF !important;\n"
                                      "font: 12pt \'Century Gothic\';\n"
                                      "text-transform: uppercase;\n"
                                      "text-decoration: none;\n"
                                      "background: #A9742A;\n"
                                      "padding: 0px;\n"
                                      "border: 4px solid #494949 !important;\n"
                                      "display: inline-block;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "QPushButton:hover {\n"
                                      "color: #ffffff !important;\n"
                                      "background: #f6b93b;\n"
                                      "border-color: #f6b93b !important;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}")
        self.toolButton.setObjectName("toolButton")

        self.id5 = QtWidgets.QLabel(self.centralwidget)
        self.id5.setGeometry(QtCore.QRect(40, 460, 30, 30))
        self.id5.setStyleSheet("QLabel {\n"
                                 "color: #AAAAAA !important;\n"
                                 "font: 17pt \'Century Gothic\';\n"
                                 "overflow-wrap: break-word;\n"
                                 "text-transform: uppercase;\n"
                                 "text-decoration: none;\n"
                                 "padding: 0px;\n"
                                 "}")
        self.id5.setObjectName("id5")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(70, 460, 380, 41))
        self.pushButton.setStyleSheet("QPushButton {\n"
                                      "color: #FFFFFF !important;\n"
                                      "font: 12pt \'Century Gothic\';                                   \n"
                                      "text-transform: uppercase;\n"
                                      "text-decoration: none;\n"
                                      "background: #A9742A;\n"
                                      "padding: 0px;\n"
                                      "border: 4px solid #494949 !important;\n"
                                      "display: inline-block;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}\n"
                                      "QPushButton:hover {\n"
                                      "color: #ffffff !important;\n"
                                      "background: #f6b93b;\n"
                                      "border-color: #f6b93b !important;\n"
                                      "transition: all 0.4s ease 0s;\n"
                                      "}")
        self.pushButton.setObjectName("pushButton")

        self.id6 = QtWidgets.QLabel(self.centralwidget)
        self.id6.setGeometry(QtCore.QRect(40, 530, 30, 30))
        self.id6.setStyleSheet("QLabel {\n"
                               "color: #AAAAAA !important;\n"
                               "font: 17pt \'Century Gothic\';\n"
                               "overflow-wrap: break-word;\n"
                               "text-transform: uppercase;\n"
                               "text-decoration: none;\n"
                               "padding: 0px;\n"
                               "}")
        self.id6.setObjectName("id6")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(70, 530, 380, 41))
        self.pushButton_2.setStyleSheet("QPushButton {\n"
                                        "color: #FFFFFF !important;\n"
                                        "font: 12pt \'Century Gothic\';                                   \n"
                                        "text-transform: uppercase;\n"
                                        "text-decoration: none;\n"
                                        "background: #A9742A;\n"
                                        "padding: 0px;\n"
                                        "border: 4px solid #494949 !important;\n"
                                        "display: inline-block;\n"
                                        "transition: all 0.4s ease 0s;\n"
                                        "}\n"
                                        "QPushButton:hover {\n"
                                        "color: #ffffff !important;\n"
                                        "background: #f6b93b;\n"
                                        "border-color: #f6b93b !important;\n"
                                        "transition: all 0.4s ease 0s;\n"
                                        "}")
        self.pushButton_2.setObjectName("pushButton_2")

        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 521, 21))
        self.menubar.setObjectName("menubar")
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def getFolderPSD(self):
        self.dir_path = QFileDialog.getExistingDirectory()
        self.label.setText(self.dir_path)
    def getFolderImages(self):
        self.dir_path = QFileDialog.getExistingDirectory()
        self.img_label.setText(self.dir_path)


    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SI Creation"))
        self.id1.setText(_translate("MainWindow", "1. "))
        self.id2.setText(_translate("MainWindow", "2. "))
        self.id3.setText(_translate("MainWindow", "3. "))
        self.id4.setText(_translate("MainWindow", "4. "))
        self.id5.setText(_translate("MainWindow", "5. "))
        self.id6.setText(_translate("MainWindow", "6. "))

        self.instruction1.setText(_translate("MainWindow", "Rename images for script. "))
        self.instruction2.setText(_translate("MainWindow", "The images and copy are ready. Prepare excel templates for script."))

        self.label.setPlaceholderText(QCoreApplication.translate("Dialog", u"Choose folder with PSD files", None))
        self.img_label.setPlaceholderText(QCoreApplication.translate("Dialog", u"Choose folder with images", None))
        self.toolButton.setText(_translate("MainWindow", "..."))
        self.img_toolButton.setText(_translate("MainWindow", "..."))
        self.pushButton.setText(_translate("MainWindow", "Generate Excel templates"))
        self.pushButton_2.setText(_translate("MainWindow", "Upload Excel templates"))
        self.img_pushButton.setText(_translate("MainWindow", "Change image names"))
        self.img_pushButton_2.setText(_translate("MainWindow", "Generate Excel with images"))

################################################################################################################################
# Feed processing
################################################################################################################################
def generateExcel_psd_templates():
    import_folder_psd = ui.label.toPlainText()
    if import_folder_psd == "":
        pass
    else:
        # create folder export
        export_folder = os.path.join(desktop, "!BulkSI\\Export SI")
        if os.path.exists(export_folder) == False:
            os.mkdir(export_folder)
        export_folder_excel = os.path.join(export_folder, "_Excel Files_")
        if os.path.exists(export_folder_excel) == False:
            os.mkdir(export_folder_excel)
        export_folder_exportfiles = os.path.join(export_folder, "_Export Files_")
        if os.path.exists(export_folder_exportfiles) == False:
            os.mkdir(export_folder_exportfiles)

        # main iteration
        for psd in os.listdir(import_folder_psd):
            if psd.endswith('.psd'):
                full_path = os.path.join(import_folder_psd, psd)
                index_of_dot = psd.index('.')
                file_name_without_extension = psd[:index_of_dot]
                task_type_list = ['text-size', 'text-color', 'text-content', 'background-color', 'image-url', 'image-url-delete-white-bg', 'icon-url', 'image-desktop', 'image-desktop-delete-white-bg', 'icon-desktop']


                # work on PSD
                with Session(full_path, action="open", auto_close=True) as ps:
                    doc = ps.active_document
                    document = PSDImage.open(full_path)
                    excel_path = os.path.join(export_folder_excel, file_name_without_extension + ".xlsx")
                    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
                    layers_in_artboards = []
                    layers_without_artboards = []


                    for i in document:
                        try:
                            if i.kind == "artboard":
                                artboard_layers = []
                                for x in i:
                                    artboard_layer = {"id": x.layer_id, "kind": x.kind, "name": x.name, "task-type": ""}
                                    artboard_layers.insert(0, artboard_layer)
                                artboard = {"kind": i.kind, "id": i.layer_id, "name": i.name, "artboard_layers": artboard_layers}
                                layers_in_artboards.append(artboard)
                            else:
                                layer = {"id": i.layer_id, "kind": i.kind, "name": i.name, "task-type": ""}
                                layers_without_artboards.insert(0, layer)
                        except:
                            print("Error with artboard: " + i.name)


                    if layers_in_artboards:
                        for i in layers_in_artboards:
                            try:
                                df = pd.DataFrame(i["artboard_layers"]).transpose()
                                df.to_excel(writer, sheet_name=i["name"])
                                worksheet = writer.sheets[i["name"]]
                                worksheet.data_validation('B5:AAA5', {'validate': 'list',
                                                                  'source': task_type_list})
                            except:
                                print("Error 1", i)


                    try:
                        if layers_without_artboards:
                            df = pd.DataFrame()
                            for i in layers_without_artboards:
                                df = df.append(i,ignore_index=True)
                            df = df.transpose()
                            df.to_excel(writer)
                            worksheet = writer.sheets["Sheet1"]
                            worksheet.data_validation('B4:AAA4', {'validate': 'list',
                                                             'source': task_type_list})
                    except:
                        print("Error 2")

                    writer.save()
def uploadExcel_psd_templates():
    import_folder_psd = ui.label.toPlainText()
    if import_folder_psd == "":
        pass
    else:
        for xlsx in os.listdir(export_folder_excel):
            try:
                if xlsx.endswith('.xlsx'):
                    index_of_dot = xlsx.index('.')
                    file_name_without_extension = xlsx[:index_of_dot]
                    excel_file = os.path.join(export_folder_excel, xlsx)
                    excel_file_object = pd.ExcelFile(excel_file)
                    psd_path = os.path.join(import_folder_psd, file_name_without_extension + ".psd")


                    with Session(psd_path, action="open", auto_close=False) as ps:
                        doc = ps.active_document
                        document = PSDImage.open(psd_path)
                        app = ph.Application()


                        artboard_id = 0
                        for sheet_name in excel_file_object.sheet_names:
                            artboard = document[artboard_id]
                            parsed_sheet = excel_file_object.parse(sheet_name=sheet_name)
                            df_sheet = pd.DataFrame(parsed_sheet)
                            df_sheet_columns = list(df_sheet.columns)[1:]
                            df_sheet_columns_len = len(df_sheet_columns)
                            df_sheet_rows = list(range(4, len(df_sheet.index)))


                            if df_sheet_rows:
                                for row in df_sheet_rows:

                                    # FOLDERS
                                    product_ident = str(df_sheet.iat[row, 0])
                                    product_folder = os.path.join(export_folder_exportfiles, product_ident)
                                    img_folder = os.path.join(product_folder, "Icons and images")
                                    if os.path.exists(product_folder) == False:
                                        os.mkdir(product_folder)
                                    if os.path.exists(img_folder) == False:
                                        os.mkdir(img_folder)
                                    # CHANGES IN LAYERS
                                    print(" ======================== START PRODUCT: " + str(product_ident) + " /// Artboard: " + str(sheet_name))
                                    df_sheet_empty_cell = 0
                                    for column in df_sheet_columns:
                                        layer_ident = str(df_sheet.iat[0, column+1])
                                        layer_name = str(df_sheet.iat[2, column + 1])
                                        if str(df_sheet.iat[row, column+1]) == "nan":
                                            df_sheet_empty_cell += 1
                                            pass
                                        else:
                                            print("Corrections start", "Artboard: " + sheet_name, "Layer: " + layer_ident + " " + str(layer_name))
                                            change_active_layer = r"""
                                            sTT = stringIDToTypeID;
                                            ref = new ActionReference();
                                            ref.putIdentifier(sTT('layer'), "layer_id");
                                            dsc = new ActionDescriptor();
                                            dsc.putReference(sTT('null'), ref);
                                            executeAction(sTT('select'), dsc);
                                            """.replace("layer_id", str(layer_ident))
                                            app.doJavaScript(change_active_layer)
                                            Layer = doc.activeLayer
                                            # print("Active layer: " + str(doc.activeLayer.id))


                                            image_to_change_JPG = os.path.join(img_folder, str(product_ident) + "_" + str(column) + ".jpg")
                                            image_to_change_PNG = os.path.join(img_folder, str(product_ident) + "_" + str(column) + ".png")
                                            image_to_change_AI = os.path.join(img_folder, str(product_ident) + "_" + str(column) + ".ai")


        ################################################ TextLayer ################################################
                                            try:
                                                if Layer.kind == ps.LayerKind.TextLayer:
                                                    print("Layer type: " + str(Layer.kind))
                                                    if df_sheet.iat[3, column+1] == "text-content":
                                                        Text_Content = df_sheet.iat[row, column + 1]
                                                        ChangeLayer_Text_Content(ps, app, Text_Content)
                                                    elif df_sheet.iat[3, column+1] == "text-size":
                                                        Layer.textItem.size = df_sheet.iat[row, column + 1]
                                                    elif df_sheet.iat[3, column+1] == "text-color":
                                                        text_color = df_sheet.iat[row, column+1]
                                                        ChangeLayer_Text_Color(ps,Layer,text_color)

            ################################################ SolidFillLayer ################################################
                                                if Layer.kind == ps.LayerKind.SolidFillLayer:
                                                    print("Layer type: " + str(Layer.kind))
                                                    if df_sheet.iat[3, column+1] == "background-color":
                                                        solid_color = df_sheet.iat[row, column + 1]
                                                        ChangeLayer_Solid_Color(ps,Layer,solid_color)

            ################################################ SmartObjectLayer ################################################
                                                if Layer.kind == ps.LayerKind.SmartObjectLayer:
                                                    print("Layer type: " + str(Layer.kind))
                                                    if df_sheet.iat[3, column+1] == "image-url":
                                                        image_url = df_sheet.iat[row, column+1]
                                                        image_save_from_url_to_JPG(image_url, image_to_change_JPG)
                                                        ChangeLayer_Image(ps, app, image_to_change_JPG)

                                                    elif df_sheet.iat[3, column + 1] == "image-url-delete-white-bg":
                                                        image_url = df_sheet.iat[row, column + 1]
                                                        convert_to_PNG_from_url(image_url, image_to_change_PNG)
                                                        ChangeLayer_Image_delete_white_bg(ps, app, image_to_change_PNG)

                                                    elif df_sheet.iat[3, column+1] == "icon-url":
                                                        image_url = df_sheet.iat[row, column+1]
                                                        icon_save_from_url(image_url, image_to_change_PNG)
                                                        extension = image_to_change_PNG[-3:]
                                                        if extension == "jpg":
                                                            convert_to_PNG_from_url(image_url, image_to_change_JPG, image_to_change_PNG)
                                                        elif extension == "png":
                                                            pass
                                                        ChangeLayer_Image(ps, app, image_to_change_PNG)


                                                    elif df_sheet.iat[3, column+1] == "image-desktop":
                                                        image_url = df_sheet.iat[row, column+1]
                                                        image_save_from_desktop_to_JPG(image_url, image_to_change_JPG)
                                                        ChangeLayer_Image(ps, app, image_to_change_JPG)

                                                    elif df_sheet.iat[3, column + 1] == "image-desktop-delete-white-bg":
                                                        image_url = df_sheet.iat[row, column + 1]
                                                        convert_to_PNG_from_desktop(image_url, image_to_change_PNG)
                                                        ChangeLayer_Image_delete_white_bg(ps, app, image_to_change_PNG)

                                                    elif df_sheet.iat[3, column+1] == "icon-desktop":
                                                        image_url = df_sheet.iat[row, column+1]
                                                        extension = image_url[-3:]
                                                        if extension == "jpg":
                                                            convert_to_PNG_from_desktop(image_url, image_to_change_PNG)
                                                            ChangeLayer_Image(ps, app, image_to_change_PNG)
                                                        elif extension == "png":
                                                            icon_save_from_desktop(image_url, image_to_change_PNG)
                                                            ChangeLayer_Image(ps, app, image_to_change_PNG)
                                                print("Corrections done", "Artboard: " + sheet_name, "Layer: " + layer_ident)
                                            except Exception:
                                                traceback.print_exc()
                                                error = str(product_ident) + " /// Artboard: " + str(sheet_name) + " /// LayerID: " + str(layer_ident) + " /// LayerName: " + str(layer_name)
                                                errors_list.append(error)


                                    # SAVE
                                    if df_sheet_empty_cell == df_sheet_columns_len:
                                        pass
                                    else:
                                        if len(excel_file_object.sheet_names) == 1:
                                            print("Start saving")
                                            saved_JPG = os.path.join(product_folder, file_name_without_extension + "_" + product_ident + ".jpg")
                                            options_JPG = ps.JPEGSaveOptions()
                                            options_JPG.quality = 10
                                            doc.saveAs(saved_JPG, options_JPG, True)

                                            saved_PSD = os.path.join(product_folder, file_name_without_extension + "_" + product_ident + ".psd")
                                            options_PSD = ps.PhotoshopSaveOptions()
                                            doc.saveAs(saved_PSD, options_PSD, True)
                                        else:
                                            print("Start saving")
                                            for layerrr in doc.layers:
                                                if layerrr.id == artboard.layer_id:
                                                    pass
                                                else:
                                                    layerrr.Delete()

                                            ps.active_document.trim(ps.TrimType.TransparentPixels, True, True, True, True)
                                            ps.active_document.resizeImage(2000,2000)
                                            saved_JPG = os.path.join(product_folder, file_name_without_extension + "_" + product_ident + "_" + str(artboard.name) + ".jpg")
                                            options_JPG = ps.JPEGSaveOptions()
                                            options_JPG.quality = 5
                                            doc.saveAs(saved_JPG, options_JPG, True)
                                            saved_PSD = os.path.join(product_folder, file_name_without_extension + "_" + product_ident + "_" + str(artboard.name) + ".psd")
                                            options_PSD = ps.PhotoshopSaveOptions()
                                            doc.saveAs(saved_PSD, options_PSD, True)


                                            # JS back to first history state
                                            app = ph.Application()
                                            jsx = r"""
                                            var doc = app.activeDocument;
                                            doc.activeHistoryState = doc.historyStates[0];
                                            """
                                            app.doJavaScript(jsx)


                            else:
                                print("Not")

                            if len(excel_file_object.sheet_names) == 1:
                                pass
                            else:
                                artboard_id += 1
                    print("ErrorsList: ", errors_list)
            except Exception:
                traceback.print_exc()
def Change_Image_Names():
    import_folder_images = ui.img_label.toPlainText()
    if import_folder_images == "":
        pass
    else:
        for root, subdirectories, files in os.walk(import_folder_images):
            for subdirectory in subdirectories:
                subdirectory_path = os.path.join(root, subdirectory)
                subdirectory_dir = os.listdir(subdirectory_path)
                image_id = 1
                for image in subdirectory_dir:
                    if image.endswith('.jpg'):
                        image_path = os.path.join(subdirectory_path, image)
                        new_image_name = str(subdirectory) + "_id" + str(image_id) + "_SI " + "_IMG 1.jpg"
                        new_image_path = os.path.join(subdirectory_path, new_image_name)
                        os.rename(image_path, new_image_path)
                        image_id += 1
                    elif image.endswith('.png'):
                        image_path = os.path.join(subdirectory_path, image)
                        new_image_name = str(subdirectory) + "_id" + str(image_id) + "_SI " + "_IMG 1.png"
                        new_image_path = os.path.join(subdirectory_path, new_image_name)
                        os.rename(image_path, new_image_path)
                        image_id += 1
def generateExcel_with_image_urls():
    import_folder_images = ui.img_label.toPlainText()

    data = {
        'Product Name': [],
        'Image Name': [],
        'Image Path': [],
        'SI number': [],
        'Layer number': []
    }

    if import_folder_images == "":
        pass
    else:
        for root, subdirectories, files in os.walk(import_folder_images):
            for subdirectory in subdirectories:
                subdirectory_path = os.path.join(root, subdirectory)
                subdirectory_dir = os.listdir(subdirectory_path)
                image_id = 1
                for image in subdirectory_dir:
                    if image.endswith('.jpg') or image.endswith('.png'):
                        product_name = subdirectory
                        image_name = image
                        image_path = os.path.join(subdirectory_path, image)
                        image_SI = image[image.index("_SI ") + len("_SI "):image.index("_IMG")]
                        image_Layer = image[image.index("_IMG ") + len("_IMG "):].split(".", 1)[0]

                        data['Product Name'].append(product_name)
                        data['Image Name'].append(image_name)
                        data['Image Path'].append(image_path)
                        data['SI number'].append(image_SI)
                        data['Layer number'].append(image_Layer)

    df = pd.DataFrame(data, columns=['Product Name', 'Image Name', 'Image Path', 'SI number', 'Layer number'])
    excel_path = os.path.join(import_folder_images, "Image path.xlsx")
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    df.to_excel(writer)
    worksheet = writer.sheets["Sheet1"]
    worksheet.autofilter(0, 0, df.shape[0], df.shape[1])
    auto_adjust_xlsx_column_width(df, writer, sheet_name="Sheet1", margin=0)



    writer.save()

################################################################################################################################
# Photoshop file processing
################################################################################################################################
def ChangeLayer_Solid_Color(ps,Layer,solid_color):
    app = ph.Application()
    solid_color_red = tuple(int(solid_color[i:i + 2], 16) for i in (0, 2, 4))[0]
    solid_color_green = tuple(int(solid_color[i:i + 2], 16) for i in (0, 2, 4))[1]
    solid_color_blue = tuple(int(solid_color[i:i + 2], 16) for i in (0, 2, 4))[2]
    ps.active_document.activeLayer = Layer
    solid_color = r"""
        var idsetd = charIDToTypeID( "setd" );
            var desc5 = new ActionDescriptor();
            var idnull = charIDToTypeID( "null" );
                var ref1 = new ActionReference();
                var idcontentLayer = stringIDToTypeID( "contentLayer" );
                var idOrdn = charIDToTypeID( "Ordn" );
                var idTrgt = charIDToTypeID( "Trgt" );
                ref1.putEnumerated( idcontentLayer, idOrdn, idTrgt );
            desc5.putReference( idnull, ref1 );
            var idT = charIDToTypeID( "T   " );
                var desc6 = new ActionDescriptor();
                var idClr = charIDToTypeID( "Clr " );
                    var desc7 = new ActionDescriptor();
                    var idRd = charIDToTypeID( "Rd  " );
                    desc7.putDouble( idRd, {red} );
                    var idGrn = charIDToTypeID( "Grn " );
                    desc7.putDouble( idGrn, {green} );
                    var idBl = charIDToTypeID( "Bl  " );
                    desc7.putDouble( idBl, {blue} );
                var idRGBC = charIDToTypeID( "RGBC" );
                desc6.putObject( idClr, idRGBC, desc7 );
            var idsolidColorLayer = stringIDToTypeID( "solidColorLayer" );
            desc5.putObject( idT, idsolidColorLayer, desc6 );
        executeAction( idsetd, desc5, DialogModes.NO );                                                
            """.format(red=solid_color_red, green=solid_color_green, blue=solid_color_blue)
    app.doJavaScript(solid_color)
def ChangeLayer_Text_Color(ps,app, text_color):
    text_color_red = tuple(int(text_color[i:i + 2], 16) for i in (0, 2, 4))[0]
    text_color_green = tuple(int(text_color[i:i + 2], 16) for i in (0, 2, 4))[1]
    text_color_blue = tuple(int(text_color[i:i + 2], 16) for i in (0, 2, 4))[2]
    Text_Color = ps.SolidColor()
    Text_Color.rgb.red = text_color_red
    Text_Color.rgb.green = text_color_green
    Text_Color.rgb.blue = text_color_blue

    text_color = r"""
            templateDocument = app.activeDocument;
            templateName = templateDocument.name.replace('.psd', '');
            executeAction(stringIDToTypeID("placedLayerEditContents"));
    """
    app.doJavaScript(text_color)
    contentLayer = app.activeDocument.activeLayer
    contentLayer.textItem.color = Text_Color
def ChangeLayer_Text_Content(ps, app, Text_Content):
    Text_Content = Text_Content.replace("\n", "\\r ")
    new_contents_len = str(len(Text_Content))

    # contentLayer.textItem.useAutoLeading = false;
    #
    # var contentLayer_leading =  desc.getList(stringIDToTypeID('textStyleRange')).getObjectValue(0).getObjectValue(stringIDToTypeID('textStyle')).getDouble (stringIDToTypeID('leading'));
    # if (desc.hasKey(stringIDToTypeID('transform')))
    # {
    #         var mFactor = desc.getObjectValue(stringIDToTypeID('transform')).getUnitDoubleValue (stringIDToTypeID("yy") );
    #         contentLayer_leading = (contentLayer_leading* mFactor).toFixed(0);
    # }


    new_text = r"""
    

        var ref = new ActionReference();
        ref.putEnumerated( charIDToTypeID("Lyr "), charIDToTypeID("Ordn"), charIDToTypeID("Trgt") );
        var desc = executeActionGet(ref).getObjectValue(stringIDToTypeID('textKey'));
        var contentLayer_textSize =  desc.getList(stringIDToTypeID('textStyleRange')).getObjectValue(0).getObjectValue(stringIDToTypeID('textStyle')).getDouble (stringIDToTypeID('size'));
        if (desc.hasKey(stringIDToTypeID('transform')))
        {
                var mFactor = desc.getObjectValue(stringIDToTypeID('transform')).getUnitDoubleValue (stringIDToTypeID("yy") );
                contentLayer_textSize = (contentLayer_textSize* mFactor).toFixed(0);
        }
        

        
        smartObject = app.activeDocument;

        contentLayer = app.activeDocument.activeLayer;
        contentLayer_contents = contentLayer.textItem.contents
        contentLayer_contents_len = contentLayer_contents.length

        contentLayer.textItem.contents = "Text_Content"

        
        contentLayer.textItem.useAutoLeading = true;
        contentLayer.textItem.hyphenation = false;
        


        """.replace("Text_Content", Text_Content).replace("new_contents_len", new_contents_len)
    app.doJavaScript(new_text)

    # koef = contentLayer_contents_len / new_contents_len
    #
    # if (koef < 1) {
    #     contentLayer.textItem.size = contentLayer_textSize * koef
    #     contentLayer.textItem.baselineShift = -contentLayer_textSize * (1-koef)
    # }
    # else {
    #     contentLayer.textItem.size = contentLayer_textSize * 1
    # }
    #

def image_save_from_desktop_to_JPG(URL, JPG):
    if URL != "":
        img = Image.open(URL)
        img.save(JPG)
def image_save_from_desktop_to_PNG(URL, PNG):
    if URL != "":
        img = Image.open(URL)
        img.save(PNG)

def image_save_from_url_to_JPG(URL, JPG):
    if URL != "":
        p = requests.get(URL)
        out = open(JPG, 'wb')
        out.write(p.content)
        out.close()

        img = Image.open(JPG)
        img.save(JPG)
def image_save_from_url_to_PNG(URL, PNG):
    if URL != "":
        p = requests.get(URL)
        out = open(PNG, 'wb')
        out.write(p.content)
        out.close()

        img = Image.open(PNG)
        img.save(PNG)

def icon_save_from_desktop(URL, PNG):
    if URL != "":
        img = Image.open(URL)
        img.save(PNG)
def icon_save_from_url(URL, PNG):
    if URL != "":
        p = requests.get(URL)
        out = open(PNG, 'wb')
        out.write(p.content)
        out.close()

        img = Image.open(PNG)
        img.save(PNG)

def convert_to_PNG_from_url(URL, JPG, PNG):
    if URL != "":
        p = requests.get(URL)
        out = open(JPG, 'wb')
        out.write(p.content)
        out.close()

        img = Image.open(JPG)
        img = img.convert("RGBA")
        datas = img.getdata()
        newData = []

        for item in datas:
            if item[0] == 255 and item[1] == 255 and item[2] == 255:
                newData.append((255, 255, 255, 0))
            else:
                newData.append(item)

        img.putdata(newData)
        img.save(PNG, "PNG")
def convert_to_PNG_from_desktop(URL, PNG):
    if URL != "":
        img = Image.open(URL)
        img = img.convert("RGBA")
        datas = img.getdata()
        newData = []

        for item in datas:
            if item[0] == 255 and item[1] == 255 and item[2] == 255:
                newData.append((255, 255, 255, 0))
            else:
                newData.append(item)

        img.putdata(newData)
        img.save(PNG, "PNG")

def ChangeLayer_Image_delete_white_bg(ps, app, image_to_change_PNG):
    image_to_change = image_to_change_PNG.replace("\\", "\\\\")
    new_image = r"""
        templateDocument = app.activeDocument;
        templateName = templateDocument.name.replace('.psd', '');
        executeAction(stringIDToTypeID("placedLayerEditContents"));

        smartObject = app.activeDocument;
        smartObject_height = smartObject.height
        smartObject_width = smartObject.width

        var idOpn = charIDToTypeID( "Opn " );
        var desc643 = new ActionDescriptor();
        var iddontRecord = stringIDToTypeID( "dontRecord" );
        desc643.putBoolean( iddontRecord, false );
        var idforceNotify = stringIDToTypeID( "forceNotify" );
        desc643.putBoolean( idforceNotify, true );
        var idnull = charIDToTypeID( "null" );
        desc643.putPath( idnull, new File( "image_to_change" ) );
        var idDocI = charIDToTypeID( "DocI" );
        desc643.putInteger( idDocI, 787 );
        executeAction( idOpn, desc643, DialogModes.NO );
        
        var idtrim = stringIDToTypeID( "trim" );
        var desc613 = new ActionDescriptor();
        var idtrimBasedOn = stringIDToTypeID( "trimBasedOn" );
        var idtrimBasedOn = stringIDToTypeID( "trimBasedOn" );
        var idTrns = charIDToTypeID( "Trns" );
        desc613.putEnumerated( idtrimBasedOn, idtrimBasedOn, idTrns );
        var idTop = charIDToTypeID( "Top " );
        desc613.putBoolean( idTop, true );
        var idBtom = charIDToTypeID( "Btom" );
        desc613.putBoolean( idBtom, true );
        var idLeft = charIDToTypeID( "Left" );
        desc613.putBoolean( idLeft, true );
        var idRght = charIDToTypeID( "Rght" );
        desc613.putBoolean( idRght, true );
        executeAction( idtrim, desc613, DialogModes.NO );

        imageSource = app.activeDocument;
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width
        koef_w = smartObject_width / imageSource.width
        koef_h = smartObject_height / imageSource.height
        min_koef = Math.min(koef_w, koef_h)
        if (smartObject_height >= smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * min_koef,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * koef_w,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        if (smartObject_height < smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * koef_h,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * min_koef,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        imageSource = app.activeDocument;
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width

        
        imageSource.artLayers[0].duplicate(smartObject)
        imageSourceName = imageSource.name.replace(".jpg", '')
        imageSource.close(SaveOptions.DONOTSAVECHANGES)
        smartObject.artLayers[1].remove()


        if (smartObject_height >= imageSource_height) {
        translate = (smartObject_height - imageSource_height) / 2
        app.activeDocument.activeLayer.translate(0, translate)
        }

        if (smartObject_width >= imageSource_width) {
        translate = (smartObject_width - imageSource_width) / 2
        app.activeDocument.activeLayer.translate(translate, 0)
        }

        smartObject.save();
        smartObject.close();


        """.replace("image_to_change", image_to_change)
    app.doJavaScript(new_image)
def ChangeLayer_Image(ps, app, image_to_change_PNG):
    image_to_change = image_to_change_PNG.replace("\\", "\\\\")
    new_image = r"""
        templateDocument = app.activeDocument;
        templateName = templateDocument.name.replace('.psd', '');
        executeAction(stringIDToTypeID("placedLayerEditContents"));

        smartObject = app.activeDocument;
        smartObject_height = smartObject.height
        smartObject_width = smartObject.width

        var idOpn = charIDToTypeID( "Opn " );
        var desc643 = new ActionDescriptor();
        var iddontRecord = stringIDToTypeID( "dontRecord" );
        desc643.putBoolean( iddontRecord, false );
        var idforceNotify = stringIDToTypeID( "forceNotify" );
        desc643.putBoolean( idforceNotify, true );
        var idnull = charIDToTypeID( "null" );
        desc643.putPath( idnull, new File( "image_to_change" ) );
        var idDocI = charIDToTypeID( "DocI" );
        desc643.putInteger( idDocI, 787 );
        executeAction( idOpn, desc643, DialogModes.NO );
        
        imageSource = app.activeDocument;
        imageSource.artLayers[0].duplicate(imageSource)
        imageSource.artLayers[1].remove()
        
        
        var idtrim = stringIDToTypeID( "trim" );
        var desc613 = new ActionDescriptor();
        var idtrimBasedOn = stringIDToTypeID( "trimBasedOn" );
        var idtrimBasedOn = stringIDToTypeID( "trimBasedOn" );
        var idTrns = charIDToTypeID( "Trns" );
        desc613.putEnumerated( idtrimBasedOn, idtrimBasedOn, idTrns );
        var idTop = charIDToTypeID( "Top " );
        desc613.putBoolean( idTop, true );
        var idBtom = charIDToTypeID( "Btom" );
        desc613.putBoolean( idBtom, true );
        var idLeft = charIDToTypeID( "Left" );
        desc613.putBoolean( idLeft, true );
        var idRght = charIDToTypeID( "Rght" );
        desc613.putBoolean( idRght, true );
        executeAction( idtrim, desc613, DialogModes.NO );

        
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width
        koef_w = smartObject_width / imageSource.width
        koef_h = smartObject_height / imageSource.height
        min_koef = Math.min(koef_w, koef_h)
        if (smartObject_height >= smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * min_koef,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * koef_w,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        if (smartObject_height < smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * koef_h,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * min_koef,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        imageSource = app.activeDocument;
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width


        imageSource.artLayers[0].duplicate(smartObject)
        imageSourceName = imageSource.name.replace(".jpg", '')
        imageSource.close(SaveOptions.DONOTSAVECHANGES)
        smartObject.artLayers[1].remove()


        if (smartObject_height >= imageSource_height) {
        translate = (smartObject_height - imageSource_height) / 2
        app.activeDocument.activeLayer.translate(0, translate)
        }

        if (smartObject_width >= imageSource_width) {
        translate = (smartObject_width - imageSource_width) / 2
        app.activeDocument.activeLayer.translate(translate, 0)
        }

        smartObject.save();
        smartObject.close();


        """.replace("image_to_change", image_to_change)
    app.doJavaScript(new_image)
def ChangeLayer_AI(ps, app, image_to_change_PNG):
    image_to_change = image_to_change_PNG.replace("\\", "\\\\")
    new_image = r"""
        templateDocument = app.activeDocument;
        templateName = templateDocument.name.replace('.psd', '');
        executeAction(stringIDToTypeID("placedLayerEditContents"));

        smartObject = app.activeDocument;
        smartObject_height = smartObject.height
        smartObject_width = smartObject.width

        
        // =======================================================
        var idOpn = charIDToTypeID( "Opn " );
            var desc72 = new ActionDescriptor();
            var iddontRecord = stringIDToTypeID( "dontRecord" );
            desc72.putBoolean( iddontRecord, false );
            var idforceNotify = stringIDToTypeID( "forceNotify" );
            desc72.putBoolean( idforceNotify, true );
            var idnull = charIDToTypeID( "null" );
            desc72.putPath( idnull, new File( "D:\\Photo and Design\\AIMSURE\\Logo\\aimsure (без текста).ai" ) );
            var idsmartObject = stringIDToTypeID( "smartObject" );
            desc72.putBoolean( idsmartObject, true );
            var idDocI = charIDToTypeID( "DocI" );
            desc72.putInteger( idDocI, 254 );
        executeAction( idOpn, desc72, DialogModes.NO );
        


        imageSource = app.activeDocument;
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width
        koef_w = smartObject_width / imageSource.width
        koef_h = smartObject_height / imageSource.height
        min_koef = Math.min(koef_w, koef_h)
        if (smartObject_height >= smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * min_koef,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * koef_w,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        if (smartObject_height < smartObject_width) {
            if (imageSource.height >= imageSource.width) {
                imageSource.resizeImage(null,UnitValue(imageSource.height * koef_h,"px"),null,ResampleMethod.BICUBIC);
            }
            else {
                imageSource.resizeImage(UnitValue(imageSource.width * min_koef,"px"),null,null,ResampleMethod.BICUBIC);
            };
        }
        imageSource = app.activeDocument;
        imageSource_height = imageSource.height
        imageSource_width = imageSource.width


        
        imageSource.artLayers[0].duplicate(smartObject)
        imageSource.close(SaveOptions.DONOTSAVECHANGES)
        smartObject.artLayers[1].remove()
        
        
        if (smartObject_height >= imageSource_height) {
        translate = (smartObject_height - imageSource_height) / 2
        app.activeDocument.activeLayer.translate(0, translate)
        }

        if (smartObject_width >= imageSource_width) {
        translate = (smartObject_width - imageSource_width) / 2
        app.activeDocument.activeLayer.translate(translate, 0)
        }


        smartObject.save();
        smartObject.close();


        """.replace("image_to_change", image_to_change)
    app.doJavaScript(new_image)

def Remove_Background():
    app = ph.Application()
    remove_background = r"""
    var idremoveBackground = stringIDToTypeID("removeBackground");
    executeAction(idremoveBackground, undefined, DialogModes.NO);
    """
    app.doJavaScript(remove_background)


################################################################################################################################
# UI
################################################################################################################################
app = QtWidgets.QApplication(sys.argv)

# init
Dialog = QtWidgets.QDialog()
ui = Ui_MainWindow()
ui.setupUi(Dialog)
Dialog.show()

Get_Folder_PSD = ui.toolButton.clicked.connect(ui.getFolderPSD)
Get_Folder_Images = ui.img_toolButton.clicked.connect(ui.getFolderImages)
generateExcel_psd_templates = ui.pushButton.clicked.connect(generateExcel_psd_templates)
uploadExcel_psd_templates = ui.pushButton_2.clicked.connect(uploadExcel_psd_templates)
Change_Image_Names = ui.img_pushButton.clicked.connect(Change_Image_Names)
generateExcel_with_image_urls = ui.img_pushButton_2.clicked.connect(generateExcel_with_image_urls)

# Main loop
sys.exit(app.exec_())




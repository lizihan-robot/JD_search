import sys,os
from PySide2.QtCore import Qt, Slot
from PySide2.QtGui import QPainter, QImage, QPixmap
from PySide2.QtWidgets import (QAction, QApplication, QHeaderView, QHBoxLayout, QLabel, QLineEdit,
                               QMainWindow, QPushButton, QTableWidget, QTableWidgetItem,
                               QVBoxLayout, QWidget, QMessageBox)
from PySide2.QtWidgets import QFileDialog
from PySide2.QtCharts import QtCharts
from PySide2.QtUiTools import QUiLoader
from update_JD  import check_update, change_format
from load_docx import load_file
import win32com.client as wc
from docx import Document
import pythoncom

import win32api
import win32con
from win32comext.shell.shell import ShellExecuteEx
from win32comext.shell import shellcon
import win32process
import time
import win32event
import _thread


class MainWindow(QMainWindow):
    def __init__(self):
        """
        """
        QMainWindow.__init__(self)
        self.JD_dir_path = "C:/Users/李子汉/Desktop/intv"
        self.ui = QUiLoader().load('ui_of.ui')
        self.ui.button_import_folder.clicked.connect(self.handle_load_folder)#导入文件夹-按钮
        self.ui.button_check_update.clicked.connect(self.handle_check_update)#检查更新-按钮
        self.ui.button_load_docx.clicked.connect(self.handle_load_docx)#转化docx文件-按钮
        self.ui.listWidget.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)#以滚动窗口显示-按钮
        self.ui.button_search.clicked.connect(self.handle_key_world_search)#关键词搜索-按钮
        self.ui.listWidget.currentItemChanged.connect(self.chose_JD_file)#主窗口listWideget
        self.ui.button_clean_up.clicked.connect(self.clean_up)#清屏-按钮

    def clean_up(self):
        "清屏"
        self.ui.listWidget.clear()

    def show_JD_format(self,file_list):
        """将搜索结果打印在窗口上"""
        search = "共搜索了"+str(file_list[0])+"个文件"
        print("search",search)
        self.ui.listWidget.addItem(search)
        del file_list[0]
        for JD_file in file_list:
            self.ui.listWidget.addItem(JD_file)

    def open_JD_file(self,JD_file):
        print("JD_file",JD_file)
        if os.path.splitext(JD_file)[1] ==".doc" or os.path.splitext(JD_file)[1] ==".docx":
            pythoncom.CoInitialize()
            word = wc.Dispatch("Word.Application")
            word.Visible = 1
            # doc.DisplayAlerts = 1
            doc = word.Documents.Open(JD_file)
        if os.path.splitext(JD_file)[1] ==".pdf":
            # handle = win32api.ShellExecuteEx(0, 'open', JD_file, '', '', 1) 
            procInfo = ShellExecuteEx(nShow=win32con.SW_SHOWNORMAL,
                                  fMask=shellcon.SEE_MASK_NOCLOSEPROCESS,
                                  lpVerb='open',
                                  lpFile=JD_file,
                                  lpParameters='')
            # handle = win32process.CreateProcess(exePath, helplFile, None, None, 0, win32process.CREATE_NO_WINDOW, None, None, win32process.STARTUPINFO())                   


    def chose_JD_file(self):
        JD_file=self.ui.listWidget.currentItem().text()
        self.open_JD_file(JD_file)
        print("-------JD_file:",JD_file)

    @Slot()
    def handle_key_world_search(self):
        """点击搜索关键词"""
        key_world1 = self.ui.find_key_world1.toPlainText()
        key_world2 = self.ui.find_key_world2.toPlainText()
        key_world3 = self.ui.find_key_world3.toPlainText()
        key_world4 = self.ui.find_key_world4.toPlainText()
        key_world5 = self.ui.find_key_world5.toPlainText()
        key_world6 = self.ui.find_key_world6.toPlainText()
        key_world7 = self.ui.find_key_world7.toPlainText()
        key_world8 = self.ui.find_key_world8.toPlainText()
        key_world_list = [key_world1, key_world2, key_world3, key_world4, key_world5, key_world6, key_world7, key_world8]
        key_world_list = [i for i in key_world_list if i!='']
        lib_path = self.JD_dir_path+'/JD_lib_docx'
        # load_file = load_file()
        res = load_file().get_list_from_lib(lib_path, key_world_list)
        if len(res) == 1:
            self.ui.listWidget.toPlainText("无法找到匹配的信息")
        else:
            self.show_JD_format(res)
        



    
    @Slot()
    def handle_load_folder(self):
        """        点击导入文件夹        """
        self.JD_dir_path=QFileDialog.getExistingDirectory(self,"choose directory")
        print(self.JD_dir_path)
        self.ui.text_floder_path.setText(self.JD_dir_path)
        return self.JD_dir_path

        

    @Slot()
    def handle_check_update(self):
        """        点击检查库更新情况        """
        # load_folder = self.handle_load_folder()
        result = check_update().check_JD_files(self.JD_dir_path)
        print("result:",result)
        if result["result"]:
            self.ui.listWidget.addItem("文件更新成功!!!")
        else:
            self.ui.listWidget.addItem("文件更新失败!!!{list}没有录入".format(list=result["result_list"]))
        return result

    @Slot()    
    def handle_load_docx(self):
        """        转化导入docx        """
        print("导入docx路径",self.JD_dir_path)
        try:
            result = change_format().change_format_to_docx(self.JD_dir_path)    
            if result:
                self.ui.listWidget.addItem("文件转化导入成功!!!")
            else:
                self.ui.listWidget.addItem("文件转化导入失败!!!")
        except Exception as conversion_error:
            self.ui.listWidget.addItem("文件转化出错!!!{error}".format(error=conversion_error))
        return result
        









if __name__=="__main__":
    app = QApplication(sys.argv)

    mainw = MainWindow()
    mainw.ui.show()
    app.exec_()






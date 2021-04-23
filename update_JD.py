import os,shutil
import sys
import importlib
from io import StringIO
from win32com import client as wc
import warnings
from threading import Thread

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice
from pdfminer.layout import LAParams
from pdfminer.converter import PDFPageAggregator

import warnings
from docx import Document
from win32comext.shell.shell import ShellExecuteEx
from win32comext.shell import shellcon
import win32process
import win32event
import win32con
class check_lib():
    def __init__(self,folder_path):
        self.path = folder_path
        self.path = folder_path
        if not os.path.exists(folder_path+'/JD_lib_docx'):
            print("未发现JD_lib_docx库!!!新建中.....")
            os.makedirs(folder_path+'/JD_lib_docx')
            print("JD_lib_docx库!!新建完成")
        else:
            print("已发现JD_lib_docx库!!!读取中.....")

class check_update():
    """
    检查库更新状况
    """
    def load_lib_file(self,path, file_list=[]):
        """
        加载已经转化后的文件夹
        """
        # 查询库文件
        lib_path = os.path.join(path,"JD_lib_docx")
        dir = os.listdir(lib_path)
        for i in dir:
            file_list.append(i)
        print("load_lib_file",file_list)
        return file_list
    

    def load_JD_folder(self,path,file_list=[]):
        """
        加载未转化的原始文件夹
        """
        dir = os.listdir(path)
        for i in dir:
            if not os.path.isdir(path+"/"+i):
                file_list.append(i)
        print("load_JD_folder",file_list)
        return file_list
    
    def check_JD_files(self,path):
        """
        比较库和文件夹list
        """
        result={"result":False,"result_list":[]}
        check_lib(path)
        turn_load_folder_list=[]
        for i in self.load_JD_folder(path):
            if os.path.splitext(i)[1] ==".doc":
                turn_load_folder_list.append(os.path.splitext(i)[0]+".docx")
            else:
                turn_load_folder_list.append(i)
        print("turn_load_folder_list:",turn_load_folder_list)
        JD_lib_files_list = self.load_lib_file(path)
        result["result"] = set(turn_load_folder_list)<=set(JD_lib_files_list)
        equal_list = [x for x in turn_load_folder_list if x not in JD_lib_files_list]
        if set(turn_load_folder_list)<=set(JD_lib_files_list):
            result["result"] = True
            result["result_list"] = "库更新正确"
        else:
            result["result"] = False
            result["result_list"] = equal_list
        print("返回结果result:",result)
        return result
        
class change_format():
    """    更改文件格式    """
    def change_format_to_docx(self,path):
        check_lib(path)
        file_list = self.find_file(path)
        print("读取文件信息list",file_list)
        # try:
        for file in file_list:
            if os.path.splitext(file)[1] == ".doc":
                self.doc_to_docx(file)
            if os.path.splitext(file)[1] == ".pdf" or os.path.splitext(file)[1] == ".docx":
                try:
                    lib_path = os.path.join(path,"JD_lib_docx")
                    print("lib_path:",lib_path)
                    lib_path_file = os.path.join((os.path.join(path,"JD_lib_docx")),os.path.split(file)[1])
                    print("~~~~~~~~~~~",lib_path_file)
                    shutil.copy(file, lib_path_file)
                except IOError as e:
                    print("Unable to copy file. %s" % e)
                except:
                    print("Unexpected error:", sys.exc_info())

        print("----文件转化完成")
        return True

    def pdf_to_docx(self,path_name):
        """pdf文件转化为docx"""
        # # import pdb; pdb.set_trace()
        path = open(path_name, 'rb')
        save_name= os.path.split(os.path.splitext(path_name)[0])[1]+".docx"
        # print("save_name",save_name)
        save_path = os.path.split(path_name)[0]+"/JD_lib_docx"
        # print("save_path:",save_path)
        save_path_name = os.path.join(save_path,save_name)
        print("save_path_name:",save_path_name)
        parser = PDFParser(path)
        document = PDFDocument(parser)
        if not document.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            rsrcmgr=PDFResourceManager()
            laparams=LAParams()
            device=PDFPageAggregator(rsrcmgr,laparams=laparams)
            interpreter=PDFPageInterpreter(rsrcmgr,device)
            for page in PDFPage.create_pages(document):
                interpreter.process_page(page)
                layout=device.get_result()
                with open('%s'%(save_path_name),"a") as f:
                    for x in layout:
                        result_text = x.get_text()
                        if len(result_text)!=0:
                            f.write(result_text+'\n')
            device.close()                                                                                              

    def remove_control_characters(content):
        mpa = dict.fromkeys(range(32))
        return content.translate(mpa)

    def find_file(self,path, file_list=[]):
        """
        加载所有文件
        """
        dir = os.listdir(path)
        for i in dir:
            if os.path.isdir(i):
                pass
            else:
                file_list.append(os.path.join(path,i))
        # print("file_list:",file_list)
        return file_list
    
    def doc_to_docx(self,path_name):
        """
        doc文件转化为docx
        """
        try:
            # import pdb; pdb.set_trace()
            try:
                word=wc.gencache.EnsureDispatch('kwps.application')
            except:
                word=wc.gencache.EnsureDispatch('wps.application')
            else:
                word=wc.gencache.EnsureDispatch('word.application')

            word = wc.Dispatch('Word.Application')
            (filepath,tempfilename) = os.path.split(path_name)
            doc = word.Documents.Open(path_name)  # 目标路径下的文件
            new_file_path_name = os.path.join(os.path.join(filepath,"JD_lib_docx"),os.path.splitext(tempfilename)[0]+".docx")
            doc.SaveAs(new_file_path_name, FileFormat = 12)  # 转化后路径下的文件
            print("------%s -文件转化成功"%tempfilename)
            doc.Close()
        except:
            print("------%s -文件转化失败"%tempfilename)
        finally:
            try:
                wps.Documents.Close()
                wps.Documents.Close(wc.wdDoNotSaveChanges)
                wps.Close()
            except:
                pass
    

    def to_docx(self,path):
        pass

if __name__=="__main__":
    floder_path = "C:/Users/李子汉/Desktop/intv"
    # a=change_format()
    # a1=a.change_format_to_docx(floder_path)
    # # a1 = a.pdf_to_docx(floder_path)
    # print("-----------转化结果----------",a1)
    a= check_update()
    a.check_JD_files(floder_path)





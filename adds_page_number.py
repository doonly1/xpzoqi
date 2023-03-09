import os
from docx import Document
from xpzoqi import save_docx
from mystyle import my_number_style,set_page

from win32com import client
from win32com.client import constants

BASEDIR = os.path.dirname(__file__)



def doc_to_docx(BASEDIR):
    word = client.Dispatch("Word.Application")
    for file in os.listdir(BASEDIR):
        if file.endswith('.doc') and not file.startswith("~$"):
            print('转化docx：{}'.format(file))
            file = BASEDIR+'/'+file
            doc = word.Documents.Open(file) 
            doc.SaveAs("{}x".format(file), 12)  #另存为后缀为".docx"的文件，其中参数12指docx文件       
            doc.Close()    #关闭原来word文件
            os.remove(file)   #删除原.doc文件

def set_page_number(docx_name):
    word = client.Dispatch("Word.Application") 
    doc = word.Documents.Open(docx_name)
    for wd_section in doc.Sections:   #section内部成员编号是从1开始的
        wd_section.Footers(constants.wdHeaderFooterPrimary).PageNumbers.Add(PageNumberAlignment=2)  #添加页码
        wd_section.Footers(constants.wdHeaderFooterPrimary).PageNumbers.NumberStyle=57
    doc.Save()
    doc.Close()

def adds_page_number(BASEDIR):
    #批量加页码
    print('当前工作目录（adds_num）：',BASEDIR)
    doc_to_docx(BASEDIR)

    digit_files=0
    for file in os.listdir(BASEDIR):        #确定digitfiles数量
        if file.endswith('.docx') and file[:4].isdigit():
            digit_files += 1
    if digit_files == 0:                        #如果没有，就生成
        for file in os.listdir(BASEDIR):
            if file.endswith('.docx') and not file.startswith("~$"):
                doc=Document(file)
                save_docx(doc,file)

    for file in os.listdir(BASEDIR):        #对生成的digitfiles加页码
        if file.endswith('.docx') and file[:4].isdigit():
            print('正在添加页码：',file)
            set_page_number(BASEDIR+'/'+file)   #激活页码样式，启用关闭一次文档
            
            doc=Document(file)
            set_page(doc)
            my_number_style(doc)  #设置页码样式
            doc.save(file)

            set_page_number(BASEDIR+'/'+file)   #激活页码样式，启用关闭又一次
            print('添加成功。\n')



if __name__ == '__main__':
    adds_page_number(BASEDIR)


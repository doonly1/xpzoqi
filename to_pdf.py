import os
from win32com import client
from adds_number import doc_to_docx


def convert_to_pdf(BASEDIR):
    try:
        doc_to_docx(BASEDIR)
    except:
        pass
    for file in os.listdir(BASEDIR):
        if file.endswith('.docx') and not file.startswith("~$"):
            file = BASEDIR+'//'+file
            word = client.Dispatch("Word.Application") 
            doc = word.Documents.Open(file)
            doc.SaveAs(file[:-4]+'pdf', 17)
            doc.Close()
    word.Quit()
    print('全部转化成功')

   
if __name__ == '__main__':
    BASEDIR = os.path.dirname(__file__)
    convert_to_pdf(BASEDIR)

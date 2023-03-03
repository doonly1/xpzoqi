import os
from win32com import client
from adds_page_number import doc_to_docx

def convert_to_pdf():        
    current_dir = os.path.abspath('./')
    os.chdir(current_dir) 
    try:
        doc_to_docx(current_dir)
    except:
        pass
    for file in os.listdir(current_dir):
        if file.endswith('.docx') and not file.startswith("~$"):
            file = current_dir+'//'+file
            word = client.Dispatch("Word.Application") 
            doc = word.Documents.Open(file)
            doc.SaveAs(file[:-4]+'pdf', 17)
            doc.Close()
    word.Quit()
    print('全部转化成功')
    
if __name__ == '__main__':
    convert_to_pdf()

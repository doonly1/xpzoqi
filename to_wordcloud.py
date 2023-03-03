
from docx import *
from adds_page_number import doc_to_docx
import os,jieba
from wordcloud import WordCloud


def show_wordcloud():
    current_dir = os.path.abspath('./')
    os.chdir(current_dir)
    try:
        doc_to_docx(current_dir)
    except:
        pass

    for file in os.listdir(current_dir):
        if file.endswith('.docx') and not file.startswith("~$"):
            print('正在生成词云：',file)
            doc=Document(file)
            paras_texts=''
            paras=doc.paragraphs
            for para in paras:
                paras_texts = paras_texts + para.text
            words = ' '.join(jieba.lcut(paras_texts))
            font = r'C:\Windows\Fonts\simfang.ttf'
            wordcloud = WordCloud(collocations=False, font_path=font, stopwords={"的","和"},\
                                  width=1980, height=1080, margin=2).generate(words)
            wordcloud.to_file(file[:-4]+'png')
            doc.save(file)
            print('词云成功：',file[:-4]+'png')

if __name__ == '__main__':
    show_wordcloud()
    os.system('pause')

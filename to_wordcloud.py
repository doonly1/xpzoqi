from docx import Document
from wordcloud import WordCloud
import os,jieba

from adds_number import doc_to_docx


def show_wordcloud(BASEDIR):
    os.chdir(BASEDIR)
    try:
        doc_to_docx(BASEDIR)
    except:
        pass

    for file in os.listdir(BASEDIR):
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
    BASEDIR = os.path.dirname(__file__)
    show_wordcloud(BASEDIR)

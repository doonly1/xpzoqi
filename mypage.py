from docx import *
from docx.shared import *

def set_page(doc):
    #设置页面
    secs=doc.sections	#赋值节给sec
    for sec in secs:
        sec.left_margin=Cm(2.8)	 #为每一节设置页边距
        sec.right_margin=Cm(2.6)
        sec.top_margin=Cm(3.7)
        sec.bottom_margin=Cm(3.5)
        sec.gutter=Cm(0)			 #装订线
        sec.header_distance=Cm(2)
        sec.footer_distance=Cm(2.5)

        doc.settings.odd_and_even_pages_header_footer = False  #是否奇偶页不同
        hed1=sec.header
        hed2=sec.even_page_header
        fot1=sec.footer
        fot2=sec.even_page_footer
         
        for hed in (hed1,hed2):     #页眉
            hed.is_linked_to_previous = True
            
        for fot in (fot1,fot2):		        #页脚
            fot.is_linked_to_previous = True	#去掉原有页脚
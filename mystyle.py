from docx import *
from docx.shared import *
from docx.enum.style import *
from docx.enum.text import *

import itertools

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

def clear_styles(doc):
    for style in itertools.chain(doc.styles,doc.styles.latent_styles):
        style_attr(style,style.priority)
        if style.name not in [ 'Normal','page number','Title',"Heading 1","Heading 2","Heading 3","Heading 4"]:
            try:
                style.quick_style = False
                style.delete()
            except:
                print('未删除：',style.name) 
    try:
       run_fm(doc.styles['Normal'],'仿宋',16,0,0,0)
       para_f=doc.styles['Normal'].paragraph_format
       para_fm(para_f,0,0,28.95,0,0,32,'J')
    except:
       pass

def add_my_styles(doc):
    try:
        add_style(doc,'unindent',1)
        add_style(doc,'Norm',2)
        add_style(doc,'Tit',3)
        add_style(doc,'H1',4)
        add_style(doc,'H2',5)
        add_style(doc,'H3',6)
        add_style(doc,'H4',21)
        add_style(doc,'Apdix',9)
        add_style(doc,'Apdix 1',10)
        add_style(doc,'Apdix 2',11)
        add_style(doc,'Blackbody',12)
        add_style(doc,'dater',13)
        add_style(doc,'Sign',14)
        add_style(doc,'heder',15)
        add_style(doc,'foter',16)
        add_style(doc,'Regular',17)
    except:
        print('添加样式失败')
    
    try:
        run_fm(doc.styles['Norm'],'仿宋',16,0,0,0)
        para_f=doc.styles['Norm'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,32,'J')

        run_fm(doc.styles['unindent'],'仿宋',16,0,0,0)
        para_f=doc.styles['unindent'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,0,'L')

        run_fm(doc.styles['Tit'],'方正小标宋简体',22,0,0,0)
        para_f=doc.styles['Tit'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,0,'C')

        run_fm(doc.styles['H1'],'黑体',16,0,0,0)
        para_f=doc.styles['H1'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,32,'J')

        run_fm(doc.styles['H2'],'楷体',16,0,0,0)
        para_f=doc.styles['H2'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,32,'J')

        run_fm(doc.styles['H3'],'仿宋',16,0,0,0)
        para_f=doc.styles['H3'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,32,'J')
        
        run_fm(doc.styles['H4'],'仿宋',16,0,0,0)
        para_f=doc.styles['H4'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,32,'J')
        
        run_fm(doc.styles['Apdix'],'仿宋',16,0,0,0)
        para_f=doc.styles['Apdix'].paragraph_format
        para_fm(para_f,0,0,28.95,80,0,-48,'J')

        run_fm(doc.styles['Apdix 1'],'仿宋',16,0,0,0)
        para_f=doc.styles['Apdix 1'].paragraph_format
        para_fm(para_f,0,0,28.95,96,0,-64,'J')

        run_fm(doc.styles['Apdix 2'],'仿宋',16,0,0,0)
        para_f=doc.styles['Apdix 2'].paragraph_format
        para_fm(para_f,0,0,28.95,96,0,-16,'J')

        run_fm(doc.styles['Blackbody'],'黑体',16,0,0,0)
        para_f=doc.styles['Blackbody'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,0,'J')

        run_fm(doc.styles['Regular'],'楷体',16,0,0,0)
        para_f=doc.styles['Regular'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,0,'J')

        run_fm(doc.styles['heder'],'楷体',14,0,0,0)
        para_f=doc.styles['heder'].paragraph_format
        para_fm(para_f,0,0,1,16,16,0,'C')
        
        run_fm(doc.styles['foter'],'宋体',14,0,0,0)
        para_f=doc.styles['foter'].paragraph_format
        para_fm(para_f,0,0,1,16,16,0,'L')
        
        run_fm(doc.styles['dater'],'仿宋',16,0,0,0)
        para_f=doc.styles['dater'].paragraph_format
        para_fm(para_f,0,0,28.95,0,64,0,'R')
        
        run_fm(doc.styles['Sign'],'仿宋',16,0,0,0)
        para_f=doc.styles['Sign'].paragraph_format
        para_fm(para_f,0,0,28.95,0,0,0,'R')
    except:
        pass

def style_attr(style,priority):
    style.hidden = False
    style.unhide_when_used = False
    style.quick_style = True
    style.priority = priority
    style.locked = False

def para_fm(para_name,spc_bef,spc_af,line_spc,left_ind,right_ind,first_l_ind,align):
    try :
        para_f=para_name.paragraph_format
    except :
        para_f=para_name
    para_f.space_before = Pt(spc_bef)   #段前间距
    para_f.space_after = Pt(spc_af)		#段后间距
    if  line_spc>3:
        para_f.line_spacing = Pt(line_spc)  #段中
    elif line_spc<=3:
        para_f.line_spacing = line_spc
    para_f.left_indent = Pt(left_ind)   #左缩进
    para_f.right_indent = Pt(right_ind)	  #右缩进
    para_f.first_line_indent = Pt(first_l_ind)  #首行缩进，负值表示悬挂
    align_dic={'L':'WD_PARAGRAPH_ALIGNMENT.LEFT',\
               'R':'WD_PARAGRAPH_ALIGNMENT.RIGHT',\
               'C':'WD_PARAGRAPH_ALIGNMENT.CENTER',\
               'J':'WD_PARAGRAPH_ALIGNMENT.JUSTIFY'}
    para_f.alignment = eval(align_dic[align])
    para_f.widow_control=False
    para_f.keep_with_next=False
    para_f.page_break_before=False
    para_f.keep_together=False

def run_fm(run,font_type='仿宋',font_size=16,r=0,g=0,b=0,font_name='Times New Roman'):
    font3=run.font
    font3.name = font_name		      #字体类型
    from docx.oxml.ns import qn			   #设置中文字体
    font3.element.rPr.rFonts.set(qn('w:eastAsia'),font_type)
    font3.size = Pt(font_size)		  #字体大小
    font3.color.rgb = RGBColor(r,g,b)
    font3.snap_to_grid=False

def add_style(doc,style,priority,hidden=1,quick=1):
    styles = doc.styles
    style = styles.add_style(style, WD_STYLE_TYPE.PARAGRAPH)
    style.hidden = eval(['True','False'][hidden])
    style.unhide_when_used=False
    style.quick_style = eval(['False','True'][quick])
    style.priority = priority
    style.base_style = styles['Normal']

def my_number_style(doc):
    tok=0    
    try:
        run_fm(doc.styles['page number'],'宋体',14,0,0,0,'SimSun-ExtB')
        style_attr(doc.styles['page number'],20)
    except:
        tok=1
    try:
        run_fm(doc.styles.latent_styles['page number'],'宋体',14,0,0,0,'SimSun-ExtB')
        style_attr(doc.styles.latent_styles['page number'],24)
    except:
        if tok:
            add_style(doc,'page number',20)
            para_f=doc.styles['page number'].paragraph_format
            run_fm(doc.styles['page number'],'宋体',14,0,0,0,'SimSun-ExtB')
            para_fm(para_f,0,0,1,16,16,0,'R')
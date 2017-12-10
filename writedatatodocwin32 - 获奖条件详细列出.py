#-*- coding:utf-8 -*-
#说明：因为docx不能打包，所以改为win32com的方式实现
import os  
import win32com 
from win32com.client import Dispatch, constants 

import time

import readstuandunivinfo as rsui
import dataprocess

#设置Range对象的内容和格式
def SetRangeTextAndFormat(Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    Range.Font.Name = fontname
    Range.Font.Size = fontsize
    Range.Font.Italic = itaflag#是否斜体
    Range.Font.Bold = boldflag#是否粗体
    Range.ParagraphFormat.Alignment = alignmentflag # 012左中右
    Range.ParagraphFormat.LeftIndent = leftindent#设置段落格式左缩进
    Range.ParagraphFormat.FirstLineIndent = firstlfind#首行缩进
    
    Range.InsertBefore(txt) # 插入内容
    
    #Range.Text = txt

#设置段落内容和格式
def SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    '''
    功能：设置段落内容和格式
    '''
    p = document.Paragraphs.Add()
    SetRangeTextAndFormat(p.Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    return p
    '''
    p.Range.Font.Name = fontname
    p.Range.Font.Size = fontsize
    p.Range.Font.Italic = itaflag#是否斜体
    p.Range.Font.Bold = boldflag#是否粗体
    p.Range.ParagraphFormat.Alignment = alignmentflag # 012左中右
    p.Range.ParagraphFormat.LeftIndent = leftindent#设置段落格式左缩进
    p.Range.ParagraphFormat.FirstLineIndent = firstlfind#首行缩进
    #p.Range.ParagraphFormat.LineSpacing = 12 # 行间距
    p.Range.InsertBefore(txt) # 插入内容
    '''
#设置表格的内容和格式
def SetTableTextAndFormat(cell,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag):
    SetRangeTextAndFormat(cell.Range,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        
def WriteFrontCoverPage(document,StuInfoData):#写封面页
    '''
    功能：写报告的封面
    '''
    ##先添加图像（公司logo图片为logo.png）
    curospath = os.path.abspath('.')
    picpath = curospath + r'\pictures\logo.png'
    p = document.Paragraphs.Add()
    p.Range.InlineShapes.AddPicture(picpath, False,True)

    #两行空行
    SetParagraphTextAndFormat(document,'',16,u'黑体',0,0,0,0,0)
    
    #居中写产品名
    txt = u'北京展梦学业规划指导中心“一元咨询”产品'
    SetParagraphTextAndFormat(document,txt,16,u'黑体',1,0,0,0,0)
    
    #居中写报告名
    txt = u'自主招生院校专业定位及申报建议报告书'
    SetParagraphTextAndFormat(document,txt,22,u'黑体',1,0,0,0,0)

    #下面再输入1行空格
    SetParagraphTextAndFormat(document,'',22,u'黑体',1,0,0,0,0)
    
    #下面居中输入报告的内容简介    
    txt = u'本报告分为如下几个部分：'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（1）专业取向分析及建议'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（2）院校定位及建议'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)
    txt = u'（3）附件：自主招生内参资料'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,120,0,0,0)

    #插入两行空行
    SetParagraphTextAndFormat(document,'\n\n',12,u'宋体',0,120,0,0,0)

    #输出学生信息
    
    name = StuInfoData['name'].decode('utf-8')
    txt = u'姓  名：' + name
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    school = StuInfoData['school'].decode('utf-8')
    txt = u'学  校：' + school
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    grade = StuInfoData['grade'].decode('utf-8')
    txt = u'年  级：' + grade
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    curtime = time.strftime('%Y-%m-%d',time.localtime(time.time()))
    txt = u'时  间：' + curtime
    SetParagraphTextAndFormat(document,txt,13,u'黑体',0,120,0,0,0)
    

    #插入3行空行
    SetParagraphTextAndFormat(document,'\n\n\n',12,u'宋体',0,120,0,0,0)
    #输入备注信息：
    txt = u'本报告由北京展梦学业规划指导中心提供，报告内容仅对本次采集的数据负责，如有问题请致电400-88888888' + \
    u'或发送邮件至学业邮箱：xf700@qq.com联系我们！'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,0,25,0,0)

    #插入3行空行
    SetParagraphTextAndFormat(document,'\n\n\n',12,u'宋体',0,120,0,0,0)
        
    #祝福语
    txt = u'北京展梦学业规划指导中心祝莘莘学子心想事成，金榜题名！'
    SetParagraphTextAndFormat(document,txt,12,u'楷体',1,0,0,1,0)

    #插入分页
    p = document.Paragraphs.Add()
    p.Range.InsertBreak()
      
#写报告正文第一页，包括个人信息和专业建议结果
def WriteFirstPage(document,StuInfoData,MBTIData,SubMajData):#写第一页
    '''
    功能：写报告正文第一页，包括个人信息和专业建议结果
    '''
    #写在前面的话：
    if StuInfoData['identity'].find('学生') > -1:
        txt = StuInfoData['name'].decode('utf-8') + u'同学，您好！由衷的感谢您选择北京展梦学业规划指导中心来为您规划学业，'
    else:
        txt = u'尊敬的家长，您好！由衷的感谢您选择北京展梦学业规划指导中心来为您的孩子规划学业，'
    txt += u'根据您提供的信息，我们为您的自主招生咨询作出如下建议，敬请参考，谢谢！'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,0,25,0,0)

    #一级标题，个人信息及专业建议
    txt = u'一、个人信息及专业建议'
    SetParagraphTextAndFormat(document,txt,20,u'黑体',0,0,0,0,0)

    #居中写表格名
    txt = u'表1 学生个人信息及专业建议'
    SetParagraphTextAndFormat(document,txt,10,u'黑体',1,0,0,False,False)

    #插入表格
    p = document.Paragraphs.Add()
    table = document.Tables.Add(p.Range, 10, 7)   # 新增一个10*7表格
    #设置表格为网格型
    #if table.Style <> u"网格型":
    table.Style = u"网格型"
    #设置部分表格的高度
    table.Cell(7,1).Height = 200
    table.Cell(8,1).Height = 60
    table.Cell(9,1).Height = 80
    table.Cell(10,1).Height = 100
    #设置部分表格的宽度
    #table.Cell(1,1).Width = 50
    #table.Cell(1,7).Width = 50
    '''
    table.Rows.Add()     # 新增一個Row
    table.Columns.Add()     # 新增一個Column
    '''
    #合并单元格
    cell = table.Cell(1,1)
    cell.Merge(table.Cell(5,1))#合并个人信息
    cell = table.Cell(6,1)
    cell.Merge(table.Cell(9,1))#合并MBTI性格测试项    

    cell = table.Cell(2,6)
    cell.Merge(table.Cell(2,7))#合并学校内容项
    cell = table.Cell(2,3)
    cell.Merge(table.Cell(2,4))#合并邮箱内容项
    cell = table.Cell(4,6)
    cell.Merge(table.Cell(4,7))#合并排斥学科内容项
    cell = table.Cell(4,3)
    cell.Merge(table.Cell(4,4))#合并优势学科内容项
    cell = table.Cell(5,6)
    cell.Merge(table.Cell(5,7))#合并意向城市内容项
    cell = table.Cell(5,3)
    cell.Merge(table.Cell(5,4))#合并意向城市内容项
    cell = table.Cell(6,3)
    cell.Merge(table.Cell(6,7))#合并性格类型内容项
    cell = table.Cell(7,3)
    cell.Merge(table.Cell(7,7))#合并性格描述内容项
    cell = table.Cell(8,3)
    cell.Merge(table.Cell(8,7))#合并适合领域内容项
    cell = table.Cell(9,3)
    cell.Merge(table.Cell(9,7))#合并适合职业内容项
    cell = table.Cell(10,2)
    cell.Merge(table.Cell(10,7))#合并建议专业内容项

    #表格字体
    fontsize = 10
    fontname = u'宋体'
    alignmentflag = 0
    leftindent = 0
    firstlfind = 0
    itaflag = False 
    boldflag = False
    #表格结构内容

    txt = u'\n\n个人\n信息'#个人信息
    SetTableTextAndFormat(table.Cell(1,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'姓名'
    SetTableTextAndFormat(table.Cell(1,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'性别'
    SetTableTextAndFormat(table.Cell(1,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'手机'
    SetTableTextAndFormat(table.Cell(1,6),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'邮箱'
    SetTableTextAndFormat(table.Cell(2,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'学校'
    SetTableTextAndFormat(table.Cell(2,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'分科'
    SetTableTextAndFormat(table.Cell(3,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'年级'
    SetTableTextAndFormat(table.Cell(3,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'成绩'
    SetTableTextAndFormat(table.Cell(3,6),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'优势学科'
    SetTableTextAndFormat(table.Cell(4,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'排斥学科'
    SetTableTextAndFormat(table.Cell(4,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'意向城市'
    SetTableTextAndFormat(table.Cell(5,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'意向高校水平'
    SetTableTextAndFormat(table.Cell(5,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'\n\n\n\n\n\n\n\n\nMBTI\n性格\n测试'
    SetTableTextAndFormat(table.Cell(6,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'性格类型'
    SetTableTextAndFormat(table.Cell(6,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'\n\n\n\n\n性格描述'
    SetTableTextAndFormat(table.Cell(7,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'\n\n适合领域'
    SetTableTextAndFormat(table.Cell(8,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'\n\n适合职业'
    SetTableTextAndFormat(table.Cell(9,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'\n建议\n选择\n的专\n业'
    SetTableTextAndFormat(table.Cell(10,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

    
    #表格内容
    txt = StuInfoData['name'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(1,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['gender'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(1,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['phonenum'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(1,7),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['email'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(2,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['school'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(2,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['artsorsci'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(3,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['grade'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(3,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['graderank'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(3,7),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['advantagesubject'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(4,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['disadvtsubject'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(4,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['intentunivlocation'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(5,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = StuInfoData['intentuniverrank'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(5,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTIData['type'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(6,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    #格式
    alignmentflag = 0
    firstlfind = 0#首行缩进已在原始数据中设置，所以不再设置

    txt = MBTIData['description'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(7,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTIData['field'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(8,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = MBTIData['profession'].decode('utf-8')
    SetTableTextAndFormat(table.Cell(9,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    
    #写入专业数据
    txt = SubMajData.decode('utf-8')
    SetTableTextAndFormat(table.Cell(10,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    
    
    #插入分页
    p = document.Paragraphs.Add()
    p.Range.InsertBreak()

def WriteUniversityInfoPage(document,StuUniversityApplyConditionsData):#写学校的信息
    '''
    功能：写学校信息
    '''
    #一级标题，推荐的学校
    txt = u'二、自主招生院校定位及申报建议'
    SetParagraphTextAndFormat(document,txt,20,u'黑体',0,0,0,False,False)
    #段前说明
    txt = u'根据您提供的数据，我们推荐您考虑以下院校的自主招生。需要请您注意的是2018年度自招的信息在2018年3月发布，'+\
    u'为方便您备考，下面给出相应学校在2017年度的自招信息。'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,0,25,False,False)#首行缩进20磅
    #居中写表名
    txt = u'表2 自主招生院校定位'
    SetParagraphTextAndFormat(document,txt,10,u'黑体',1,0,0,False,False)

    #增加表格写学校信息
    #先设置表格字体
    fontsize = 10
    fontname = u'宋体'
    alignmentflag = 0
    leftindent = 0
    firstlfind = 0
    itaflag = False 
    boldflag = False
    #创建表格
    p = document.Paragraphs.Add()
    table = document.Tables.Add(p.Range, 1, 6)   
    table.Style = u'网格型'
    
    txt = u'序号'
    SetTableTextAndFormat(table.Cell(1,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'学校名称'
    SetTableTextAndFormat(table.Cell(1,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'地点'
    SetTableTextAndFormat(table.Cell(1,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'学校层次'
    SetTableTextAndFormat(table.Cell(1,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'限报学校数'
    SetTableTextAndFormat(table.Cell(1,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    txt = u'招生人数'
    SetTableTextAndFormat(table.Cell(1,6),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    #再增加表格元素
    schoolnum = len(StuUniversityApplyConditionsData)
    for i in xrange(schoolnum):
        table.Rows.Add()     # 新增一個Row
        txt = str(i+1).decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,1),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['schoolname'].decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,2),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['schoollocation'].decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,3),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['schoolrank'].decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,4),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['schoollimit'].decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,5),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['plan'].decode('utf-8')
        SetTableTextAndFormat(table.Cell(i+2,6),txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

    fontsize = 12
    firstlfind = 20
    #分段写每个学校的招考要求
    txt = u'上表是北京展梦学业规划指导中心给您的自主招生院校定位，下面详细给出每所学校的招考要求，敬请您参考，谢谢！'
    SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)#首行缩进20磅
    ##循环填写学校的其他信息
    for i in range(schoolnum):
        ##二级标题，学校名称
        
        fontsize = 16
        fontname = u'黑体'
        leftindent = 0
        firstlfind = 0
        
        txt = (str(i+1) + '. ' + StuUniversityApplyConditionsData[i]['schoolname']).decode('utf-8')
        SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        #分条目写其他信息
        
        fontsize = 12
        fontname = u'宋体'
        leftindent = 20
        firstlfind2 = 30
        
        infoi = 1#报名时间
        txt = u'（' + str(infoi) + u'）报名时间：' 
        SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['applytime'].decode('utf-8')
        SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind2,itaflag,boldflag)
        #邮寄材料截止时间
        infoi += 1
        txt = u'（' + str(infoi) + u'）邮寄材料截止时间：'
        par = SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['materialpostdeadline'].decode('utf-8')
        par = SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind2,itaflag,boldflag)

        #笔面试时间
        infoi += 1
        txt = u'（' + str(infoi) + u'）笔面试时间：'
        SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        txt = StuUniversityApplyConditionsData[i]['examtime'].decode('utf-8')
        SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind2,itaflag,boldflag)
        
        #专利论文要求：
        paper = StuUniversityApplyConditionsData[i]['paper'].decode('utf-8')
        patent = StuUniversityApplyConditionsData[i]['patent'].decode('utf-8')
        if len(paper) > 0 or len(patent) > 0:

            leftindent = 20
            firstlfind = 0
            
            infoi += 1
            txt = u'（' + str(infoi) + u'）专利论文要求：'
            SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

            firstlfind = 30
            
            infoix = 0
            if len(paper) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 论文：' + paper
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(patent) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 专利：' + patent
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        #学科竞赛要求
        #先判断有无该项要求
        math = StuUniversityApplyConditionsData[i]['math'].decode('utf-8')
        phis = StuUniversityApplyConditionsData[i]['physics'].decode('utf-8')
        chem = StuUniversityApplyConditionsData[i]['chemistry'].decode('utf-8')
        bio = StuUniversityApplyConditionsData[i]['biology'].decode('utf-8')
        info = StuUniversityApplyConditionsData[i]['infomatics'].decode('utf-8')
        if len(math) > 0 or len(phis) > 0 or len(chem) > 0 or len(bio) > 0 or len(info) > 0:

            leftindent = 20
            firstlfind = 0
            
            infoi += 1
            txt = u'（' + str(infoi) + u'）学科竞赛要求：'
            SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

            firstlfind = 30
            
            infoix = 0
            if len(math) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 数学：' + math
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(phis) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 物理：' + phis
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(chem) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 化学：' + chem
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(bio) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 生物：' + bio
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(info) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 信息学：' + info
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
                
        #文科类竞赛要求
        newcon = StuUniversityApplyConditionsData[i]['newconceptcomposition'].decode('utf-8')
        innocom = StuUniversityApplyConditionsData[i]['innovativecomposition'].decode('utf-8')
        chinesenews = StuUniversityApplyConditionsData[i]['chinesenewspapercup'].decode('utf-8')
        yeshengtao = StuUniversityApplyConditionsData[i]['yeshengtaocup'].decode('utf-8')
        peking = StuUniversityApplyConditionsData[i]['pekinguculture'].decode('utf-8')
        innoeng = StuUniversityApplyConditionsData[i]['innovativeenglish'].decode('utf-8')
        engcom = StuUniversityApplyConditionsData[i]['englishcompetition'].decode('utf-8')
        if len(newcon) > 0 or len(innocom) > 0 or len(chinesenews) > 0 or len(yeshengtao) > 0 or len(peking) > 0 \
            or len(innoeng) > 0 or len(engcom) > 0:
            
            leftindent = 20
            firstlfind = 0
            
            infoi += 1
            txt = u'（' + str(infoi) + u'）文科类竞赛要求：'
            SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

            firstlfind = 30
            
            infoix = 0
            if len(newcon) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 新概念作文大赛：' + newcon
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(innocom) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 创新作文大赛：' + innocom
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(chinesenews) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 语文报杯：' + chinesenews
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(yeshengtao) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 叶圣陶杯：' + yeshengtao
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(peking) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 北大学培文化：' + peking
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(innoeng) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 创新英语竞赛：' + innoeng
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(engcom) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 英语能力竞赛：' + engcom
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
        
        #科技创新活动大赛
        young = StuUniversityApplyConditionsData[i]['youngsciinnocom'].decode('utf-8')
        tomm = StuUniversityApplyConditionsData[i]['tomorrowscientist'].decode('utf-8')
        prim = StuUniversityApplyConditionsData[i]['primsecdschcomputer'].decode('utf-8')
        intsci = StuUniversityApplyConditionsData[i]['intenationalscieng'].decode('utf-8')
        intenv = StuUniversityApplyConditionsData[i]['intenationalenvironmentsciprj'].decode('utf-8')

        if len(young) > 0 or len(tomm) > 0 or len(prim) > 0 or len(intsci) > 0 or len(intenv) > 0:

            leftindent = 20
            firstlfind = 0
            
            infoi += 1
            txt = u'（' + str(infoi) + u'）科技创新竞赛要求：'
            SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

            firstlfind = 30
            
            infoix = 0
            if len(young) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 青少年科技创新大赛：' + young
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(tomm) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 明天小小科学技术科创大赛：' + tomm
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(prim) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 中小学电脑制作大赛：' + prim
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(intsci) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 国际科学与工程竞赛：' + intsci
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
            if len(intenv) > 0:
                infoix += 1
                txt = u'[' + str(infoix) + u'] 国际环境科研项目：' + intenv
                SetParagraphTextAndFormat(document,txt,fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)
    
        ###############################################
    #写结束语
    #下面再输入3行空行
    SetParagraphTextAndFormat(document,'\n\n',fontsize,fontname,alignmentflag,leftindent,firstlfind,itaflag,boldflag)

    #段前说明
    txt = u'上述内容为北京展梦学业规划指导中心根据您提供的信息为您提供的自主招生学校定位信息，敬请参考！如对结果有疑问，请致电400-88888888' + \
    u'或发送邮件至学业邮箱：xf700@qq.com 联系北京展梦学业规划指导中心的黄老师（学业规划）或侯老师（自主招生）！'
    SetParagraphTextAndFormat(document,txt,12,u'宋体',0,0,20,False,False)#首行缩进20磅 



#将信息写入word文件总调用函数
def WriteStuAndUinversityDatatoDoc(FilePathStu,iStu):
    '''
    功能：写doc文件总调用函数
    '''
    curospath = os.path.abspath('.')
    FilePathMBTI = curospath + '//data//MBTI_Data.xlsx'

    flag,MBTIData = rsui.ReadMBTIDataFromExcel(FilePathMBTI)
    if flag == 0:
        return 0
    #读取不同学生和学校的报考信息
    StuData,sugmajor,stumbti,stuanduniverinfo = dataprocess.GetAllStudentsUniversityandApplyConditions(FilePathStu,iStu)        
    
    stunum = len(StuData)

    LoopList = []
    if len(iStu) == 1:
        if iStu[0] >= 0 and iStu[0] < stunum:
            LoopList.append(iStu[0])
        else:
            LoopList = range(stunum)
    else:
        #检查有效性
        for i in range(len(iStu)):
            if iStu[i] < 0 or iStu[i] >= stunum:
                del iStu[i]
        LoopList = iStu
        
    #word写入准备
    #一次打开word引擎
    #打开word引擎
    w = win32com.client.Dispatch('Word.Application') 
    # 后台运行，不显示，不警告
    w.Visible = 1
    w.DisplayAlerts = 0 

    ip = 0
    for i in LoopList:#循环写入
        document = w.Documents.Add() # 创建新的文档，对于每个学生创建一个新文档
        WriteFrontCoverPage(document,StuData[i])#写封面页
        WriteFirstPage(document,StuData[i],MBTIData[stumbti[ip]],sugmajor[ip])#写第一页
        WriteUniversityInfoPage(document,stuanduniverinfo[ip])#写学校的信息
        ip += 1
        docpath = curospath + '\\报告\\' + StuData[i]['name'] + '.docx'#str(i) + '.docx'#
        docpath = docpath.decode('utf-8')
        document.SaveAs(docpath)
        document.Close()

    w.Quit()
    
    return 1

'''
if __name__=="__main__":
    FilePath = r'C:\Users\Think\Desktop\data\stuinfo.xlsx'
    iStu = [0]
    WriteStuAndUinversityDatatoDoc(FilePath,iStu)
'''
'''
#打开word引擎
w = win32com.client.Dispatch('Word.Application') 
# 后台运行，不显示，不警告
w.Visible = 1 
w.DisplayAlerts = 0
document = w.Documents.Add() # 创建新的文档
##先添加图像（公司logo图片为logo.png）

StuData = ['a','b']
#WriteFrontCoverPage(document,StuData[0])#写封面页
MBTIData = ''
sugmajor = ''
WriteFirstPage(document,StuData,MBTIData,sugmajor)
'''

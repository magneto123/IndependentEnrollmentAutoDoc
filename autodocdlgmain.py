# -*- coding: utf-8 -*- 

###########################################################################
## Python code generated with wxFormBuilder (version Jun 17 2015)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################
import os
import wx
#import wx.xrc

import readstuandunivinfo as rsu
import dataprocess
#from docx import Document
import writedatatodocwin32 as wdd

import productaccredit as pa

from win32com import client as wc#将输出的文档另存为pdf文件
###########################################################################
## Class AutoDocDlg
###########################################################################

class AutoDocDlg ( wx.Dialog ):
    
    def __init__( self, parent ):
        wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"一元咨询产品报告自动编写软件 - 先峰教育", pos = wx.DefaultPosition, size = wx.Size( 601,406 ), style = wx.DEFAULT_DIALOG_STYLE )
        
        self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
        
        fgSizer1 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer1.SetFlexibleDirection( wx.BOTH )
        fgSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_buttonOpenStuInfoFile = wx.Button( self, wx.ID_ANY, u"打开问卷信息文件", wx.DefaultPosition, wx.Size( 150,-1 ), 0 )
        fgSizer1.Add( self.m_buttonOpenStuInfoFile, 0, wx.ALL, 5 )
        
        self.m_textStuInfoFilePath = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.Size( 420,-1 ), wx.TE_READONLY )
        fgSizer1.Add( self.m_textStuInfoFilePath, 0, wx.ALL, 5 )
        
        bSizer1 = wx.BoxSizer( wx.VERTICAL )
        
        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, u"学生信息列表：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )
        bSizer1.Add( self.m_staticText1, 0, wx.ALL, 5 )
        
        m_listBoxStuInfoListChoices = []
        self.m_listBoxStuInfoList = wx.ListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( 150,300 ), m_listBoxStuInfoListChoices, wx.LB_ALWAYS_SB)#|wx.LB_MULTIPLE )
        bSizer1.Add( self.m_listBoxStuInfoList, 0, wx.ALL, 5 )
        
        
        fgSizer1.Add( bSizer1, 1, wx.EXPAND, 5 )
        
        bSizer2 = wx.BoxSizer( wx.VERTICAL )
        
        fgSizer2 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer2.SetFlexibleDirection( wx.BOTH )
        fgSizer2.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        fgSizer3 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer3.SetFlexibleDirection( wx.BOTH )
        fgSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_staticText2 = wx.StaticText( self, wx.ID_ANY, u"姓名：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText2.Wrap( -1 )
        fgSizer3.Add( self.m_staticText2, 0, wx.ALL, 5 )
        
        self.m_textName = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textName, 0, wx.ALL, 5 )
        
        self.m_staticText3 = wx.StaticText( self, wx.ID_ANY, u"手机：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText3.Wrap( -1 )
        fgSizer3.Add( self.m_staticText3, 0, wx.ALL, 5 )
        
        self.m_textPhonenum = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textPhonenum, 0, wx.ALL, 5 )
        
        self.m_staticText4 = wx.StaticText( self, wx.ID_ANY, u"学校：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText4.Wrap( -1 )
        fgSizer3.Add( self.m_staticText4, 0, wx.ALL, 5 )
        
        self.m_textSchool = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textSchool, 0, wx.ALL, 5 )
        
        self.m_staticText5 = wx.StaticText( self, wx.ID_ANY, u"分科：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText5.Wrap( -1 )
        fgSizer3.Add( self.m_staticText5, 0, wx.ALL, 5 )
        
        self.m_textArtsorsci = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textArtsorsci, 0, wx.ALL, 5 )
        
        self.m_staticText6 = wx.StaticText( self, wx.ID_ANY, u"优势学科：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText6.Wrap( -1 )
        fgSizer3.Add( self.m_staticText6, 0, wx.ALL, 5 )
        
        self.m_textAdvsubject = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textAdvsubject, 0, wx.ALL, 5 )
        
        self.m_staticText7 = wx.StaticText( self, wx.ID_ANY, u"意向大学位置：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText7.Wrap( -1 )
        fgSizer3.Add( self.m_staticText7, 0, wx.ALL, 5 )
        
        self.m_textIntentunivloca = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer3.Add( self.m_textIntentunivloca, 0, wx.ALL, 5 )
        
        
        fgSizer2.Add( fgSizer3, 1, wx.EXPAND, 5 )
        
        fgSizer4 = wx.FlexGridSizer( 0, 2, 0, 0 )
        fgSizer4.SetFlexibleDirection( wx.BOTH )
        fgSizer4.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_staticText8 = wx.StaticText( self, wx.ID_ANY, u"性别：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText8.Wrap( -1 )
        fgSizer4.Add( self.m_staticText8, 0, wx.ALL, 5 )
        
        self.m_textGender = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textGender, 0, wx.ALL, 5 )
        
        self.m_staticText9 = wx.StaticText( self, wx.ID_ANY, u"邮箱：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText9.Wrap( -1 )
        fgSizer4.Add( self.m_staticText9, 0, wx.ALL, 5 )
        
        self.m_textEmail = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textEmail, 0, wx.ALL, 5 )
        
        self.m_staticText10 = wx.StaticText( self, wx.ID_ANY, u"年级：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText10.Wrap( -1 )
        fgSizer4.Add( self.m_staticText10, 0, wx.ALL, 5 )
        
        self.m_textGrade = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textGrade, 0, wx.ALL, 5 )
        
        self.m_staticText11 = wx.StaticText( self, wx.ID_ANY, u"成绩：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText11.Wrap( -1 )
        fgSizer4.Add( self.m_staticText11, 0, wx.ALL, 5 )
        
        self.m_textGraderank = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textGraderank, 0, wx.ALL, 5 )
        
        self.m_staticText12 = wx.StaticText( self, wx.ID_ANY, u"排斥学科：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText12.Wrap( -1 )
        fgSizer4.Add( self.m_staticText12, 0, wx.ALL, 5 )
        
        self.m_textDisadvsubject = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textDisadvsubject, 0, wx.ALL, 5 )
        
        self.m_staticText13 = wx.StaticText( self, wx.ID_ANY, u"意向大学层次：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText13.Wrap( -1 )
        fgSizer4.Add( self.m_staticText13, 0, wx.ALL, 5 )
        
        self.m_textIntentunivrank = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        fgSizer4.Add( self.m_textIntentunivrank, 0, wx.ALL, 5 )
        
        
        fgSizer2.Add( fgSizer4, 1, wx.EXPAND, 5 )
        
        
        bSizer2.Add( fgSizer2, 1, wx.EXPAND, 5 )
        
        gbSizer3 = wx.GridBagSizer( 0, 80 )
        gbSizer3.SetFlexibleDirection( wx.BOTH )
        gbSizer3.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_staticText16 = wx.StaticText( self, wx.ID_ANY, u"MBTI性格测试分类：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText16.Wrap( -1 )
        gbSizer3.Add( self.m_staticText16, wx.GBPosition( 0, 0 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        '''
        m_choiceMBTIChoices = ['','ENFJ','ENFP','ENTJ','ENTP','ESFJ','ESFP','ESTJ','ESTP',\
        'INFJ','INFP','INTJ','INTP','ISFJ','ISFP','ISTJ','ISTP']
        self.m_choiceMBTI = wx.Choice( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( 120,-1 ), m_choiceMBTIChoices, 0 )
        self.m_choiceMBTI.SetSelection( 0 )
        gbSizer3.Add( self.m_choiceMBTI, wx.GBPosition( 0, 1 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        '''
        self.m_textMbtiType = wx.TextCtrl( self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, wx.TE_READONLY )
        gbSizer3.Add( self.m_textMbtiType, wx.GBPosition( 0, 1 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        bSizer2.Add( gbSizer3, 1, wx.EXPAND, 5 )
        
        gbSizer1 = wx.GridBagSizer( 0, 10 )
        gbSizer1.SetFlexibleDirection( wx.BOTH )
        gbSizer1.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_staticText15 = wx.StaticText( self, wx.ID_ANY, u"文件输出：", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText15.Wrap( -1 )
        gbSizer1.Add( self.m_staticText15, wx.GBPosition( 0, 0 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        self.m_checkBoxPDFFile = wx.CheckBox( self, wx.ID_ANY, u"同时保存为PDF文件", wx.DefaultPosition, wx.DefaultSize, 0 )
        gbSizer1.Add( self.m_checkBoxPDFFile, wx.GBPosition( 0, 2 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        
        bSizer2.Add( gbSizer1, 1, wx.EXPAND, 5 )
        
        gbSizer2 = wx.GridBagSizer( 0, 20 )
        gbSizer2.SetFlexibleDirection( wx.BOTH )
        gbSizer2.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
        
        self.m_buttonCreateDocument = wx.Button( self, wx.ID_ANY, u"生成报告", wx.DefaultPosition, wx.Size( 180,-1 ), 0 )
        gbSizer2.Add( self.m_buttonCreateDocument, wx.GBPosition( 0, 0 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        #self.m_buttonCreateAndOpenDocument = wx.Button( self, wx.ID_ANY, u"生成报告并打开", wx.DefaultPosition, wx.Size( 120,-1 ), 0 )
        #gbSizer2.Add( self.m_buttonCreateAndOpenDocument, wx.GBPosition( 0, 1 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        self.m_buttonCreateBatchDocument = wx.Button( self, wx.ID_ANY, u"批量生成报告", wx.DefaultPosition, wx.Size( 180,-1 ), 0 )
        gbSizer2.Add( self.m_buttonCreateBatchDocument, wx.GBPosition( 0, 2 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
        
        
        bSizer2.Add( gbSizer2, 1, wx.EXPAND, 5 )
        '''
        self.m_gaugeProgress = wx.Gauge( self, wx.ID_ANY, 100, wx.DefaultPosition, wx.Size( 420,10 ), wx.GA_HORIZONTAL )
        self.m_gaugeProgress.SetValue( 0 ) 
        bSizer2.Add( self.m_gaugeProgress, 0, wx.ALL, 5 )
        '''
        fgSizer1.Add( bSizer2, 1, wx.EXPAND, 5 )
        
        
        self.SetSizer( fgSizer1 )
        self.Layout()
        
        self.Centre( wx.BOTH )
        
        # Connect Events
        self.m_buttonOpenStuInfoFile.Bind( wx.EVT_BUTTON, self.OnOpenStuInfoFile )
        self.m_buttonCreateDocument.Bind( wx.EVT_BUTTON, self.OnCreateDocument )
        #self.m_buttonCreateAndOpenDocument.Bind( wx.EVT_BUTTON, self.OnCreateAndOpenDocument )
        self.m_buttonCreateBatchDocument.Bind( wx.EVT_BUTTON, self.OnCreateBatchDocument )
        self.m_listBoxStuInfoList.Bind( wx.EVT_LISTBOX, self.OnChooseAStudent )
    
    def __del__( self ):
        pass
    
    #将信息写入word文件总调用函数
    def WriteStuAndUinversityDatatoDoc(self,FilePathStu,iStu,pdfflag):
        '''
        功能：写doc文件总调用函数
        '''
        #计算状态提示
        if len(iStu) == 1:
            blabel = u'生成报告'
            button = self.m_buttonCreateDocument
        else:
            blabel = u'批量生成报告'
            button = self.m_buttonCreateBatchDocument
            
        vstr = blabel + u'（初始化中...）'
        button.SetLabel(vstr)
        
        curospath = os.path.abspath('.')
        FilePathMBTI = curospath + '\\data\\MBTI_Data.xlsx'
        #创建一个新的目录用于存放生成的报告
        newdir = curospath + u'\\报告'.encode('gbk')
        if not os.path.exists(newdir):
            os.mkdir(newdir)
            
        flag,MBTIData = rsu.ReadMBTIDataFromExcel(FilePathMBTI)
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
            
        vstr = blabel + u'（' + str(0).decode('utf-8') + u'%...）'
        button.SetLabel(vstr)

        #如果要转为pdf文件，先启动wordapplication
        wordapp = wc.Dispatch('Word.Application')
        wordapp.Visible = 0
        wordapp.DisplayAlerts = 0
        
        ip = 0
        for i in LoopList:#循环写入
            document = wordapp.Documents.Add()
            wdd.WriteFrontCoverPage(document,StuData[i])#写封面页
            wdd.WriteFirstPage(document,StuData[i],MBTIData[stumbti[ip]],sugmajor[ip])#写第一页
            wdd.WriteUniversityInfoPage(document,stuanduniverinfo[ip])#写学校的信息
            ip += 1
            docpath = newdir + '\\' + StuData[i]['name'].decode('utf-8').encode('gbk') + '.docx'#str(i) + '.docx'#
            document.SaveAs(docpath)
            
            #转为pdf格式
            if pdfflag:
                pdfpath = newdir + '\\' + StuData[i]['name'].decode('utf-8').encode('gbk') + '.pdf'
                document.SaveAs(pdfpath, 17) #17对应于下表中的pdf文件
            document.Close() 
            
            #将计算进度显示到按钮标签上
            prog = int(float(ip)/float(len(LoopList))*100)
            vstr = blabel + u'（' + str(prog).decode('utf-8') + u'%...）'
            button.SetLabel(vstr)
            

        wordapp.Quit()

        #完成，将标签置为初始值
        button.SetLabel(vstr)
        #提示信息
        txt = u'当前工作完成，总共导出' + str(len(LoopList)).decode('utf-8') + u'份报告!'
        wx.MessageBox(txt)
         
        return 1
    
    # Virtual event handlers, overide them in your derived class
    #将信息添加到对话框中
    def FillStuDataIntoDlg(self,StuData,i):
        '''
        功能：将数据填充到对话框中
        '''
        self.m_textName.SetValue(StuData[i]['name'].decode('utf-8'))
        self.m_textGender.SetValue(StuData[i]['gender'].decode('utf-8'))
        self.m_textPhonenum.SetValue(StuData[i]['phonenum'].decode('utf-8'))
        self.m_textEmail.SetValue(StuData[i]['email'].decode('utf-8'))
        self.m_textSchool.SetValue(StuData[i]['school'].decode('utf-8'))
        self.m_textGrade.SetValue(StuData[i]['grade'].decode('utf-8'))
        self.m_textArtsorsci.SetValue(StuData[i]['artsorsci'].decode('utf-8'))
        self.m_textGraderank.SetValue(StuData[i]['graderank'].decode('utf-8'))
        advsubject = StuData[i]['advantagesubject'].decode('utf-8')
        self.m_textAdvsubject.SetValue(advsubject)
        disadvsubject = StuData[i]['disadvtsubject'].decode('utf-8')
        self.m_textDisadvsubject.SetValue(disadvsubject)
        self.m_textIntentunivloca.SetValue(StuData[i]['intentunivlocation'].decode('utf-8'))
        self.m_textIntentunivrank.SetValue(StuData[i]['intentuniverrank'].decode('utf-8'))
        self.m_textMbtiType.SetValue(StuData[i]['mbtitype'].decode('utf-8'))
        
    def OnOpenStuInfoFile( self, event ):
    #打开文件对话框选择一个xls或xlsx文件
        wildcard = "Excel Files (*.xls)|*.xls |Excel Files (*.xlsx)|*.xlsx" 
        dlg = wx.FileDialog(self, "Choose a file", os.getcwd().decode('gbk'), "", wildcard, wx.FD_OPEN) 
    
        if dlg.ShowModal() == wx.ID_OK:
            self.InfoFilePath = dlg.GetPath()
        else:
            self.InfoFilePath = ''
        
        if len(self.InfoFilePath) > 10:
            #将文件路径添加到文本框中
            self.m_textStuInfoFilePath.SetValue(self.InfoFilePath)
            
            #读取文件中的数据并将文件填充到列表中
            flag,self.StuInfoArray = rsu.ReadStuInfoDataFromExcel(self.InfoFilePath)
            if flag == 0:
                wx.MessageBox(u'学生信息文件不存在！')
                return
            
            stunum = len(self.StuInfoArray)
            self.m_listBoxStuInfoList.Clear()#先清空列表框
            for i in range(stunum):
                string = self.StuInfoArray[i]['name'].decode('utf-8')
                self.m_listBoxStuInfoList.Append(string)
            #设置选中项
            self.m_listBoxStuInfoList.SetSelection(0)
            #将选中项的数据填充到对话框
            self.FillStuDataIntoDlg(self.StuInfoArray,0)
    
    def OnCreateDocument( self, event ):
        if len(self.InfoFilePath) > 10:
            stuchoosei = self.m_listBoxStuInfoList.GetSelection()
            pdff = self.m_checkBoxPDFFile.GetValue()
            if stuchoosei > -1:
                iStu = [stuchoosei]
                self.WriteStuAndUinversityDatatoDoc(self.InfoFilePath,iStu,pdff)
            else:
                wx.MessageBox(u'请选择一个学生信息!')
                return
        else:
            wx.MessageBox(u'请先打开学生信息!')
            return
    
    #def OnCreateAndOpenDocument( self, event ):
    #    event.Skip()
    
    def OnCreateBatchDocument( self, event ):
        if len(self.InfoFilePath) > 10:
            pdff = self.m_checkBoxPDFFile.GetValue()

            iStu = range(len(self.StuInfoArray))
            self.WriteStuAndUinversityDatatoDoc(self.InfoFilePath,iStu,pdff)
        else:
            wx.MessageBox(u'请先打开学生信息!')
            return
                
    def OnChooseAStudent( self, event ):
        '''
        功能：选中一个学生然后显示他的信息
        '''
        choosei = self.m_listBoxStuInfoList.GetSelection()
        self.FillStuDataIntoDlg(self.StuInfoArray,choosei)
        
    
if __name__=="__main__":
    '''
    '''
    app = wx.App(False) 
    #先检查软件是否注册
    flag = pa.ProductAccredit()
    flag = 1
    if flag == 1:
        dlg = AutoDocDlg(None) 
        dlg.Show(True) 
        #start the applications 
        app.MainLoop()
    else:
        if flag == 0:
            wx.MessageBox(u'软件注册信息不正确，请联系管理员重新注册!')
        if flag == -1:
            wx.MessageBox(u'未找到注册文件，请联系管理员进行注册!')

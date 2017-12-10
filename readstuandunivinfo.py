#-*- coding:utf-8 -*-
import numpy as np
import xlrd
import os #判断文件是否存在

StudentInfoDataType = np.dtype({'names': ['identity','name','gender', 'phonenum','qq','email','artsorsci',\
                                             'school', 'grade', 'graderank', 'intentunivlocation','intentaftgradu',\
                                             'intentuniverrank','intentunivercategory', 'advantagesubject','disadvtsubject','mbtitype'],
                            'formats': ['S50', 'S50', 'S20', 'S30', 'S30', 'S50', 'S50', 'S50', 'S50', 'S50', 'S100', 'S50',\
                                         'S50', 'S50', 'S50','S50', 'S50','S50']}, align = True)

# Universities = np.array([("清华大学","1003",1,140,10,150,8,145,13,135,23,150,15,"一本","综合","985","中国牛逼高校联盟","计算机、建筑、土木工程"),\
                         # ("中国石油大学（华东）","1023",1,80,10000,90,12100,80,23482,78,10342,80,12345,"一本","工科","211/(小985)","石油科技人才的摇篮","石油工程、资源勘查工程")],
    # dtype=UniversityInfoDataType)
#MBTI性格测试数据格式
MBTIDataType = np.dtype({'names': ['type','description','field', 'profession'],
                            'formats': ['S20', 'S3000', 'S500', 'S1500']}, align = True)
                            
#学科和专业对照表
SubjectAndMajorDataType = np.dtype({'names': ['subject','major'],
                            'formats': ['S20', 'S1000']}, align = True)

#自主招生信息数据格式
#自主招生高校名单
UniversityDataType = np.dtype({'names': ['schoollocation','schoolname','schoolrank','schoolcategory'],
                            'formats': ['S50', 'S50','S50', 'S50']}, align = True)
                            
#自主招生计划和考试时间
UniversityPlanAndExamTimeDataType = np.dtype({'names': ['schoollocation','schoolname','schoollimit',\
    'plan','applytime','materialpostdeadline','examtime'],
                            'formats': ['S50', 'S50','S30','S50','S300','S300','S300']}, align = True)
#综合成绩
ComGradeDataType = np.dtype({'names': ['schoolname','comgradcondition'],
                            'formats': ['S50', 'S1000']}, align = True)
#期刊论文                            
PaperAndPatentDataType = np.dtype({'names': ['schoollocation','schoolname','paper', 'patent'],
                            'formats': ['S20', 'S50', 'S500', 'S300']}, align = True)
#五大学科竞赛                            
SubjectCompetitionDataType = np.dtype({'names': ['schoollocation','schoolname','math', 'physics','chemistry','biology','infomatics'],
                            'formats': ['S20', 'S50', 'S300', 'S300','S300', 'S300','S300']}, align = True)                            
   
#五大学科竞赛                            
ArtsCompetitionDataType = np.dtype({'names': ['schoollocation','schoolname','newconceptcomposition', 'innovativecomposition','chinesenewspapercup','yeshengtaocup','pekinguculture','innovativeenglish','englishcompetition'],
                            'formats': ['S20', 'S50', 'S300', 'S300','S300', 'S300','S300','S300', 'S300','S300']}, align = True)
   
#科技创新活动                            
SciTechInnovationDataType = np.dtype({'names': ['schoollocation','schoolname','youngsciinnocom', 'tomorrowscientist','primsecdschcomputer','intenationalscieng','intenationalenvironmentsciprj'],
                            'formats': ['S20', 'S50', 'S300', 'S300','S300', 'S300','S300','S300']}, align = True)
   
    
def ReadStuInfoDataFromExcel(FilePath):
    '''
    功能：读取问卷信息
    '''
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0
        
    data = xlrd.open_workbook(FilePath)
    table = data.sheet_by_index(0)
        
    nrows = 0
    irow = 2#寻找数据的起始行
    while(True):
        data = table.cell(irow,0).value
        if len(data) > 0:
            irow += 1
        else:
            break
    
    StuInfoList = []
    for i in range(2,irow):
        identity = table.cell(i,2).value.encode('utf-8')
        name = table.cell(i,3).value.encode('utf-8')
        gender = table.cell(i,4).value.encode('utf-8')
        phonenum = table.cell(i,5).value.encode('utf-8')
        qq = table.cell(i,6).value.encode('utf-8')
        email = table.cell(i,7).value.encode('utf-8')
        school = table.cell(i,8).value.encode('utf-8')
        artsorsci = table.cell(i,9).value.encode('utf-8')
        grade = table.cell(i,10).value.encode('utf-8')
        graderank = table.cell(i,11).value.encode('utf-8')
        #读取意向大学省份
        intentunivlocation1 = table.cell(i,13).value.encode('utf-8')
        intentunivlocation2 = table.cell(i,14).value.encode('utf-8')
        intentunivlocation3 = table.cell(i,15).value.encode('utf-8')

        intentunivlocation = intentunivlocation1.split(' ')[1] + '/' + \
                             intentunivlocation2.split(' ')[1] + '/' + \
                             intentunivlocation3.split(' ')[1];

        intentuniverrank = table.cell(i,16).value.encode('utf-8')#预期高校水平
        intentunivercategory = table.cell(i,17).value.encode('utf-8')#预期高校类别
        
        intentaftgradu = table.cell(i,18).value.encode('utf-8')#毕业后意向
        
        
        advantagesubject = table.cell(i,19).value.encode('utf-8') + '/' + table.cell(i,20).value.encode('utf-8')
        disadvtsubject = table.cell(i,21).value.encode('utf-8') + '/' + table.cell(i,22).value.encode('utf-8')

        mbtiqt1 = table.cell(i,25).value.encode('utf-8')
        mbtiqt2 = table.cell(i,26).value.encode('utf-8')
        mbtiqt3 = table.cell(i,27).value.encode('utf-8')
        mbtiqt4 = table.cell(i,28).value.encode('utf-8')

        mbti = ''
        if 'A' in mbtiqt1:
            mbti += 'E'
        if 'B' in mbtiqt1:
            mbti += 'I'
        if 'A' in mbtiqt2:
            mbti += 'S'
        if 'B' in mbtiqt2:
            mbti += 'N'
        if 'A' in mbtiqt3:
            mbti += 'T'
        if 'B' in mbtiqt3:
            mbti += 'F'
        if 'A' in mbtiqt4:
            mbti += 'J'
        if 'B' in mbtiqt4:
            mbti += 'P'

        #mbti = table.cell(i,20).value.encode('utf-8')

        stinfo = (identity,name,gender,phonenum,qq,email,artsorsci,school,grade,\
            graderank,intentunivlocation,intentaftgradu,intentuniverrank,intentunivercategory,\
            advantagesubject,disadvtsubject,mbti)

        StuInfoList.append(stinfo)

    StuInfoDataArray = np.array(StuInfoList,dtype = StudentInfoDataType)
    return 1,StuInfoDataArray
    
#读取MBTI性格测试数据
def ReadMBTIDataFromExcel(FilePath):
    '''
    功能：读取MBTI性格测试数据
    '''
    
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0
        
    data = xlrd.open_workbook(FilePath)
    table = data.sheet_by_index(0)
    
    MbtiDataList = []
    for i in range(1,17):
        type = table.cell(i,1).value.encode('utf-8')
        description = table.cell(i,2).value.encode('utf-8')
        field = table.cell(i,3).value.encode('utf-8')
        profession = table.cell(i,4).value.encode('utf-8')

        mbtid = (type,description,field,profession)

        MbtiDataList.append(mbtid)

    MBTIDataArray = np.array(MbtiDataList,dtype = MBTIDataType)
    return 1,MBTIDataArray
    
    
#读取学科和专业对照表中的数据
def ReadSubjectAndMajor(FilePath):
    '''
    功能：读取MBTI性格测试数据
    '''
    
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0
  
    data = xlrd.open_workbook(FilePath)
    table = data.sheet_by_index(0)
    
    DList = []
    for i in range(2,13):
        sub = table.cell(i,0).value.encode('utf-8')
        maj = table.cell(i,1).value.encode('utf-8')

        d = (sub,maj)
        DList.append(d)

    SubMajDataArray = np.array(DList,dtype = SubjectAndMajorDataType)
    return 1,SubMajDataArray
    
#读取211和985工程大学数据
def ReadP211And985Universities(FilePath):
    '''
    功能：读取211和985工程大学数据
    '''
    
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0,0
        
    data = xlrd.open_workbook(FilePath)
    table = data.sheet_by_index(0)
    
    U211List = []
    U985List = []
    for i in range(2,41):
        U985 = table.cell(i,1).value.encode('utf-8')
        U985List.append(U985)
    for i in range(2,118):
        U211 = table.cell(i,3).value.encode('utf-8')
        U211List.append(U211)
    return 1,U211List,U985List
    
def ReadIndependentEnrollmentInfos(FilePath):
    '''
    功能：读取自主招生有关数据
    '''
    
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0,0,0,0,0,0,0
        
    data = xlrd.open_workbook(FilePath)
    
    #读取2017年自主招生高校名单数据
    table = data.sheet_by_index(0)
    UDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        prov = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        rank = table.cell(i,2).value.encode('utf-8')
        category = table.cell(i,3).value.encode('utf-8')
        d = (prov,univ,rank,category)
        UDataList.append(d)
    IndEnrollUniversityDataArray = np.array(UDataList,dtype = UniversityDataType)
    
    #读取学校计划和考试时间
    table = data.sheet_by_index(1)
    DataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        loca = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        limit = table.cell(i,3).value.encode('utf-8')
        plan = table.cell(i,4).value.encode('utf-8')
        apptime = table.cell(i,6).value.encode('utf-8')
        postdead = table.cell(i,7).value.encode('utf-8')
        examtime = table.cell(i,8).value.encode('utf-8')
        d = (loca,univ,limit,plan,apptime,postdead,examtime)
        DataList.append(d)
    UniversityPlanAndExamTimeDataArray = np.array(DataList,dtype = UniversityPlanAndExamTimeDataType)
    
    #读取综合成绩数据
    table = data.sheet_by_index(2)
    ComGradeDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        univ = table.cell(i,0).value.encode('utf-8')
        limitcondition = table.cell(i,1).value.encode('utf-8')
        comgrad = (univ,limitcondition)
        ComGradeDataList.append(comgrad)
    ComGradeDataArray = np.array(ComGradeDataList,dtype = ComGradeDataType)
    
    #读取期刊论文和专利发明数据
    table = data.sheet_by_index(3)
    PaperandPatentDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        provence = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        paper = table.cell(i,2).value.encode('utf-8')
        patent = table.cell(i,3).value.encode('utf-8')
        pappatd = (provence,univ,paper,patent)
        PaperandPatentDataList.append(pappatd)
    PaperAndPatentDataArray = np.array(PaperandPatentDataList,dtype = PaperAndPatentDataType)
    
    #读取五大学科竞赛数据
    table = data.sheet_by_index(4)
    SubComDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        provence = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        math = table.cell(i,2).value.encode('utf-8')
        phis = table.cell(i,3).value.encode('utf-8')
        chem = table.cell(i,4).value.encode('utf-8')
        bio = table.cell(i,5).value.encode('utf-8')
        info = table.cell(i,6).value.encode('utf-8')
        d = (provence,univ,math,phis,chem,bio,info)
        SubComDataList.append(d)
    SubjectCompetitionDataArray = np.array(SubComDataList,dtype = SubjectCompetitionDataType)
    
    #读取文科类竞赛数据
    table = data.sheet_by_index(5)
    ArtsDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        provence = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        newconcept = table.cell(i,2).value.encode('utf-8')
        innocom = table.cell(i,3).value.encode('utf-8')
        chinesenews = table.cell(i,4).value.encode('utf-8')
        yeshengtao = table.cell(i,5).value.encode('utf-8')
        peking = table.cell(i,6).value.encode('utf-8')
        innoeng = table.cell(i,7).value.encode('utf-8')
        engcom = table.cell(i,8).value.encode('utf-8')
        d = (provence,univ,newconcept,innocom,chinesenews,yeshengtao,peking,innoeng,engcom)
        ArtsDataList.append(d)
    ArtsCompetitionDataArray = np.array(ArtsDataList,dtype = ArtsCompetitionDataType)
    
    #读取科技创新类竞赛数据
    table = data.sheet_by_index(6)
    SciDataList = []
    irow = 2#寻找数据的起始行
    rows = table.nrows#先获取总行数，下面检查有效数据的行数
    while(irow < rows):
        v = table.cell(irow,0).value
        if len(v) > 0:
            irow += 1
        else:
            break
    for i in range(2,irow):
        provence = table.cell(i,0).value.encode('utf-8')
        univ = table.cell(i,1).value.encode('utf-8')
        young = table.cell(i,2).value.encode('utf-8')
        tomorrow = table.cell(i,3).value.encode('utf-8')
        computer = table.cell(i,4).value.encode('utf-8')
        intsci = table.cell(i,5).value.encode('utf-8')
        intenv = table.cell(i,6).value.encode('utf-8')
        d = (provence,univ,young,tomorrow,computer,intsci,intenv)
        SciDataList.append(d)
    SciTechInnovationDataArray = np.array(SciDataList,dtype = SciTechInnovationDataType)
    return 1,IndEnrollUniversityDataArray,UniversityPlanAndExamTimeDataArray,ComGradeDataArray,\
        PaperAndPatentDataArray,SubjectCompetitionDataArray,ArtsCompetitionDataArray,SciTechInnovationDataArray
''' 
if __name__=="__main__":   
    FilePath = 'C:\Users\CXJ-PC\Desktop\stuinfo.xlsx'
    StuData = ReadStuInfoDataFromExcel(FilePath)
    printStuData(StuData)
'''

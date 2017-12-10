#-*- coding:utf-8 -*-
'''
说明：数据处理，查找符号条件的学校
'''
import os
import numpy as np
import readstuandunivinfo as rsi

#自主招生高校及报考条件
IndEnrollUniversityAndApplyConditionsDataType = np.dtype({'names': ['schoolname','schoollocation','schoolrank','schoolcategory','schoollimit',\
        'plan','applytime','materialpostdeadline','examtime','comgradcondition',\
        'paper', 'patent','math', 'physics','chemistry','biology','infomatics','newconceptcomposition', 'innovativecomposition',\
        'chinesenewspapercup','yeshengtaocup','pekinguculture','innovativeenglish','englishcompetition',\
        'youngsciinnocom', 'tomorrowscientist','primsecdschcomputer','intenationalscieng','intenationalenvironmentsciprj'],
    'formats': ['S50', 'S50', 'S50','S50','S30','S50','S300','S300','S300','S1000', 'S500', 'S300', 'S300', 'S300','S300', 'S300','S300',\
        'S300', 'S300','S300', 'S300','S300','S300', 'S300','S300','S300', 'S300','S300', 'S300','S300','S300']}, align = True)

#定义数据文件路径
#获取 当前路径
curospath = os.path.abspath('.')
#211和985数据
P211and985UFilePath = curospath + '\\data\\P211_985Universities.xlsx'
#MBTI性格测试数据
MBTIDataFilePath = curospath + '\\data\\MBTI_Data.xlsx'
#学科和专业对照数据
SubMajFilePath = curospath + '\\data\\subjectandmajor.xlsx'
#自主招生信息数据
IndependentEnrollmentInfoFilePath = curospath + '\\data\\IndependentEnrollmentInfos.xlsx'
    
#获取筛选条件
def GetStuConditions(StuInfo):
    '''
    功能：对学生输入的信息进行分析，得到其设置的条件
    '''
    #Conditions = {}.fromkeys(['artsorsci','schoolrank','schoollocation','schoolcategory'])
    Conditions = {'artsorsci':0,'schoolrank':[],'schoollocation':[],'schoolcategory':[]}
    #判断学生文科或理科
    Conditions['artsorsci'] = 0#默认理科
    if '文科' in StuInfo['artsorsci']:
        Conditions['artsorsci'] = 1
    #先判断211或985
    #Conditions['schoolrank'] = 0#先设置一个默认值，默认所有高校
    if '211' in StuInfo['intentuniverrank']:
        Conditions['schoolrank'].append(211)
    if '985' in StuInfo['intentuniverrank']:
        Conditions['schoolrank'].append(985)
    if '普通' in StuInfo['intentuniverrank']:
        Conditions['schoolrank'].append(1)
    if len(Conditions['schoolrank']) == 0:
        Conditions['schoolrank'].append(0)
    #判断意向省份
    #中国省份列表
    Provences = ['北京','广东','山东','江苏','河南','上海','河北','浙江','香港',\
        '陕西','湖南','重庆','福建','天津','云南','四川','广西','安徽','海南',\
        '江西','湖北','山西','辽宁','台湾','黑龙江','内蒙古','澳门','贵州','甘肃'\
        '青海','新疆','西藏','吉林','宁夏']
    for i in range(len(Provences)):
        if Provences[i] in StuInfo['intentunivlocation']:
            Conditions['schoollocation'].append(Provences[i])
            
    #判断类别，优先级要大于地点
    #学校类别列表
    Category = ['全部','综合','工科','农业','林业','医药','师范','语言','财经','政治','体育','艺术','民族']
    for i in range(len(Category)):
        if Category[i] in StuInfo['intentunivercategory']:
            Conditions['schoolcategory'].append(Category[i])
            
    return Conditions
    
#根据筛选条件选择学校并遍历其报考条件
def GetUniversitiesAndApplyConditionsFitTheLimitations(LimitConditions):
    '''
    功能：根据输入的条件选择学校
    '''
    #读取211和985高校信息
    flag,U211List,U985List = rsi.ReadP211And985Universities(P211and985UFilePath)
    if flag == 0:
        return []
    #读取自主招生学校列表
    flag,IndEnrollUniversityDataArray,UniversityPlanAndExamTimeDataArray,ComGradeDataArray,\
        PaperAndPatentDataArray,SubjectCompetitionDataArray,ArtsCompetitionDataArray,\
        SciTechInnovationDataArray = rsi.ReadIndependentEnrollmentInfos(IndependentEnrollmentInfoFilePath)
    if flag == 0:
        return []
    #先对地点进行筛选
    Uposi = []
    for i in range(len(IndEnrollUniversityDataArray)):
        if IndEnrollUniversityDataArray[i]['schoollocation'] in LimitConditions['schoollocation']:
            Uposi.append(i)

    #对学校类别进行筛选
    if not '全部' in LimitConditions['schoolcategory']:#如果没有全部，就进行筛选
        for i in range(len(Uposi)-1,-1,-1):
            if not IndEnrollUniversityDataArray[Uposi[i]]['schoolcategory'] in LimitConditions['schoolcategory']:
                del Uposi[i]

    #对学校层次进行筛选
    if not 0 in LimitConditions['schoolrank']:#只有不是所有学校的时候才进行筛选
        for i in range(len(Uposi)-1,-1,-1):
            if not 1 in LimitConditions['schoolrank']:#如果不含普通，则把普通高校去掉
                if '普通' in IndEnrollUniversityDataArray[Uposi[i]]['schoolrank']:
                    del Uposi[i]
            if not 985 in LimitConditions['schoolrank']:#如果不含985高校
                if '985' in IndEnrollUniversityDataArray[Uposi[i]]['schoolrank']:# in U985List:
                    del Uposi[i]
            if not 211 in LimitConditions['schoolrank']:#如果不含211高校
                if '211' in IndEnrollUniversityDataArray[Uposi[i]]['schoolrank']:# in U211List:
                    del Uposi[i]
    #遍历报考条件
    ApplyConditionsList = []
    for i in Uposi:
        #查找计划和时间
        findflag = 0
        for j in range(len(UniversityPlanAndExamTimeDataArray)):
            if UniversityPlanAndExamTimeDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                limit = UniversityPlanAndExamTimeDataArray[j]['schoollimit']
                plan = UniversityPlanAndExamTimeDataArray[j]['plan']
                apptime = UniversityPlanAndExamTimeDataArray[j]['applytime']
                postdead = UniversityPlanAndExamTimeDataArray[j]['materialpostdeadline']
                examtime = UniversityPlanAndExamTimeDataArray[j]['examtime']
                findflag = 1
                break
        if findflag == 0:
            limit = ''
            plan = ''
            apptime = ''
            postdead = ''
            examtime = ''
            
        #查找综合成绩
        findflag = 0
        for j in range(len(ComGradeDataArray)):
            if ComGradeDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                comgradcond = ComGradeDataArray[j]['comgradcondition']
                findflag = 1
                break
        if findflag == 0:
            comgradcond = ''
        #查找期刊论文
        findflag = 0
        for j in range(len(PaperAndPatentDataArray)):
            if PaperAndPatentDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                if LimitConditions['artsorsci'] == 0:
                    paper = PaperAndPatentDataArray[j]['paper']
                    patent = PaperAndPatentDataArray[j]['patent']
                else:
                    paper = PaperAndPatentDataArray[j]['paper']#文科只给出文章要求
                    patent = ''
                findflag = 1
                break
        if findflag == 0:
            paper = ''
            patent = ''
        #查找五大学科竞赛
        findflag = 0
        if LimitConditions['artsorsci'] == 0:#理科生
            for j in range(len(SubjectCompetitionDataArray)):
                if SubjectCompetitionDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                    math = SubjectCompetitionDataArray[j]['math']
                    phis = SubjectCompetitionDataArray[j]['physics']
                    chem = SubjectCompetitionDataArray[j]['chemistry']
                    bio = SubjectCompetitionDataArray[j]['biology']
                    info = SubjectCompetitionDataArray[j]['infomatics']
                    findflag = 1
                    break
        else:
            findflag = 0
            
        if findflag == 0:
            math = ''
            phis = ''
            chem = ''
            bio = ''
            info = ''
            
        #查找文科竞赛
        findflag = 0
        if LimitConditions['artsorsci'] == 1:#文科生
            for j in range(len(ArtsCompetitionDataArray)):
                if ArtsCompetitionDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                    newcon = ArtsCompetitionDataArray[j]['newconceptcomposition']
                    innocom = ArtsCompetitionDataArray[j]['innovativecomposition']
                    chinesenews = ArtsCompetitionDataArray[j]['chinesenewspapercup']
                    yeshengtao = ArtsCompetitionDataArray[j]['yeshengtaocup']
                    peking = ArtsCompetitionDataArray[j]['pekinguculture']
                    innoeng = ArtsCompetitionDataArray[j]['innovativeenglish']
                    engcom = ArtsCompetitionDataArray[j]['englishcompetition']
                    findflag = 1
                    break
        else:
            findflag = 0
            
        if findflag == 0:
            newcon = ''
            innocom = ''
            chinesenews = ''
            yeshengtao = ''
            peking = ''
            innoeng = ''
            engcom = ''
            
        #查找科技创新类竞赛
        findflag = 0
        for j in range(len(SciTechInnovationDataArray)):
            if SciTechInnovationDataArray[j]['schoolname'] == IndEnrollUniversityDataArray[i]['schoolname']:
                young = SciTechInnovationDataArray[j]['youngsciinnocom']
                tomm = SciTechInnovationDataArray[j]['tomorrowscientist']
                prim = SciTechInnovationDataArray[j]['primsecdschcomputer']
                intsci = SciTechInnovationDataArray[j]['intenationalscieng']
                intenv = SciTechInnovationDataArray[j]['intenationalenvironmentsciprj']
                findflag = 1
                break
        if findflag == 0:
            young = ''
            tomm = ''
            prim = ''
            intsci = ''
            intenv = ''
        
        univ = IndEnrollUniversityDataArray[i]['schoolname']
        loca = IndEnrollUniversityDataArray[i]['schoollocation']
        rank = IndEnrollUniversityDataArray[i]['schoolrank']
        cate = IndEnrollUniversityDataArray[i]['schoolcategory']
        #至此查到了所有的报考条件，现在讲报考条件添加到信息列表中
        data = (univ,loca,rank,cate,limit,plan,apptime,postdead,examtime,comgradcond,paper,patent,\
            math,phis,chem,bio,info,newcon,innocom,chinesenews,\
            yeshengtao,peking,innoeng,engcom,young,tomm,prim,intsci,intenv)
        ApplyConditionsList.append(data)
    UniversityAndApplyConditionsDataArray = np.array(ApplyConditionsList,dtype = IndEnrollUniversityAndApplyConditionsDataType)
    return UniversityAndApplyConditionsDataArray
    
    
def GetAllStudentsUniversityandApplyConditions(FilePath,iStu):
    '''
    功能：根据输入的文件路径获取所有学生的学校和申请条件
    '''
    flag,AllStuInfo = rsi.ReadStuInfoDataFromExcel(FilePath)
    flag2,SubMajData = rsi.ReadSubjectAndMajor(SubMajFilePath)
    flag3,MBTIData = rsi.ReadMBTIDataFromExcel(MBTIDataFilePath)
    
    if flag == 0 or flag2 == 0:
        print u'读取学生信息时出现问题'
    else:
        stunum = len(AllStuInfo)
        #专业
        Subject = ['语文','英语','政治','历史','地理','化学','数学','物理','生物','艺术','体育','信息技术']#学科
        MBTI = ['ENFJ','ENFP','ENTJ','ENTP','ESFJ','ESFP','ESTJ','ESTP',\
            'INFJ','INFP','INTJ','INTP','ISFJ','ISFP','ISTJ','ISTP']
        AllStuUniversityAndApplyconditionsList = []
        SuggestMajorList = []
        StuMbtiTypei = []
        
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
            
        for i in LoopList:
            #判断专业
            SuggestMajor = ''
            for j in range(len(Subject)):
                if Subject[j] in AllStuInfo[i]['advantagesubject']:
                    SuggestMajor += SubMajData[j]['major'] + '\n'
            SuggestMajorList.append(SuggestMajor)
            #判断MBTI类型
            type = -1
            for j in range(len(MBTI)):
                if MBTI[j] in AllStuInfo[i]['mbtitype']:
                    type = j
            StuMbtiTypei.append(type)
            
            limitconditions = GetStuConditions(AllStuInfo[i])
            UACD = GetUniversitiesAndApplyConditionsFitTheLimitations(limitconditions)

            AllStuUniversityAndApplyconditionsList.append(UACD)
        
        return AllStuInfo,SuggestMajorList,StuMbtiTypei,AllStuUniversityAndApplyconditionsList
    
#FilePath = r'C:\Users\CXJ-PC\Desktop\data\stuinfo.xlsx'
#GetAllStudentsUniversityandApplyConditions(FilePath)
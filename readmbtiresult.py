#-*- coding: utf - 8 -*-
import numpy as np
import xlrd
import os

#MBTI性格测试数据格式
MBTIDescriptionDataType = np.dtype({'names': ['type','characterdescription','possibleblindspot','howtouse',\
                                'advantage','disadvantage','suitablepost','suitableprofession','suggestions'],
                            'formats': ['S300', 'S10000','S10000','S10000',\
                                'S10000','S10000','S10000','S10000','S10000',]}, align = True)
                          
#学生信息数据格式
StuInfoDataType = np.dtype({'names': ['name','gender','school', 'grade' , 'phoneno' , 'email'],
                            'formats': ['S30', 'S10', 'S50', 'S30' , 'S20' , 'S50']}, align = True)

                            

def ReadStuMBTITestResult(FilePath):
    '''
    功能：从Excel中读取MBTI测试客户数据
    '''
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0,0
        
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
    AnswerList = []
    for i in range(2,irow):
        name = table.cell(i,2).value.encode('utf-8')#姓名
        gender = table.cell(i,3).value.encode('utf-8')#性别
        school = table.cell(i,4).value.encode('utf-8')#学校
        grade = table.cell(i,5).value.encode('utf-8')#年级
        phoneno = table.cell(i,6).value.encode('utf-8')#电话
        email = table.cell(i,7).value.encode('utf-8')#邮箱
    
        stinfo = (name,gender,school,grade,phoneno,email)
        StuInfoList.append(stinfo)#信息添加到学生信息结构体中
        
        #下面读取答案数据
        answer = []
        for j in range(8,101):#总共93道题目
            answerstr = table.cell(i,j).value.encode('utf-8')
            if 'A' in answerstr:
                answer.append('A')
            if 'B' in answerstr:
                answer.append('B')
        AnswerList.append(answer)
    StuInfoDataArray = np.array(StuInfoList,dtype = StuInfoDataType)
    
    return 1,StuInfoDataArray,AnswerList
    
def ReadMBTIDescriptionData(FilePath):
    '''
    功能：读取MBTI各种类型的分析信息
    '''
    if not os.path.exists(FilePath):#如果文件不存在
        return 0,0
        
    data = xlrd.open_workbook(FilePath)
    table = data.sheet_by_index(0)
    
    MbtiDataList = []
    #读取内容
    for i in range(1,17):
        type = table.cell(i,1).value.encode('utf-8')
        characterdescription = table.cell(i,2).value.encode('utf-8')
        possibleblindspot = table.cell(i,3).value.encode('utf-8')
        howtouse = table.cell(i,4).value.encode('utf-8')
        advantage = table.cell(i,5).value.encode('utf-8')
        disadvantage = table.cell(i,6).value.encode('utf-8')
        suitablepost = table.cell(i,7).value.encode('utf-8')
        suitableprofession = table.cell(i,8).value.encode('utf-8')
        suggestions = table.cell(i,9).value.encode('utf-8')

        mbtid = (type,characterdescription,possibleblindspot,howtouse,advantage,disadvantage,suitablepost,suitableprofession,suggestions)

        MbtiDataList.append(mbtid)

    MBTIDataArray = np.array(MbtiDataList,dtype = MBTIDescriptionDataType)
    return 1,MBTIDataArray

def GetMBTIType(answer):
    '''
    功能：根据学生答案判断所属的MBTI性格类型
    '''
    EPA = [3,7,10,19,23,32,62,74,79,81,83]#选A让E加分的选项
    EPB = [13,16,26,38,42,57,68,77,85,91]#选B让E加分的选项
    SPA = [2,9,25,30,34,39,50,52,54,60,63,73,92]#选A让S加分的选项
    SPB = [5,11,18,22,27,44,46,48,65,67,69,71,82]#选B让S加分的选项
    TPA = [31,33,35,43,45,47,49,56,58,61,66,75,87]#选A让T加分的选项
    TPB = [6,15,21,29,37,40,51,53,70,72,89]#选B让T加分的选项
    JPA = [1,4,12,14,20,28,36,41,64,76,86]#选A让J加分的选项
    JPB = [8,17,24,55,59,78,80,84,88,90,93]#选B让J加分的选项
    
    EScore = 0
    SScore = 0
    TScore = 0
    JScore = 0
    #计算E和I得分
    for i in EPA:
        if answer[i-1] == 'A':
            EScore += 1
    for i in EPB:
        if answer[i-1] == 'B':
            EScore += 1
    IScore = 21 - EScore
    
    #计算S和N得分
    for i in SPA:
        if answer[i-1] == 'A':
            SScore += 1
    for i in SPB:
        if answer[i-1] == 'B':
            SScore += 1
    NScore = 26 - SScore
    
    #计算T和F得分
    for i in TPA:
        if answer[i-1] == 'A':
            TScore += 1
    for i in TPB:
        if answer[i-1] == 'B':
            TScore += 1
    FScore = 24 - TScore
    
    #计算J和P得分
    for i in JPA:
        if answer[i-1] == 'A':
            JScore += 1
    for i in JPB:
        if answer[i-1] == 'B':
            JScore += 1
    PScore = 22 - JScore
    #MBTI得分
    MBTIScore = [EScore,IScore,SScore,NScore,TScore,FScore,JScore,PScore]
    #根据得到判断MBTI类型
    MBTIType = ''
    if EScore >= IScore:
        MBTIType += 'E'
    else:
        MBTIType += 'I'
        
    if SScore >= NScore:
        MBTIType += 'S'
    else:
        MBTIType += 'N'
        
    if TScore >= FScore:
        MBTIType += 'T'
    else:
        MBTIType += 'F'
        
    if JScore >= PScore:
        MBTIType += 'J'
    else:
        MBTIType += 'P'
    #判断MBTI类型的程度
    MBTITypeLevel = []
    if 13 >= EScore >= 11:#外向
        MBTITypeLevel.append('轻微外向')
    if 16 >= EScore >= 14:
        MBTITypeLevel.append('中等外向')
    if 19 >= EScore >= 17:
        MBTITypeLevel.append('明显外向')
    if 21 >= EScore >= 20:
        MBTITypeLevel.append('绝对外向')
        
    if 13 >= IScore >= 11:#内向
        MBTITypeLevel.append('轻微内向')
    if 16 >= IScore >= 14:
        MBTITypeLevel.append('中等内向')
    if 19 >= IScore >= 17:
        MBTITypeLevel.append('明显内向')
    if 21 >= IScore >= 20:
        MBTITypeLevel.append('绝对内向')
        
    if 15 >= SScore >= 13:#感觉
        MBTITypeLevel.append('轻微感觉')
    if 20 >= SScore >= 16:
        MBTITypeLevel.append('中等感觉')
    if 24 >= SScore >= 21:
        MBTITypeLevel.append('明显感觉')
    if 26 >= SScore >= 25:
        MBTITypeLevel.append('绝对感觉')
        
    if 15 >= NScore >= 13:#直觉
        MBTITypeLevel.append('轻微直觉')
    if 20 >= NScore >= 16:
        MBTITypeLevel.append('中等直觉')
    if 24 >= NScore >= 21:
        MBTITypeLevel.append('明显直觉')
    if 26 >= NScore >= 25:
        MBTITypeLevel.append('绝对直觉')
        
    if 14 >= TScore >= 12:#思考
        MBTITypeLevel.append('轻微思考')
    if 18 >= TScore >= 15:
        MBTITypeLevel.append('中等思考')
    if 22 >= TScore >= 19:
        MBTITypeLevel.append('明显思考')
    if 24 >= TScore >= 23:
        MBTITypeLevel.append('绝对思考')
        
    if 14 >= FScore >= 12:#情感
        MBTITypeLevel.append('轻微情感')
    if 18 >= FScore >= 15:
        MBTITypeLevel.append('中等情感')
    if 22 >= FScore >= 19:
        MBTITypeLevel.append('明显情感')
    if 24 >= FScore >= 23:
        MBTITypeLevel.append('绝对情感')
        
    if 13 >= JScore >= 11:#判断
        MBTITypeLevel.append('轻微判断')
    if 16 >= JScore >= 14:
        MBTITypeLevel.append('中等判断')
    if 20 >= JScore >= 17:
        MBTITypeLevel.append('明显判断')
    if 22 >= JScore >= 21:
        MBTITypeLevel.append('绝对判断')
        
    if 13 >= PScore >= 11:#知觉
        MBTITypeLevel.append('轻微知觉')
    if 16 >= PScore >= 14:
        MBTITypeLevel.append('中等知觉')
    if 20 >= PScore >= 17:
        MBTITypeLevel.append('明显知觉')
    if 22 >= PScore >= 21:
        MBTITypeLevel.append('绝对知觉')
    
    return MBTIType,MBTITypeLevel,MBTIScore
    
    
'''    
FilePath = r'E:\xfedu_workstation\MBTITestDoc\mbtiresult.xlsx'
flag,StuInfo,Answer = ReadMBTIResult(FilePath)

stunum = len(StuInfo)
for i in range(stunum-1):
    print StuInfo[i]['name'],StuInfo[i]['school']
    if len(Answer[i]) == 93:
        MBTIType,MBTITypeLevel,MBTIScore = GetMBTIType(Answer[i])
        print MBTIType
        for x in range(4):
            print x,MBTITypeLevel[x].decode('utf-8')
        for s in range(8):
            print s,MBTIScore[s]
    #for j in range(len(Answer[0])):
        #print Answer[i][j]
'''

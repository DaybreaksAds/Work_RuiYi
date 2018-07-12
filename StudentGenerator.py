import xlwt
import xlrd
from xlutils.copy import copy
import random
import numpy as np

def readTestMeanScore(QuestNo):
    ExcelFile=xlrd.open_workbook(r'D:\TestMeanScore.xlsx')
    sheet=ExcelFile.sheet_by_index(0)
    ScoreList=sheet.row_values(QuestNo[2]*3-1)
    return ScoreList

def readTestFullScore(QuestNo):
    ExcelFile=xlrd.open_workbook(r'D:\TestMeanScore.xlsx')
    sheet=ExcelFile.sheet_by_index(0)
    ScoreList=sheet.row_values(QuestNo[2]*3-2)
    return ScoreList

def getTitle(QuestNo):
    ExcelFile=xlrd.open_workbook(r'D:\Title.xlsx')
    sheet=ExcelFile.sheet_by_index(0)
    TitleList=[sheet.row_values(QuestNo[2]*3-3),sheet.row_values(QuestNo[2]*3-2),]
    return TitleList

if __name__=='__main__':
    workbook=xlwt.Workbook(encoding='utf-8')

    TestName=('YuWen_L','ShuXue_L','YingYu_L','ShengWu_L','HuaXue_L','WuLi_L','YuWen_W','ShuXue_W','YingYu_W','DiLi_W','ZhengZhi_W','LiShi_W')
    #TestName=(questNo_keguan,questNo_zhuguan,sequenceInExcel)
    Test_QuestNo={'YuWen_L':(13,10,1),'ShuXue_L':(12,7,2),'YingYu_L':(60,3,3),'ShengWu_L':(6,5,4),'HuaXue_L':(7,4,5),'WuLi_L':(8,4,6),
                  'YuWen_W':(13,10,7),'ShuXue_W':(12,7,8),'YingYu_W':(60,3,9),'DiLi_W':(11,5,10),'ZhengZhi_W':(12,4,11),'LiShi_W':(12,4,12)}

    StudentNum_L=1018
    StudentNum_W=500

    #i为考试列表里第i个考试
    for i in range(len(Test_QuestNo)):
        #第i个考试各个考题的平均值
        MeanScoreList=readTestMeanScore(Test_QuestNo[TestName[i]])
        #第i个考试各个考题的满分
        FullScoreList=readTestFullScore(Test_QuestNo[TestName[i]])
        #第i个考试的列标
        Title=getTitle(Test_QuestNo[TestName[i]])
        worksheet=workbook.add_sheet(TestName[i])

        #在Excel里打印列标
        for ix in range(len(Title[0])):
            worksheet.write(0,ix,Title[0][ix])
            worksheet.write(1,ix,Title[1][ix])

        ClassNum_L=18
        ClassNum_W=9

        if TestName[i][-1]=='L':
            StudentNum=StudentNum_L
            #BaseNum为学号的起始,ClassSize为一个班的人数
            BaseNum=30000
            ClassSize=50
        else:
            StudentNum=StudentNum_W
            BaseNum=32000
            ClassSize=50

        #首先生成学生信息,ix为学生信息的列数
        for ixx in range(2,StudentNum+2):
            x=int(ixx/ClassSize)+1
            y=ixx%int((StudentNum/ClassSize))
            worksheet.write(ixx,0,'高三年级')
            worksheet.write(ixx,1,'行政%s班'%x)
            worksheet.write(ixx,2,'教学%s班'%y)
            worksheet.write(ixx,3,ixx+BaseNum)
            worksheet.write(ixx,4,ixx+BaseNum)

        #j为第i个考试里的第j题，Test_QuestNo[TestName[i]][0]代表客观题
        for j in range(1,Test_QuestNo[TestName[i]][0]+Test_QuestNo[TestName[i]][1]+1):
            MeanScore=MeanScoreList[j]
            FullScore=FullScoreList[j]
            SumScore=0;

            if j<=Test_QuestNo[TestName[i]][0]:
                #k代表第k个学生在第i个考试的第j题的表现
                for k in range(StudentNum):
                    #生成简单均匀分布的分数
                    StudentScore=random.randint(0,FullScore*100)

                    if StudentScore<MeanScore*100:
                        FinalScore=FullScore
                    else:
                        FinalScore=0
                    SumScore=SumScore+FinalScore
                    #学生信息占了5列，再加姓名共六列
                    worksheet.write(k+2,j+5,FinalScore)

            #else情况代表主观题
            else:
                #k代表第k个学生在第i个考试的第j题的表现
                for k in range(StudentNum):
                    #生成正太分布的分数，正太分布的得分就为最后得分
                    if FullScore-MeanScore>MeanScore:
                        StudentScore=int(np.random.normal(MeanScore, MeanScore/4))
                    else:
                        StudentScore=int(np.random.normal(MeanScore, (FullScore-MeanScore)/4))

                    '''
                    #对学生分数进行修正
                    if k<StudentNum/10:
                        StudentScore=StudentScore+random.uniform(0,(FullScore-StudentScore))
                    elif k<StudentNum/7:
                        StudentScore=StudentScore+random.uniform(0,(FullScore-StudentScore)*0.5)
                    elif k>StudentNum*7/8:
                        StudentScore=StudentScore-random.uniform(0,StudentScore)
                    '''

                    FinalScore=int(StudentScore)
                    SumScore=SumScore+FinalScore
                    #学生信息占了5列，再加姓名共六列
                    worksheet.write(k+2,j+5,FinalScore)
            worksheet.write(StudentNum+2,j+5,SumScore/StudentNum)

    workbook.save('D:\My_Excel.xls')

    #对生成的分数进行计分，从原来的excel中读取分数，然后新生成一个excel
    ExcelFile=xlrd.open_workbook(r'D:\My_Excel.xls')
    New_ExcelFile=copy(ExcelFile)

    for i in range(len(Test_QuestNo)):
        sheet=ExcelFile.sheet_by_index(i)
        New_sheet=New_ExcelFile.get_sheet(i)

        ObjectiveNum=Test_QuestNo[TestName[i]][0]
        SubjectiveNum=Test_QuestNo[TestName[i]][1]

        if TestName[i][-1]=='L':
            StudentNum=StudentNum_L
        else:
            StudentNum=StudentNum_W
        for j in range(StudentNum+1):
            Row=sheet.row_values(j+2)
            Sum_Objective=0
            Sum_Subjective=0
            for k in range(6,6+ObjectiveNum):
                Sum_Objective=Sum_Objective+Row[k]

            for k in range(6+ObjectiveNum,6+ObjectiveNum+SubjectiveNum):
                Sum_Subjective=Sum_Subjective+Row[k]

            New_sheet.write(j+2,6+ObjectiveNum+SubjectiveNum,Sum_Objective)
            New_sheet.write(j+2,7+ObjectiveNum+SubjectiveNum,Sum_Subjective)
            New_sheet.write(j+2,8+ObjectiveNum+SubjectiveNum,Sum_Objective+Sum_Subjective)

    New_ExcelFile.save('D:\My_Excel_Final.xls')

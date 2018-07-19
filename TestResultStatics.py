import xlrd
import xlwt

ExcelFile=xlrd.open_workbook(r'D:\StudentResult.xls')
WorkBook=xlwt.Workbook(encoding='utf-8')

FullScoreList=[150,110,150,100]
SubjectNameList=['理科语文','理科物理','文科语文','文科历史']
ChooseClass='J'

for i in range(4):
    sheet=ExcelFile.sheet_by_index(i)

    ClassList_X=sheet.col_values(1)[2:]
    ClassList_J=sheet.col_values(2)[2:]
    NameList=sheet.col_values(3)[2:]
    ScoreList=sheet.col_values(4)[2:]

    #选择行政班或教学班
    if ChooseClass=='X':
        ClassList=ClassList_X
        ClassLogic=1
    else:
        ClassList=ClassList_J
        ClassLogic=0

    ScoreList_1=[]

    '''
    #去除平均数那一行
    ClassList_X.pop()
    ClassList_J.pop()
    NameList.pop()
    ScoreList.pop()
    '''

    StudentNum=len(NameList)-1

    #理科班人数
    StudentNum_L=[0 for i in range(33)]
    #文科班人数
    StudentNum_W=[0 for i in range(33)]
    #文理
    StudentNum_LW=[StudentNum_L,StudentNum_W]

    #A,B,C,D分别代表年级的优秀，良好，及格，低分人数
    ANum=0;
    BNum=0;
    CNum=0;
    DNum=0;

    #A,B,C,D分别代表年级的优秀，良好，及格，低分人数
    ANum_1=0;
    BNum_1=0;
    CNum_1=0;
    DNum_1=0;

    #理科班的平均分
    AverageScore_L=[0 for i in range(33)]
    #文科班的平均分
    AverageScore_W=[0 for i in range(33)]
    #文理
    AverageScore_LW=[AverageScore_L,AverageScore_W]

    #樊子仡，胥明成，李雪霜，吕蕊含的班级排名
    RankInClass_Fan=0
    RankInClass_Xu=0
    RankInClass_Li=0
    RankInClass_Lv=0

    #樊子仡，胥明成，李雪霜，吕蕊含的年级排名
    RankInGrade_Fan=0
    RankInGrade_Xu=0
    RankInGrade_Li=0
    RankInGrade_Lv=0

    #教学班理科从二班开始，行政班时ClassLogic为1，教学班时为0，理科时i>1为0，文科为1
    #           ClassLogic         i<2                结果         逻辑运算结果
    #理科行政     1                  0             从1班开始            1
    #文科行政     1                  1             从1班开始            1
    #理科教学     0                  0             从2班开始            0
    #文科教学     0                  1             从1班开始            1
    ii=int(ClassLogic|(i>1))
    #减法后变为0010，对应下标
    ii=1-ii

    #ii为下标，匹配文字时加一
    for j in range(StudentNum):
        #计算优秀，良好，及格，低分人数
        if ScoreList[j]<FullScoreList[i]*0.8:
            #ANum计算出小于0.8的有多少人
            ANum=ANum+1
            #加入是1班的就加一
            ANum_1=ANum_1+1*(ClassList[j][2:4]=='0'+str(ii+1))
        if ScoreList[j]<FullScoreList[i]*0.75:
            #Num计算出小于0.75的有多少人
            BNum=BNum+1
            BNum_1=BNum_1+1*(ClassList[j][2:4]=='0'+str(ii+1))
        if ScoreList[j]<FullScoreList[i]*0.6:
            #ANum计算出小于0.6的有多少人
            CNum=CNum+1
            CNum_1=CNum_1+1*(ClassList[j][2:4]=='0'+str(ii+1))
        if ScoreList[j]<FullScoreList[i]*0.45:
            #ANum计算出小于0.45的有多少人
            DNum=DNum+1
            DNum_1=DNum_1+1*(ClassList[j][2:4]=='0'+str(ii+1))

        #计算各班人数
        StudentNum_LW[i<2][int(ClassList[j][2:4])-1]=StudentNum_LW[i<2][int(ClassList[j][2:4])-1]+1
        AverageScore_LW[i<2][int(ClassList[j][2:4])-1]=AverageScore_LW[i<2][int(ClassList[j][2:4])-1]+ScoreList[j]

        #+1代表要匹配字符串中的班号
        if int(ClassList[j][2:4])==ii+1:
            ScoreList_1.append(ScoreList[j])

        if NameList[j]=='樊子仡':
            RankInClass_Fan=StudentNum_LW[i<2][int(ClassList[j][2:4])-1]
            RankInGrade_Fan=j+1

        if NameList[j]=='胥明成':
            RankInClass_Xu=StudentNum_LW[i<2][int(ClassList[j][2:4])-1]
            RankInGrade_Xu=j+1

        if NameList[j]=='李雪霜':
            RankInClass_Li=StudentNum_LW[i<2][int(ClassList[j][2:4])-1]
            RankInGrade_Li=j+1

        if NameList[j]=='吕蕊含':
            RankInClass_Lv=StudentNum_LW[i<2][int(ClassList[j][2:4])-1]
            RankInGrade_Lv=j+1

    #将各班总分化为平均分
    for k in range(len(StudentNum_L)):
        if StudentNum_LW[i<2][k]!=0:
            AverageScore_LW[i<2][k]=AverageScore_LW[i<2][k]/StudentNum_LW[i<2][k]

    ANum=StudentNum-ANum
    BNum=StudentNum-BNum-ANum
    CNum=StudentNum-CNum-ANum-BNum

    ANum_1=StudentNum_LW[i<2][ii]-ANum_1
    BNum_1=StudentNum_LW[i<2][ii]-BNum_1-ANum_1
    CNum_1=StudentNum_LW[i<2][ii]-CNum_1-ANum_1-BNum_1

    print(SubjectNameList[i])

    print('总人数：%d,最高分：%d，最低分：%d，平均分：%.3f'%(StudentNum,ScoreList[0],ScoreList[-2],sum(ScoreList[:-1])/StudentNum))
    print('优秀人数：%d，优秀率：%.4f，良好人数：%d，良好率：%.4f，及格人数：%d，及格率：%.4f，低分人数：%d，低分率：%.4f'
          %(ANum,ANum/StudentNum,BNum,BNum/StudentNum,CNum,CNum/StudentNum,DNum,DNum/StudentNum))

    print('1班总人数：%d,最高分：%d，最低分：%d，平均分：%.3f'%(StudentNum_LW[i<2][ii],ScoreList_1[0],ScoreList_1[-1],AverageScore_LW[i<2][ii]))
    print('优秀人数：%d，优秀率：%.4f，良好人数：%d，良好率：%.4f，及格人数：%d，及格率：%.4f，低分人数：%d，低分率：%.4f'
          %(ANum_1,ANum_1/StudentNum_LW[i<2][ii],BNum_1,BNum_1/StudentNum_LW[i<2][ii],CNum_1,CNum_1/StudentNum_LW[i<2][ii],DNum_1,DNum_1/StudentNum_LW[i<2][ii]))

    #1班的排名
    Rank_1=1
    for k in AverageScore_LW[i<2]:
        if AverageScore_LW[i<2][ii]<k:
            Rank_1+=1

    print('一班排名：%d'%Rank_1)

    if i<2:
        print('樊子仡年级排名：%d，樊子仡班级排名：%d，胥明成年级排名：%d，胥明成班级排名：%d'%(RankInGrade_Fan,RankInClass_Fan,RankInGrade_Xu,RankInClass_Xu))
    else:
        print('李雪霜年级排名：%d，李雪霜班级排名：%d，吕蕊含年级排名：%d，吕蕊含班级排名：%d'%(RankInGrade_Li,RankInClass_Li,RankInGrade_Lv,RankInClass_Lv))

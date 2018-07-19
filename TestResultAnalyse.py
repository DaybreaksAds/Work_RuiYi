import pymysql

Conn=pymysql.connect(host="127.0.0.1",port=3306,user="root",passwd="0723",db="schooltestresult",charset="utf8")
Cur=Conn.cursor()
Subject=['理科语文','理科物理','文科语文','文科历史']

def GradeRanking():
    for i in range(4):
        GradeScore=[]
        for j in [1,2]:
            Sql='select avg(score) from school%d_1_%d'%(j,i+1)
            Cur.execute(Sql)
            GradeScore.append(Cur.fetchall())
        print(Subject[i])
        if GradeScore[0]>GradeScore[1]:
            print('学校1排名：1，平均分：%.3f\n学校2排名：2，平均分：%.3f\n'%(GradeScore[0][0][0],GradeScore[1][0][0]))
        else:
            print('学校2排名：1，平均分：%.3f\n学校1排名：2，平均分：%.3f\n'%(GradeScore[0][0][0],GradeScore[1][0][0]))

def Ranking(IniRank,Score,ScoreList,):
    for i in ScoreList:
        if Score<i[0]:
            IniRank+=1
    return IniRank

def ClassRanking(ClassKind):
    if ClassKind=='X':
        ClassLogic=0
        ClassName='ClassNo_X'
    else:
        ClassLogic=1
        ClassName='ClassNo_J'

    for i in range(4):
        ClassRank=1

        #计算第一个班的班号
        #           ClassLogic         i<2                结果         逻辑运算结果
        #理科行政       0                1              从1班开始            0
        #文科行政       0                0              从1班开始            0
        #理科教学       1                1              从2班开始            1
        #文科教学       1                0              从1班开始            0
        ii=int(ClassLogic&int(i<2))+1

        Sql='select avg(score) from School1_1_%d where %s=%d group by %s'%(i+1,ClassName,ii,ClassName)
        Cur.execute(Sql)
        ClassScore_1=Cur.fetchall()[0][0]

        Sql='select avg(score) from School1_1_%d group by %s'%(i+1,ClassName)
        Cur.execute(Sql)
        ClassScore=Cur.fetchall()
        ClassRank=Ranking(ClassRank,ClassScore_1,ClassScore)

        Sql='select avg(score) from School1_2_%d group by %s'%(i+1,ClassName)
        Cur.execute(Sql)
        ClassScore=Cur.fetchall()
        ClassRank_2=Ranking(1,ClassScore_1,ClassScore)
        if ClassRank>ClassRank_2:
            Str='进步'
        else:
            Str='退步'
        print('%s：%d班第一次考试排名：%d，第二次考试排名：%d，比上次%s%d名'%(Subject[i],ii,ClassRank,ClassRank_2,Str,abs(ClassRank-ClassRank_2)))

        Sql='select avg(score) from School2_1_%d group by %s'%(i+1,ClassName)
        Cur.execute(Sql)
        ClassScore=Cur.fetchall()
        ClassRank=Ranking(ClassRank,ClassScore_1,ClassScore)
        print('%s：%d班第一次考试在联校中排名：%d'%(Subject[i],ii,ClassRank))

def StudentRanking(ClassKind,Major):
    if ClassKind=='X':
        ClassName='ClassNo_X'
    else:
        ClassName='ClassNo_J'

    if Major=='L':
        NameList=['樊子仡','胥明成']
        Test=[1,2]
    else:
        NameList=['李雪霜','吕蕊含']
        Test=[3,4]

    for i in NameList:
        for j in Test:
            StudentRank=1

            #找出学生的成绩
            Sql='select score from school1_1_%d where name=\'%s\''%(j,i)
            Cur.execute(Sql)
            Score=Cur.fetchall()[0][0]

            #找出学生所在班级的成绩
            Sql='select score from school1_1_%d where %s=(select %s from school1_1_%d where name=\'%s\')'%(j,ClassName,ClassName,j,i)
            Cur.execute(Sql)
            ClassScore=Cur.fetchall()

            Sql='select score from school1_2_%d where %s=(select %s from school1_2_%d where name=\'%s\')'%(j,ClassName,ClassName,j,i)
            Cur.execute(Sql)
            ClassScore_2=Cur.fetchall()

            #找出学生所在年级的成绩
            Sql='select score from school1_1_%d'%j
            Cur.execute(Sql)
            GradeScore=Cur.fetchall()

            Sql='select score from school1_2_%d'%j
            Cur.execute(Sql)
            GradeScore_2=Cur.fetchall()


            #找出学生所在联校的成绩
            Sql='select score from school2_1_%d'%j
            Cur.execute(Sql)
            MultiSchoolScore=Cur.fetchall()

            StudentRank=Ranking(StudentRank,Score,ClassScore)
            StudentRank_Grade=Ranking(1,Score,GradeScore)
            StudnetRank_MultiSchool=Ranking(StudentRank_Grade,Score,MultiSchoolScore)

            StudentRank_2=Ranking(StudentRank,Score,ClassScore_2)
            StudentRank_Grade_2=Ranking(1,Score,GradeScore_2)

            if StudentRank>StudentRank_2:
                Str1='进步'
            else:
                Str1='退步'

            if StudentRank_Grade>StudentRank_Grade_2:
                Str2='进步'
            else:
                Str2='退步'
            print(i,Subject[j-1]+'第一次班级排名为:',StudentRank,'第二次排名为：',StudentRank_2,Str1,abs(StudentRank-StudentRank_2),'名','班级人数为：',len(ClassScore),
                  '第一次年级排名为:',StudentRank_Grade,'第二次排名为：',StudentRank_Grade_2,Str2,abs(StudentRank_Grade-StudentRank_Grade_2),'名','年级人数为：',len(GradeScore),
                  '第一次联校排名为：',StudnetRank_MultiSchool,'联校人数为：',len(GradeScore)+len(MultiSchoolScore))

def MaxMinAvgScore(ClassKind):
    if ClassKind=='X':
        ClassName='ClassNo_X'
    else:
        ClassName='ClassNo_J'

    for i in range(4):
        print(Subject[i])

        Sql='select avg(score) from School1_1_%d'%(i+1)
        Cur.execute(Sql)
        GradeScore_1=Cur.fetchall()[0][0]
        Sql='select avg(score) from School2_1_%d'%(i+1)
        Cur.execute(Sql)
        GradeScore_2=Cur.fetchall()[0][0]
        AvgGradeScore=(GradeScore_1+GradeScore_2)/2
        if GradeScore_1>GradeScore_2:
            print('%s联校年级最高分：%.3f，联校年级最低分：%.3f，联校年级平均分：%.3f'%(Subject[i],GradeScore_1,GradeScore_2,AvgGradeScore))
        else:
            print('%s联校年级最高分：%.3f，联校年级最低分：%.3f，联校年级平均分：%.3f'%(Subject[i],GradeScore_2,GradeScore_1,AvgGradeScore))

        ScoreList=[]
        Sql='select avg(score) from school1_1_%d group by %s'%(i+1,ClassName)
        Cur.execute(Sql)
        ClassScore=Cur.fetchall()
        for j in ClassScore:
            ScoreList.append(j[0])


        print('%s年级班级最高分：%.3f，年级班级最低分：%.3f，年级班级平均分：%.3f'%(Subject[i],max(ScoreList),min(ScoreList),sum(ScoreList)/len(ScoreList)))

        Sql='select avg(score) from school2_1_%d group by %s'%(i+1,ClassName)
        Cur.execute(Sql)
        ClassScore=Cur.fetchall()
        for j in ClassScore:
            ScoreList.append(j[0])
        print('%s联校班级最高分：%.3f，联校班级最低分：%.3f，联校班级平均分：%.3f'%(Subject[i],max(ScoreList),min(ScoreList),sum(ScoreList)/len(ScoreList)))

if __name__=='__main__':
    #GradeRanking()
    GradeRanking()
    print('行政班：')
    ClassRanking('X')
    StudentRanking('X','L')
    StudentRanking('X','W')
    MaxMinAvgScore('X')
    print('教学班：')
    ClassRanking('J')
    StudentRanking('J','L')
    StudentRanking('J','W')
    MaxMinAvgScore('J')

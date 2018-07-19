import pymysql
import xlrd
import xlwt

def WriteDataToMysql(SchoolNum,TestNum):
    #按照传入的学校号与考试号打开excel
    ExcelFile=xlrd.open_workbook(r'D:\School%d_%d.xls'%(SchoolNum,TestNum))
    #与mysql建立连接
    Conn=pymysql.connect(host="127.0.0.1",port=3306,user="root",passwd="0723",db="schooltestresult",charset="utf8")
    Cur=Conn.cursor()
    Major=['L','W']

    for i in range(4):
        Sheet=ExcelFile.sheet_by_index(i)

        StudentNum=len(Sheet.col_values(0)[2:])
        for j in range(2,StudentNum+2):
            Rows=Sheet.row_values(j)
            if len(Rows[0])==5:
                ClassNo_X=int(Rows[0][2:4])
            else:
                ClassNo_X=int(Rows[0][2])
            if len(Rows[1])==5:
                ClassNo_J=int(Rows[1][2:4])
            else:
                ClassNo_J=int(Rows[1][2])

            Sql='insert into school%d_%d_%d values(%d,%d,%d,\'%s\',%.1f,\'%s\')'%(SchoolNum,TestNum,i+1,j-1,ClassNo_X,ClassNo_J,Rows[2],Rows[3],Major[i>2])
            Cur.execute(Sql)
            Conn.commit()

if __name__=='__main__':
    WriteDataToMysql(2,1)

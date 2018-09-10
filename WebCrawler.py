import urllib.request
from bs4 import BeautifulSoup
import json
import pymysql
import time

#数据库中的主键
id=1

#与数据库建立连接
Conn=pymysql.connect(host="127.0.0.1",port=3306,user="root",passwd="0723",db="schooldata",charset="utf8")
Cur=Conn.cursor()
#添加headers，用谷歌浏览器访问后按F12后直接复制network里的headers
headers={
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/67.0.3396.99 Safari/537.36',
    'Accept':'*/*',
    'Accept-Encoding':'gzip, deflate, br',
    'Accept-Language':'zh-CN,zh;q=0.9',
    'Connection':'keep-alive',
    'Cookie':'tool_ipuse=1.119.33.82; tool_ipprovince=11; tool_iparea=1',
    'Host':'data-gkcx.eol.cn',
    'Referer':'https://gkcx.eol.cn/soudaxue/querySpecialtyScore.html?&studentprovince=&fstype=%E7%90%86%E7%A7%91&keyWord1=&schoolyear=2017&page=3',
}

#记个时
StartTime=time.time()
Time_100=StartTime
for i in range(1,91196):
    #URL直接找JQuery那个
    URL='https://data-gkcx.eol.cn/soudaxue/querySpecialtyScore.html?messtype=jsonp&callback=jQuery18303078462773970765_1533519022238&provinceforschool=&schooltype=&page='+str(i)+'&size=10&keyWord=&schoolproperty=&schoolflag=&province=&fstype=&zhaoshengpici=&fsyear=&zytype=&_=1533519022330'

    req = urllib.request.Request(URL,None,headers)
    WeatherPage=urllib.request.urlopen(req)
    WeatherResult=BeautifulSoup(WeatherPage.read(),'html.parser')

    #处理一下解析出来的
    Result=str(WeatherResult)[41:]
    Result=Result[:-2]

    #掐头去尾后变成Json
    jd=json.loads(Result)

    for j in jd['school']:
        #格式化字符串作为SQL语句
        SQL='insert into data2 values(%d,%s,\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\',\'%s\')'%(id,j['schoolid'],j['schoolname'],j['specialtyname'],j['localprovince'],j['studenttype'],j['year'],j['batch'],j['var_score'],j['max'],j['min'])
        #主键+1
        id+=1
        #执行SQL语句
        Cur.execute(SQL)
        Conn.commit()

    #算个时间
    if i%100==0:
        EndTime=time.time()
        print(i,'次，最近一百次时间为：',EndTime-Time_100,'总时间为：',EndTime-StartTime)
        Time_100=time.time()

# -*- coding: utf-8 -*-
"""
Created on Sat Jun 30 17:30:13 2018

@author: bianl
"""

#  New teachers' Performance estimate
import datetime
import xlrd
from openpyxl import Workbook

def CalNTP(name,Index):  
    start = Index[0]
    end = Index[1]
    month_rate = Index[2]
    clsdict = {}  # 表示老师所授所有课程，及其在其中所贡献业绩的classes dictionary
    #  读入excel   
    for i in range(start, end):
        r = table.row_values(i)   
        #  教师姓名约束
        if name not in r[5]:
            continue
        #  排除已取消的课程
        if "取消" in r[2]:
            continue    
        #  排除自习课、辅导课、模考等
        if len(r[5]) == 0:
            continue         
        #  计算该班的业绩
        if isinstance(r[11],float) :
            if r[11] > 0:
                fee = r[4] * r[11]  # 学费乘以人数
            else: 
                if r[7] == "6人":
                    fee = r[4] * 6
                else: 
                    fee = r[4]
        else:
            fee = 0
        #  计算classes dictionary      
        if r[1] not in clsdict.keys():  # 首次记录该课号
            clsdict[r[1]] = (1/r[10]) * fee
        else:
            clsdict[r[1]] += (1/r[10]) * fee
    sum = 0
    for item in clsdict.keys():
        sum += clsdict[item]
    #  计算月平均业绩
    #  sum为0，或month_rate为负表示仍在培训期内，返回0
    return round(max(0,sum/month_rate),2)

def GetIndex(obdate,enddate,mindate,maxdate):
    #  startdate：enddate表示新教师入职后的测试时间段
    #  mindate：maxdate表示班表中的时间范围
    #  (s,e)[min,max] 或 [min,max](s,e)的情况
    if enddate < mindate or obdate > maxdate:
        return 
    #  [min (s,e) max]的情况
    elif obdate > mindate and enddate < maxdate:
        for i in range(2,rows - 1):  
            curdate = xlrd.xldate_as_datetime(table.row_values(i) [3],0)
            predate = xlrd.xldate_as_datetime(table.row_values(i-1) [3],0)
            if  curdate == obdate:
                start = i
                p = curdate
            if predate <= enddate and curdate > enddate:
                end = i-1   
                q = predate
                break
    #  (s ==[min, e) max]的情况
    elif obdate <= mindate and enddate < maxdate:
        start = 1
        p = mindate
        for i in range(2,rows - 1):  
            curdate = xlrd.xldate_as_datetime(table.row_values(i) [3],0)
            predate = xlrd.xldate_as_datetime(table.row_values(i-1) [3],0)
            if predate <= enddate and curdate > enddate:
                end = i-1
                q = predate
                break
    #  [min (s, max] == e)的情况
    elif obdate > mindate and enddate >= maxdate:  
        end = rows - 1
        q = maxdate              
        for i in range(2,rows - 1):      
            curdate = xlrd.xldate_as_datetime(table.row_values(i) [3],0)
            if  curdate == obdate:
                start = i
                p = curdate
                break  
    #  p和q表示最终计算的日期范围
    #  受培训期影响，p通常比obdate晚3个月
    #  month_rate为[p,q]期间工作的月份
    day = (q-p).days
    month_rate = day / 30
    return [start,end,month_rate]
        
duration = input("新教师入职x个月的产能表现，请输入x值(Deafault x = 9)：")
#  Set the default vaule of duration
if len(duration) < 1:
    duration = 12  
#  打开excel
data1 = xlrd.open_workbook("国外部教师名单6.15.xlsx")  
nametable = data1.sheets()[0]
m = nametable.nrows 
perform = {}

data2 = xlrd.open_workbook("配课表明细16.6.1-18.5.31.xlsx")  
table = data2.sheets()[0]
rows = table.nrows 
#  转换excel日期到datetime格式，并找到table中的最小和最大日期
mindate = xlrd.xldate_as_datetime(table.row_values(1) [3],0)  
maxdate = xlrd.xldate_as_datetime(table.row_values(rows - 1) [3],0)
for i in range(1, m-1):
    r = nametable.row_values(i) 
    name = r[2]  #teacher's name      
    obdate = xlrd.xldate_as_datetime(r[3],0)  #  Onboarding date
    enddate = obdate + datetime.timedelta(days = (int(duration) * 30))
    Index = GetIndex(obdate,enddate,mindate,maxdate)
    if Index is None:
        continue
    tmp = CalNTP(name,Index)
    perform[name] = tmp

#  输出结果文件：“New_Teacher_Performance.xlsx”
#  对perform的value排序
output = sorted(perform.items(),key = lambda item:item[1], reverse = True)
#  在内存中创建一个workbook对象，而自动创建一个worksheet   
wb = Workbook()   
#  获取当前活跃的worksheet，默认就是第一个worksheet
ws = wb.active
for i in range(len(perform)):
    ws.append(output[i])
wb.save("New_Teacher_Performance.xlsx")          
           



    
    
    

    

    
    
    
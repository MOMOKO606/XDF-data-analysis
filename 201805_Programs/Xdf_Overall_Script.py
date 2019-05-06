# -*- coding: utf-8 -*-
"""
Created on Fri Jun 15 00:15:47 2018

@author: bianl
"""
def isinClassType(Classtype, Classprogram, CurrentClass, StudentNum):
    """
    Input
    Classtype: dict{课程项目名(str):[课程包含关键字](list)}
    Classprogram: 课程项目名(str)
    CurrentClass: 当前课程名(str)
    StudentNum: 学生人数(str)
    
    Output
    True if 当前课程CurrentClass 属于 课程项目Classprogram; False otherwise
    """
    flag = False
    if "VIP" in Classprogram:  # VIP课一定是1对1 或 6人
        if StudentNum != "1对1" and StudentNum != "6人":
            return False
    if "班级" in Classprogram:  # 班级项目一定不是 1对1 或 6人的
        if StudentNum == "1对1" or StudentNum == "6人":
            return False              
    for i in range(len(Classtype[Classprogram])):
        if Classtype[Classprogram][i] in CurrentClass:
            flag = True
    return flag


#  根据班表，计算一段时间内，某部门（或全部），某课程（或全部）的产能，与各教师在其中的产能列表
#  输出2个txt文档：Summary.txt 和 Classes_table.txt
#  输出1个excel文档：Teacher_List.xlsx
import datetime
import xlrd
from openpyxl import Workbook

#  打开excel
data = xlrd.open_workbook("配课表明细16.6.1-18.5.31.xlsx")  
table = data.sheets()[0]
rows = table.nrows 
cols = table.ncols 
#  转换excel日期到datetime格式，并找到table中的最小和最大日期
mindate = xlrd.xldate_as_datetime(table.row_values(1) [3],0)  
maxdate = xlrd.xldate_as_datetime(table.row_values(rows - 1) [3],0)  
#  根据用户输入的时间范围，找出table中的index
while True:
    st = input("Please enter start date (e.g. 2015-12-21): ")
    ft = input("Please enter end date (e.g. 2015-12-21): ")
    
    if len(st) < 1:  # Default value, 从第1行开始（第0行为header）
        start = 1
    else:
        try:
            stmp = datetime.datetime.strptime(st, "%Y-%m-%d")  # str转datetime
            if stmp <= mindate:
                start = 1
            else:
                 for i in range(1,rows - 1):
                     curdate = xlrd.xldate_as_datetime(table.row_values(i) [3],0)
                     if  curdate == stmp:
                         start = i
                         break                
        except:
            print("Please re-enter start date in this 2015-12-21 format. ")
            continue
            
    if len(ft) < 1:  # Default value, 在最后一行结束（Python下标从0开始，故最后一行index比excel小1）
        end = rows - 1
        break
    else:
        try:
            etmp = datetime.datetime.strptime(ft, "%Y-%m-%d")
            if etmp >= maxdate:
                end = rows - 1
            else:
                for i in range(max(2,start),rows - 1):  # i至少从2开始
                    curdate = xlrd.xldate_as_datetime(table.row_values(i) [3],0)
                    predate = xlrd.xldate_as_datetime(table.row_values(i-1) [3],0)
                    if predate <= etmp and curdate > etmp:
                        end = i-1
                        break
            break
        except:
            print("Please re-enter end date in this 2015-12-21 format. ")
            continue


#  输入要查询的项目部门
#  Default为全部部门
DeptType = ["北美项目部", "英联邦项目部"]
while True:
    Dep = input("Please input department: '北美项目部', '英联邦项目部' or both(press Enter directly) ")   
    if Dep in DeptType or len(Dep) < 1:
        break
    else:
        print("Please re-enter department name")
                
#  输入要查询的课程类型
#  提示所有的课程类型
#  Default为全部类型
ClassType = {"英联邦VIP":["雅思","IELTS"],"雅思班级":["雅思","IELTS"],
             "托福VIP":["TOEFL"],"托福班级":["TOEFL"],
             "美研":["GRE","GMAT"],"美本":["SSAT","SAT","ACT","AP","北美本科精英"],
             "其他":["企业培训班","英联邦中学","国际英语实验班","TOEIC","海外留学菁英"]}        
while True:
    print("共有7种课程项目，分别为：")  
    print([keys for keys in ClassType.keys()],end = " ")   
    # Class Type constraints
    cyconsts = input("Please input class type: (press Enter directly = select all the type) ")  
    if cyconsts in ClassType.keys() or len(cyconsts) < 1:
        break
    else:
        print("Please re-enter class type name")
 
resclstab = []  # Result Class Table   
classes = {}  # 用于计算的classes dictionary
#  读入excel   
header = table.row_values(0)  # 第一行为表头
for i in range(start, end):
    r = table.row_values(i) 
    curdep = r[0]  # 当前部门
    curcls = r[2]  # 当前课程种类       
    #  Department Constraint
    if len(Dep) >= 1:  # 没有使用default的情况
        if Dep not in curdep:
            continue
    #  Class Type Constraint
    if len(cyconsts) >=1:  # 没有使用default的情况
        flag = isinClassType(ClassType, cyconsts, curcls, r[7])
        if  not flag:
            continue      
    #  排除已取消的课程
    if "取消" in r[2]:
        continue    
    #  排除自习课、辅导课、模考等
    if len(r[5]) == 0:
        continue       
    clsnum = r[1]  # 当前课号
    cur_t = r[5]  # 当前教师
    if clsnum not in classes.keys():
        #  当前课号第一次出现
        t4tc = {cur_t:1}  # teacher for this class
        #  更新classes dictionary
        classes[clsnum] = t4tc      
        #  更新Result Class Table   
        opendate = xlrd.xldate_as_datetime(r[8],0)
        closedate = xlrd.xldate_as_datetime(r[9],0)
#       tmp = [curdep,curcls,clsnum,r[4],r[10],opendate,closedate]
        tmp = [curdep,curcls,clsnum,r[4],r[10],r[7],r[11]]
        resclstab.append(tmp)
    else:
        #  当前课号已存在于classes中
        if cur_t in classes[clsnum].keys():  # 当前教师已存在于该课号中
            classes[clsnum][cur_t] += 1
        else:  # 当前教师不存在于该课号中
            classes[clsnum][cur_t] = 1
                
#  修改classes和resclstab  
totalclsn = len(resclstab)  # Total classes number      
totalfee = 0  # 总业绩初值   
teacherlist = {}  #  参与教师dict初值     
for i in range(totalclsn):
    clsnum = resclstab[i][2]
    #  计算该班的业绩
    if isinstance(resclstab[i][6],float) :
        if resclstab[i][6] > 0:
            fee = resclstab[i][3] * resclstab[i][6]  # 学费乘以人数
        else: 
            if resclstab[i][5] == "6人":
                fee = resclstab[i][3] * 6
            else: 
                fee = resclstab[i][3]
    else:
        fee = 0
    tct = resclstab[i][4]  # Total Class Times 
    totalfee += fee
    #  更新该课程中每位老师的课数、课时贡献百分比 和 对应的完成业绩
    for t in classes[clsnum].keys(): 
        if t not in teacherlist.keys():
            teacherlist[t] = 0
        times = classes[clsnum][t]
        p = times / tct
        classes[clsnum][t] = [times,round(p,2),round(p*fee,2)]
    resclstab[i].append(classes[clsnum])
totaltn = len(teacherlist.keys())  # 参与教师总数

#  输出结果文件：“Summary.txt”
if len(st) < 1:
    st = mindate.strftime("%Y-%m-%d")
if len(ft) < 1:
    ft = maxdate.strftime("%Y-%m-%d")   
if len(Dep) < 1:
    Dep = "全部项目部"
if len(cyconsts) < 1:
    cyconsts = "全部"
    
output_text =("您选择了 " + st + " 到 " + ft + "\n"
       + Dep + "\n"
       + cyconsts + "课程项目\n"
       + "\n"
       + "共开班：" + str(totalclsn) + "节\n"
       + "共" + str(totaltn) + "位教师参与教学\n"
       + "完成业绩：" + str(totalfee) + "元\n")
f = open("Summary.txt" , "w")
f.write(output_text)
f.close()

#  输出结果文件：“Classes_Table.txt”
f = open("Classes_Table.txt" , "w+")
for i in range(len(resclstab)):
    f.writelines(str(resclstab[i]) + "\n")
f.close()

#  输出结果文件：“Teacher_List.xlsx”
#  计算教师产值列表
for c in classes.keys():
     for t in classes[c].keys():                       
         individ = classes[c][t][2]  # 个人业绩
         teacherlist[t] += round(individ,2)
#  排序教师产值
#  对teacherlist（dict）的value排序
output = sorted(teacherlist.items(),key = lambda item:item[1], reverse = True)
#  在内存中创建一个workbook对象，而自动创建一个worksheet   
wb = Workbook()   
#  获取当前活跃的worksheet，默认就是第一个worksheet
ws = wb.active
for i in range(totaltn):
    ws.append(output[i])
wb.save("Teacher_List.xlsx")

    


        
        


        





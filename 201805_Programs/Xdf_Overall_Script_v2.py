# -*- coding: utf-8 -*-
"""
Created on Fri Jun 15 00:15:47 2018

@author: bianl
"""


def GetIndex( CLASS_DATE_POS ):
    """
    function: 根据用户输入的起始日期和班表中的日期，计算所需循环数据的角标index。
    :param CLASS_DATE_POS: 常量，表示班表中日期位于第几列。
    :return: 班表中起始到终止数据的index，[start, end]。
             st, ft, mindate, maxdate在输出报告中会用到。
    """

    #  转换excel日期到datetime格式，并找到table中的最小和最大日期。
    #  注意，excel中的日期格式应为2018/10/22形式，其他形式需在excel中先进行预处理。
    mindate = xlrd.xldate_as_datetime(table.row_values(1)[CLASS_DATE_POS], 0)  # 班表中的最早日期。
    maxdate = xlrd.xldate_as_datetime(table.row_values(rows - 1)[CLASS_DATE_POS], 0)  # 班表中的最晚日期。

    #  根据用户输入的时间范围，找出table中的index
    while True:
        #  输入查询的起始与终止时间，不做任何输入直接回车表示使用Default value，即班表中的全部时间。
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
                    for i in range(1, rows - 1):
                        curdate = xlrd.xldate_as_datetime(table.row_values(i)[CLASS_DATE_POS], 0)
                        if curdate == stmp:
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
                    for i in range(max(2, start), rows - 1):  # i至少从2开始
                        curdate = xlrd.xldate_as_datetime(table.row_values(i)[CLASS_DATE_POS], 0)
                        predate = xlrd.xldate_as_datetime(table.row_values(i - 1)[CLASS_DATE_POS], 0)
                        if predate <= etmp and curdate > etmp:
                            end = i - 1
                            break
                break
            except:
                print("Please re-enter end date in this 2015-12-21 format. ")
                continue
    return [start, end, st, ft, mindate, maxdate]


def GetDep( DEPT_TYPE ):
    """
    function: 根据用户输入，确定所选部门。
    :param DEPT_TYPE: 常量，包含所有的部门名称。
    :return: Dep，用户所选部门。
    """

    while True:
        #  输入要查询的项目部门。
        #  不做任何输入直接回车表示使用Default value，即班表中的所有部门。
        Dep = input("Please input department: '北美项目部', '英联邦项目部' or both(press Enter directly) ")
        if Dep in DEPT_TYPE or len(Dep) < 1:
            break
        else:
            print("Please re-enter department name")
    return Dep


def GetClassname( CLASS_TYPE ):
    """
    function: 根据用户输入，确定所选课程类型。
    :param CLASS_TYPE: 常量，包括所有的课程类型。
    :return: cyconsts， 用户所选课程类型。
    """

    while True:
        #  显示课程项目提示信息。
        print("共有7种课程项目，分别为：")
        print([keys for keys in CLASS_TYPE.keys()], end=" ")

        #  Class Type constraints
        #  输入要查询的课程类型。
        #  不做任何输入直接回车表示使用Default value，即班表中的所有课程。
        cyconsts = input("Please input class type: (press Enter directly = select all the type) ")
        if cyconsts in CLASS_TYPE.keys() or len(cyconsts) < 1:
            break
        else:
            print("Please re-enter class type name")
    return cyconsts


def isinClassType(CLASS_TYPE, Classprogram, CurrentClass, StudentNum):
    """
    function: 检查班表中的当前课程是否满足用户所选的约束条件。
    :param
    CLASS_TYPE: 包含所有的课程类型，dict{课程项目名(str):[课程包含关键字](list)}；
    Classprogram: 课程类型名(str)，一般表示用户所选的课程类型；
    CurrentClass: 当前课程名(str)；
    StudentNum: 学生人数(str)；
    :return: True if当前课程CurrentClass，属于（用户选择的）课程类型Classprogram; False otherwise。
    """

    flag = False
    if "VIP" in Classprogram:  # VIP课一定是1对1 或 6人
        if StudentNum != "1对1" and StudentNum != "6人":
            return False
    if "班级" in Classprogram:  # 班级项目一定不是 1对1 或 6人的
        if StudentNum == "1对1" or StudentNum == "6人":
            return False
    for i in range(len(CLASS_TYPE[Classprogram])):
        if CLASS_TYPE[Classprogram][i] in CurrentClass:
            flag = True
    return flag



#  --------------------------- Script --------------------------
#  根据班表，计算一段时间内，某部门（或全部），某课程（或全部）的产能，与各教师在其中的产能列表。
#  输出2个txt文档：Summary.txt 和 Classes_table.txt；
#  输出1个excel文档：Teacher_List.xlsx。

#  适用于201805的数据，配课表明细16.6.1-18.5.31.xlsx，课次 & 总课次格式。
#  相比于Xdf_Overall_Script_v1.py，v2版本：
#  1.增加了注释；
#  2.程序模块化，尽量使用多个函数来实现脚本内容；
#  3.设定了一些常量，这样在班表格式改变时做出简单的修改即可。

import datetime
import xlrd
from openpyxl import Workbook

#  需要用到的一些常量。
CLASS_SCHEDULE_FILE = "配课表明细16.6.1-18.5.31.xlsx"  # excel格式的配课表文件名。
DEPT_TYPE = ["北美项目部", "英联邦项目部"]  #  所有的项目部门。
#  所有的课程类型。
CLASS_TYPE = {"英联邦VIP": ["雅思", "IELTS"], "雅思班级": ["雅思", "IELTS"],
             "托福VIP": ["TOEFL"], "托福班级": ["TOEFL"],
             "美研": ["GRE", "GMAT"], "美本": ["SSAT", "SAT", "ACT", "AP", "北美本科精英"],
             "其他": ["企业培训班", "英联邦中学", "国际英语实验班", "TOEIC", "海外留学菁英"]}
DEPT_POS = 0  #  部门信息位于excel中的第几列（从0开始），例如，D列即为3，第3列。
CLASS_NUM = 1  #  课号信息（班级编码）位于excel中的第几列（从0开始）。
CLASS_POS = 2  # 课程种类信息位于excel中的第几列（从0开始）。
CLASS_DATE_POS = 3  #  上课日期位于excel中的第几列（从0开始）
FEE_POS = 4  #  学费信息位于excel中的第几列（从0开始）。
TEACHERS_POS = 5  #  教师姓名信息位于excel中的第几列（从0开始）。
STU_CAP_POS = 7  #  班级容量信息位于excel中的第几列（从0开始）。
OPENDATE_POS = 8  #  开班日期信息位于excel中的第几列（从0开始）。
CLOSEDATE_POS = 9  #  结课日期信息位于excel中的第几列（从0开始）。
CLASS_TIMES_POS = 10  #  总课次信息位于excel中的第几列（从0开始）。
STU_NUM_POS = 11  #  当前学生数量信息位于excel中的第几列（从0开始）。


# CLASS_SCHEDULE_FILE = "FY19国外配课表.xlsx"  # excel格式的配课表文件名。
# DEPT_TYPE = ["北美项目部", "英联邦项目部"]  #  所有的项目部门。
# #  所有的课程类型。
# CLASS_TYPE = {"英联邦VIP": ["雅思", "IELTS", "GCSE"], "雅思班级": ["雅思", "IELTS"],
#              "托福VIP": ["TOEFL"], "托福班级": ["TOEFL"],
#              "美研": ["GRE", "GMAT"], "美本": ["SSAT", "SAT", "ACT", "AP", "北美本科精英"],
#              "其他": ["企业培训班", "英联邦中学", "国际英语实验班", "TOEIC", "海外留学菁英"]}
# DEPT_POS = 17  #  部门信息位于excel中的第几列（从0开始），例如，D列即为3，第3列。
# CLASS_NUM = 0  #  课号信息（班级编码）位于excel中的第几列（从0开始）。
# CLASS_POS = 1  # 课程种类信息位于excel中的第几列（从0开始）。
# CLASS_DATE_POS = 3  #  上课日期位于excel中的第几列（从0开始）
# FEE_POS = 2  #  学费信息位于excel中的第几列（从0开始）。
# TEACHERS_POS = 11  #  教师姓名信息位于excel中的第几列（从0开始）。
#
# STU_CAP_POS = 7  #  班级容量信息位于excel中的第几列（从0开始）。
# OPENDATE_POS = 8  #  开班日期信息位于excel中的第几列（从0开始）。
# CLOSEDATE_POS = 9  #  结课日期信息位于excel中的第几列（从0开始）。
# CLASS_TIMES_POS = 10  #  总课次信息位于excel中的第几列（从0开始）。
# STU_NUM_POS = 11  #  当前学生数量信息位于excel中的第几列（从0开始）。

#  打开excel。
#  table, rows, cols类似于全局变量，可直接用于function中。
data = xlrd.open_workbook( CLASS_SCHEDULE_FILE )
table = data.sheets()[0]
rows = table.nrows
cols = table.ncols

#  STEP1. 根据用户输入的时间范围，找出table中的index。
[start, end, st, ft, mindate, maxdate] = GetIndex( CLASS_DATE_POS )

#  STEP2. 用户输入要查询的项目部门。
Dep = GetDep( DEPT_TYPE )

#  STEP3. 用户输入要查询的课程类型。
cyconsts = GetClassname( CLASS_TYPE )

#  STEP4. 计算每位教师的产能等。
#  (1).以开班编号为primary key来扫描并统计班表。
#  用resclstab和classes存放计算结果。

#  赋初值：
resclstab = []  # Result Class Table.
#  用于计算的classes dictionary.
#  班级编号为key，value为一个dict{该班级的所有老师名：其对应的上课次数}。
classes = {}
#  读入excel
header = table.row_values(0)  # 第一行为表头。
for i in range(start, end):
    r = table.row_values(i)
    curdep = r[DEPT_POS]  # 当前部门。
    curcls = r[CLASS_POS]  # 当前课程种类。

    #  Sentinels
    #  Department Constraint.
    if len(Dep) >= 1:  # 没有使用default的情况。
        if Dep not in curdep:
            continue
    #  Class Type Constraint.
    if len(cyconsts) >= 1:  # 没有使用default的情况。
        flag = isinClassType(CLASS_TYPE, cyconsts, curcls, r[STU_NUM_POS])
        if not flag:
            continue
            #  排除已取消的课程。
    if "取消" in r[CLASS_POS]:
        continue
        #  排除自习课、辅导课、模考等。
    if len(r[TEACHERS_POS]) == 0:
        continue

    clsnum = r[CLASS_NUM]  # 当前课号。
    cur_t = r[TEACHERS_POS]  # 当前教师。
    #  当前课号第一次出现时：
    if clsnum not in classes.keys():
        t4tc = {cur_t: 1}  # teacher for this class.
        #  更新classes dictionary.
        classes[clsnum] = t4tc
        #  更新Result Class Table.
        opendate = xlrd.xldate_as_datetime(r[OPENDATE_POS], 0)
        closedate = xlrd.xldate_as_datetime(r[CLOSEDATE_POS], 0)
        tmp = [curdep, curcls, clsnum, r[FEE_POS], r[CLASS_TIMES_POS], r[STU_CAP_POS], r[STU_NUM_POS]]
        resclstab.append(tmp)
    else:
        #  当前课号已存在于classes中的情况：
        if cur_t in classes[clsnum].keys():  # 当前教师已存在于该课号中。
            classes[clsnum][cur_t] += 1
        else:  # 当前教师不存在于该课号中。
            classes[clsnum][cur_t] = 1

#  (2).以教师为primary key，整理&修改classes和resclstab。
totalclsn = len(resclstab)  # Total classes number.
totalfee = 0  # 总业绩初值。
teacherlist = {}  # 参与教师dict初值。
for i in range(totalclsn):
    clsnum = resclstab[i][2]
    #  计算该班的业绩。
    #  resclstab[i][6]表示该班级的学生数量。

    #  学生数量不为空，则有可能开班。
    if isinstance(resclstab[i][6], float):
        #  学生数量大于0则按实际人数计算该班级创造的业绩。
        if resclstab[i][6] > 0:
            #  resclstab[i][3]表示每位学生的学费。
            fee = resclstab[i][3] * resclstab[i][6]  # 学费乘以人数。
        #  学生数量为0则按预期班级容量计算该班级创造的（预期）业绩。
        else:
            if resclstab[i][5] == "6人":
                fee = resclstab[i][3] * 6
            else:
                fee = resclstab[i][3]
    #  学生数量为空，则没有开班。
    else:
        fee = 0

    tct = resclstab[i][4]  # Total Class Times.
    #  该班级的业绩加入总业绩。
    totalfee += fee
    #  更新该课程中每位老师的课数、课时贡献百分比 和 对应的完成业绩。
    for t in classes[clsnum].keys():
        #  首次统计到该老师时，将其姓名加入teacherlist。
        if t not in teacherlist.keys():
            teacherlist[t] = 0
        times = classes[clsnum][t]  #  t老师在该班级上了times课。
        p = times / tct  #  t老师在该班级上课所占总课次的比例。
        #  更新classes，存入t老师在该班级上了times课，占所有课次的比例，以及对应该班级学费的业绩。
        classes[clsnum][t] = [times, round(p, 2), round(p * fee, 2)]
    #  更新resclstab中的classes信息。
    resclstab[i].append(classes[clsnum])
# 参与教师总数。
totaltn = len(teacherlist.keys())


#  输出3个结果文件。
#  输出结果文件1：“Summary.txt”.
if len(st) < 1:
    st = mindate.strftime("%Y-%m-%d")
if len(ft) < 1:
    ft = maxdate.strftime("%Y-%m-%d")
if len(Dep) < 1:
    Dep = "全部项目部"
if len(cyconsts) < 1:
    cyconsts = "全部"

output_text = ("您选择了 " + st + " 到 " + ft + "\n"
               + Dep + "\n"
               + cyconsts + "课程项目\n"
               + "\n"
               + "共开班：" + str(totalclsn) + "节\n"
               + "共" + str(totaltn) + "位教师参与教学\n"
               + "完成业绩：" + str(totalfee) + "元\n")
f = open("Summary.txt", "w")
f.write(output_text)
f.close()

#  输出结果文件2：“Classes_Table.txt”.
f = open("Classes_Table.txt", "w+")
for i in range(len(resclstab)):
    f.writelines(str(resclstab[i]) + "\n")
f.close()

#  输出结果文件3：“Teacher_List.xlsx”.
#  计算教师产值列表。
#  按每个班级中的每个教师的产能，叠加到teacherlist中该教师的总产能。
for c in classes.keys():
    for t in classes[c].keys():
        individ = classes[c][t][2]  # 个人业绩。
        teacherlist[t] += round(individ, 2)
#  排序教师产值.
#  对teacherlist（dict）的value排序。
output = sorted(teacherlist.items(), key=lambda item: item[1], reverse=True)
#  在内存中创建一个workbook对象，而自动创建一个worksheet。
wb = Workbook()
#  获取当前活跃的worksheet，默认就是第一个worksheet。
ws = wb.active
for i in range(totaltn):
    ws.append(output[i])
wb.save("Teacher_List.xlsx")














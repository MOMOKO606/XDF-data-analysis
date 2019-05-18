# -*- coding: utf-8 -*-
"""
Created on Fri Jun 15 00:15:47 2018

@author: bianl
"""
def GetDateLimit( CLASS_DATE_POS ):
    """
    function: 班表中的上课日期可能不是排序好的，此时需要遍历班表，获取最小和最大日期。
    :param CLASS_DATE_POS: 上课日期在班表中的位置（列数）。
    :return: mindate & maxdate，即此班表中的日期范围。
    """

    #  班表的行数。
    m = table.nrows
    #  赋初值。
    mindate = xlrd.xldate_as_datetime(table.row_values(1)[CLASS_DATE_POS], 0)
    maxdate = xlrd.xldate_as_datetime(table.row_values(1)[CLASS_DATE_POS], 0)

    #  excel班表为1到m行，读入table后自动转为0到m-1行。
    #  第0行为header，第1行赋初值，所以从第2行循环至第m-1行。
    for i in range(2, m):
        #  每个班上课的日期。
        tmpdate = xlrd.xldate_as_datetime(table.row_values(i)[CLASS_DATE_POS], 0)
        if tmpdate < mindate:
            mindate = tmpdate
        if tmpdate > maxdate:
            maxdate = tmpdate
    return mindate, maxdate


def GetDateRange( CLASS_DATE_POS ):
    """
    function: 根据用户输入的起始日期和班表中的日期，计算所需的日期范围。
    :param CLASS_DATE_POS: 常量，表示班表中日期位于第几列。
    :return: 日期范围startdate, enddate;
            mindate, maxdate为班表中的最小和最大日期范围，在输出报告中会用到。
    """

    #  转换excel日期到datetime格式，并找到table中的最小和最大日期。
    #  注意，excel中的日期格式应为2018/10/22形式，其他形式需在excel中先进行预处理。
    [mindate, maxdate] = GetDateLimit(CLASS_DATE_POS)

    #  根据用户输入的日期范围，确定最终的日期范围。
    while True:
        #  输入查询的起始与终止日期，不做任何输入直接回车表示使用Default value，即班表中的全部日期。
        st = input("Please enter start date (e.g. 2015-12-21): ")
        ft = input("Please enter end date (e.g. 2015-12-21): ")

        if len(st) < 1:  # Default value, 从第1行开始（第0行为header）
            startdate = mindate
        else:
            try:
                stmp = datetime.datetime.strptime(st, "%Y-%m-%d")  # str转datetime
                if stmp <= mindate:
                    startdate = mindate
                else:
                    startdate = stmp
            except:
                print("Please re-enter start date in this 2015-12-21 format. ")
                continue

        if len(ft) < 1:  # Default value, 在最后一行结束（Python下标从0开始，故最后一行index比excel小1）
            enddate = maxdate
            break
        else:
            try:
                etmp = datetime.datetime.strptime(ft, "%Y-%m-%d")
                if etmp >= maxdate:
                    enddate = maxdate
                else:
                    enddate = etmp
                break
            except:
                print("Please re-enter end date in this 2015-12-21 format. ")
                continue
    return [startdate, enddate, mindate, maxdate]

# def GetIndex( CLASS_DATE_POS ):
#     """
#     function: 根据用户输入的起始日期和班表中的日期，计算所需循环数据的角标index。
#     :param CLASS_DATE_POS: 常量，表示班表中日期位于第几列。
#     :return: 班表中起始到终止数据的index，[start, end]。
#              st, ft, mindate, maxdate在输出报告中会用到。
#     """
#
#     #  转换excel日期到datetime格式，并找到table中的最小和最大日期。
#     #  注意，excel中的日期格式应为2018/10/22形式，其他形式需在excel中先进行预处理。
#     # mindate = xlrd.xldate_as_datetime(table.row_values(1)[CLASS_DATE_POS], 0)  # 班表中的最早日期。
#     # maxdate = xlrd.xldate_as_datetime(table.row_values(rows - 1)[CLASS_DATE_POS], 0)  # 班表中的最晚日期。
#     [mindate, maxdate] = GetDateRange(CLASS_DATE_POS)
#
#     #  根据用户输入的时间范围，找出table中的index
#     while True:
#         #  输入查询的起始与终止时间，不做任何输入直接回车表示使用Default value，即班表中的全部时间。
#         st = input("Please enter start date (e.g. 2015-12-21): ")
#         ft = input("Please enter end date (e.g. 2015-12-21): ")
#
#         if len(st) < 1:  # Default value, 从第1行开始（第0行为header）
#             start = 1
#         else:
#             try:
#                 stmp = datetime.datetime.strptime(st, "%Y-%m-%d")  # str转datetime
#                 if stmp <= mindate:
#                     start = 1
#                 else:
#                     for i in range(1, rows - 1):
#                         curdate = xlrd.xldate_as_datetime(table.row_values(i)[CLASS_DATE_POS], 0)
#                         if curdate == stmp:
#                             start = i
#                             break
#             except:
#                 print("Please re-enter start date in this 2015-12-21 format. ")
#                 continue
#
#         if len(ft) < 1:  # Default value, 在最后一行结束（Python下标从0开始，故最后一行index比excel小1）
#             end = rows - 1
#             break
#         else:
#             try:
#                 etmp = datetime.datetime.strptime(ft, "%Y-%m-%d")
#                 if etmp >= maxdate:
#                     end = rows - 1
#                 else:
#                     for i in range(max(2, start), rows - 1):  # i至少从2开始
#                         curdate = xlrd.xldate_as_datetime(table.row_values(i)[CLASS_DATE_POS], 0)
#                         predate = xlrd.xldate_as_datetime(table.row_values(i - 1)[CLASS_DATE_POS], 0)
#                         if predate <= etmp and curdate > etmp:
#                             end = i - 1
#                             break
#                 break
#             except:
#                 print("Please re-enter end date in this 2015-12-21 format. ")
#                 continue
#     return [start, end, st, ft, mindate, maxdate]


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


def GetStuNum(current_num, class_name):
    """
    function: 根据班表中的2个参数，计算该班级的学生数。
    :param current_num:班表中的当前人数，即导出班表的时刻该班级中的学生数。
                         如果此时该班已结课或退课，则为0，否则则显示真实学生数。
    :param class_name:班表中的班级名称，当无法获得真实学生数时，可通过班级名称中的关键字估算学生数。
    :return: 该班级中的学生数量。
    """

    #  如果当前人数不为零，则为真实学生数，直接返回该值。
    if current_num != 0:
        return current_num
    #  当前人数为零时，则需要从班级名称中寻找关键字。
    if "二" in class_name:
        return 2
    elif "三" in class_name:
        return 3
    elif "1" in class_name:
        return 1
    elif "8" in class_name:
        return 8
    elif "15" in class_name:
        return 15
    elif ("基础" or "强化" or "全程") in class_name:
        return 8
    else:  #  "定制" in class_name
        return 15



#  --------------------------- Script --------------------------
#  根据班表，计算一段时间内，某部门（或全部），某课程（或全部）的产能，与各教师在其中的产能列表。
#  输出2个txt文档：Summary.txt 和 Classes_table.txt；
#  输出1个excel文档：Teacher_List.xlsx。

#  适用于201905的数据，FY19国外配课表.xlsx，课时 & 总课时格式。
#  相比于Xdf_Overall_Script_v2.py，v3版本：
#  1.常量值有不同，e.g. 加入了GCSE课程；
#  2.没有课容的数据；
#  3.按照课时（单位分钟），而不是课次（单位次数）计算，注意有总课时为0或空的情况；
#  4.Sentinel中“取消”字眼不代表该班彻底取消，而是已打卡的课程正常收费，后面的课程取消。

import datetime
import xlrd
from openpyxl import Workbook

#  需要用到的一些常量。
CLASS_SCHEDULE_FILE = "FY19国外配课表.xlsx"  # excel格式的配课表文件名。
DEPT_TYPE = ["北美项目部", "英联邦项目部"]  #  所有的项目部门。
#  所有的课程类型。
CLASS_TYPE = {"英联邦VIP": ["雅思", "IELTS", "GCSE"], "雅思班级": ["雅思", "IELTS"],
             "托福VIP": ["TOEFL"], "托福班级": ["TOEFL"],
             "美研": ["GRE", "GMAT"], "美本": ["SSAT", "SAT", "ACT", "AP", "北美本科精英"],
             "其他": ["企业培训班", "英联邦中学", "国际英语实验班", "TOEIC", "海外留学菁英"]}
DEPT_POS = 17  #  部门信息位于excel中的第几列（从0开始），例如，D列即为3，第3列。
CLASS_NUM = 0  #  课号信息（班级编码）位于excel中的第几列（从0开始）。
CLASS_POS = 1  # 课程种类信息位于excel中的第几列（从0开始）。
CLASS_DATE_POS = 3  #  上课日期位于excel中的第几列（从0开始）
FEE_POS = 2  #  学费信息位于excel中的第几列（从0开始）。
TEACHERS_POS = 11  #  教师姓名信息位于excel中的第几列（从0开始）。
OPENDATE_POS = 21  #  开班日期信息位于excel中的第几列（从0开始）。
CLOSEDATE_POS = 22  #  结课日期信息位于excel中的第几列（从0开始）。
CLASS_TIME_POS = 6  #  本次课时（分钟）信息位于excel中的第几列（从0开始）。
CLASS_TIMESUM_POS = 24  #  总课时(分钟)信息位于excel中的第几列（从0开始）。
STU_NUM_POS = 29  #  当前学生数量信息位于excel中的第几列（从0开始）。

#  此新班表中不存在的信息：
#  STU_CAP_POS = 7  #  班级容量信息位于excel中的第几列（从0开始）。
# CLASS_TIMES_POS = 10  #  总课次信息位于excel中的第几列（从0开始）。


#  打开excel。
#  table, rows, cols类似于全局变量，可直接用于function中。
data = xlrd.open_workbook( CLASS_SCHEDULE_FILE )
table = data.sheets()[0]
rows = table.nrows
cols = table.ncols

#  STEP1. 根据用户输入的时间范围，找出table中的index。
[startdate, enddate, mindate, maxdate] = GetDateRange( CLASS_DATE_POS )

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
for i in range(1, rows):
    r = table.row_values(i)
    xlrd.xldate_as_datetime(r[CLASS_DATE_POS], 0)
    curdate = xlrd.xldate_as_datetime(r[CLASS_DATE_POS], 0)  # 当前日期。
    curdep = r[DEPT_POS]  # 当前部门。
    curcls = r[CLASS_POS]  # 当前课程种类。
    #  Sentinels
    #  Date Constraint.
    if curdate < startdate or curdate > enddate:
        continue
    #  Department Constraint.
    if len(Dep) >= 1:  # 没有使用default的情况。
        if Dep not in curdep:
            continue
    #  Class Type Constraint.
    if len(cyconsts) >= 1:  # 没有使用default的情况。
        flag = isinClassType(CLASS_TYPE, cyconsts, curcls, r[STU_NUM_POS])
        if not flag:
            continue
    #  排除自习课、辅导课、模考等。
    if len(r[TEACHERS_POS]) == 0:
        continue
    #  排除没学费的课程。
    if r[FEE_POS] == "#N/A":
        continue

    curclstime = float(r[CLASS_TIME_POS])  # 当前班级所用学时（单位：分钟）。
    totalclstime = r[CLASS_TIMESUM_POS]  # 该班级的总学时（单位：分钟）。
    clsnum = r[CLASS_NUM]  # 当前课号。
    cur_t = r[TEACHERS_POS]  # 当前教师。
    cur_stunum = GetStuNum(int(r[STU_NUM_POS]), r[CLASS_POS])

    #  当前课号第一次出现时：
    if clsnum not in classes.keys():
        t4tc = {cur_t: curclstime}  # teacher for this class.
        #  更新classes dictionary.
        classes[clsnum] = t4tc

        #  更新Result Class Table.
        #  将课程总时长（str）转换为float。
        try:
            studyhours = float(totalclstime)
            if studyhours == 0:  #  班表中总时长为0的情况。
                studyhours = None
        except:
            studyhours = None  #  班表中总时长为空的情况。

        tmp = [curdep, curcls, clsnum, float(r[FEE_POS]), studyhours, cur_stunum]
        resclstab.append(tmp)
    else:
        #  当前课号已存在于classes中的情况：
        if cur_t in classes[clsnum].keys():  # 当前教师已存在于该课号中。
            classes[clsnum][cur_t] += curclstime
        else:  # 当前教师不存在于该课号中。
            classes[clsnum][cur_t] = curclstime

#  (2).以教师为primary key，整理&修改classes和resclstab。
totalclsn = len(resclstab)  # Total classes number.
totalfee = 0  # 总业绩初值。
teacherlist = {}  # 参与教师dict初值。

for i in range(totalclsn):
    clsnum = resclstab[i][2]
    #  计算该班的业绩。
    #  resclstab[i][3]表示该班级每位学生的学费。
    #  resclstab[i][4]表示该班级的总时长（单位：分钟）。
    #  resclstab[i][5]表示该班级的学生数量。

    fee = resclstab[i][3] * resclstab[i][5]
    #  该班级的业绩加入总业绩。
    totalfee += fee
    tct = resclstab[i][4]  # Total Class Time.

    #  不定时课程的情况，则将所有老师的课时叠加作为总课时。
    if tct is None:
        tct = 0
        for t in classes[clsnum].keys():
            tct += classes[clsnum][t]

    #  更新该课程中每位老师的课数、课时贡献百分比 和 对应的完成业绩。
    for t in classes[clsnum].keys():
        #  首次统计到该老师时，将其姓名加入teacherlist。
        if t not in teacherlist.keys():
            teacherlist[t] = 0
        time = classes[clsnum][t]  #  t老师在该班级上了time分钟的课。
        p = time / tct
        #  更新classes，存入t老师在该班级上了times课，占所有课次的比例，以及对应该班级学费的业绩。
        classes[clsnum][t] = [time, round(p, 2), round(p * fee, 2)]
    #  更新resclstab中的classes信息。
    resclstab[i].append(classes[clsnum])
# 参与教师总数。
totaltn = len(teacherlist.keys())


#  输出3个结果文件。
#  输出结果文件1：“Summary.txt”.
if len(Dep) < 1:
    Dep = "全部项目部"
if len(cyconsts) < 1:
    cyconsts = "全部"

#  datetime转字符串。
st = startdate.strftime("%Y-%m-%d %H:%M:%S")
ft = enddate.strftime("%Y-%m-%d %H:%M:%S")

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














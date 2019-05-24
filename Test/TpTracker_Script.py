import datetime
import os
import xlrd
from TpTracker import TpTrackers
from openpyxl import Workbook
from openpyxl import load_workbook



def CalcuAll( CLASS_SCHEDULE_FILE, output_mode ):
    """
    function: 试算2018.06.01-2019.05-31，所有部门，所有老师的产能。
    :param CLASS_SCHEDULE_FILE: 班表名称。
    :return: 输出2个txt, 1个excel文档。
    """

    #  定义类函数所需的parameters。
    st = str()
    ft = str()
    Dep = str()
    cyconsts = str()
    teachers = str()

    #  创建Teacher's performance Trackers实例。
    fy19 = TpTrackers( CLASS_SCHEDULE_FILE )
    #  计算Teacher's performance。
    return fy19.FilterClsTable(st, ft, Dep, cyconsts, teachers, output_mode)


def PerfomafterEntry( CLASS_SCHEDULE_FILE, ENTRYTIME_FILE, timewindow, timelimit ):
    """
    function: 试算每位老师入职后timewindow时间内的产能。
    :param CLASS_SCHEDULE_FILE: 班表名称。
    :param ENTRYTIME_FILE: 老师入职时间表。
    :param timewindow: 要计算的入职后的时间范围（单位：月）。
    :param timelimit: 要计算的开始带课后的时间范围（单位：月）。
    :return entryperfom: 每位老师入职后timewindow时间内的产能列表。
            [教师姓名，总产能，总课时，产能/课时，产能/月，首节课日期, 末节课日期, 带课时间（月），入职时间]
    """

    #  打开入职时间表。
    data1 = xlrd.open_workbook( ENTRYTIME_FILE )
    nametable = data1.sheets()[0]
    m = nametable.nrows

    #  定义常量。
    TEACHER_POS = 2  #  教师姓名在excel中的第2列（从第0列开始）。
    ENTRY_DATE_POS = 3  #  入职日期在excel中的第2列（从第0列开始）。
    NOW_DATE = datetime.datetime.strptime("2019-05-21", "%Y-%m-%d")  #  入职至今，即“今”表示的日期。

    #  创建Teacher's performance Trackers实例。
    fy19 = TpTrackers( CLASS_SCHEDULE_FILE )

    #  定义function所需的其他参数。
    Dep = str()
    cyconsts = str()

    #  对结果赋初值。
    entryperform =[]
    #  对均值赋初值。
    perform_mean = 0
    hours_mean = 0
    pperh_mean = 0
    pperm_mean = 0
    count = 0  # 新教师人数。
    header = nametable.row_values(0)  # 第一行为表头。
    for i in range(1, m):
        r = nametable.row_values(i)
        teachers = r[TEACHER_POS]    # 读取教师姓名。

        #  读取入职时间。
        entry_date = xlrd.xldate_as_datetime(r[ENTRY_DATE_POS], 0)
        end_date = entry_date + datetime.timedelta(days=(int(timewindow) * 30))

        #  datetime转str。
        entry_date_str = entry_date.strftime("%Y-%m-%d")
        end_date_str = end_date.strftime("%Y-%m-%d")

        #  计算Teacher's performance。
        [tmp, detail] = fy19.FilterClsTable(entry_date_str, end_date_str, Dep, cyconsts, teachers, False)

        #  跳过空值的计算结果。
        if len(tmp) < 1:
            continue

        delta = (NOW_DATE - entry_date).days
        #  筛选入职不超过15个月，带课不超过12个月的新老师。
        if delta < ((timewindow - 12) * 31 + 365) and tmp[0][7] <= 12:
            tmp = list(tmp[0])  # 从[[]]修改为[]
            tmp.append(entry_date.strftime("%Y-%m-%d"))
            entryperform.append(tmp)
            count += 1
            perform_mean += tmp[1]
            hours_mean += tmp[2]
            pperh_mean += tmp[3]
            pperm_mean += tmp[4]

    entryperform.append([" ", " ", " ", " ", " ", " ", " ", " ", " "])
    entryperform.append([" ", "平均产能", "平均课时", "平均产能/课时", "平均产能/月", " ", " ", " "])
    entryperform.append([" ", round(perform_mean / count), round(hours_mean / count), round(pperh_mean / count),
               round(pperm_mean / count), " ", " ", " "])

    return entryperform


#------------------------------------------------------------------
#  Script:
#  定义function的parameter。
CLASS_SCHEDULE_FILE = "FY19国外配课表.xlsx"
ENTRYTIME_FILE = "【国外部】教师名单-3.7.xlsx"
timewindow = 15
timelimit = 12
CASE = 1

if CASE == 1:
    #  test1. 试算2018.06.01-2019.05-31，所有部门，所有老师的产能。
    CalcuAll( CLASS_SCHEDULE_FILE, True )
elif CASE == 2:
    # test3. 试算每位老师入职不超过timewindow个月，带课不超过timelimit个月的产能表现。
    entryperform = PerfomafterEntry(CLASS_SCHEDULE_FILE, ENTRYTIME_FILE, timewindow, timelimit)
    #  输出到excel。
    #  在内存中创建一个workbook对象，而自动创建一个worksheet。
    wb = Workbook()
    #  获取当前活跃的worksheet，默认就是第一个worksheet。
    ws = wb.active
    ws.append(["教师姓名","总产能","总课时","产能/课时","产能/月","首节课日期","末节课日期","带课时间（月）","入职日期"])
    for i in range(len(entryperform)):
        ws.append(entryperform[i])
    #  保存excel。
    wb.save("Performance_after_entry.xlsx")



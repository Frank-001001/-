import xlrd
import xlwt
from xlutils.copy import copy
import os
import time
t=time.localtime()
book=""
data = os.listdir(os.getcwd())
for i in range(len(data)):
    if data[i].startswith('北京宏软高科信息技术有限公司_每日统计') == True:
    #if data[i].startswith(f'北京宏软高科信息技术有限公司_每日统计_{t.tm_year}{t.tm_mon}{t.tm_mday}') == True:
        book=data[i]
        break
# 检查字符串是否以指定的字符串开头


# 打开工作簿
open_book = xlrd.open_workbook(filename=book, encoding_override=True)
# 获取选项卡

sheet = open_book.sheet_by_index(0)
# 获取列和行

name_all = sheet.col_values(0, 4)  # 所有人名
# 创建字典
dict = {}
# 长期请假的列表
leave_list = []
# 没全勤的列表
no_attendance = []
# 全勤的列表
attendance_list = []

#正常请假列表
normal_leave=[]


# dict={name_all[i]:{sheet.cell_value(2,x):sheet.col_values(x,4)[i]} for i in range(len(name_all)) for x in range(9,17,2)}
# 获取所有人的打卡时间
for i in range(len(name_all)):
    for x in range(9, 19, 2):
        if name_all[i] in dict:
            dict[name_all[i]].update({sheet.cell_value(2, x): sheet.col_values(x, 4)[i]})
        else:
            dict[name_all[i]] = {
                sheet.cell_value(2, x): sheet.col_values(x, 4)[i]
            }
    for x in range(28,30):
        if sheet.cell_value(i,x) ==" ":
            continue
        elif name_all[i] in normal_leave:
            continue
        elif sheet.cell_value(i,x) =="1":
            normal_leave.append(name_all[i])

dict.pop("高永爱")
dict.pop("陈永杰")
dict.pop("贾老师")
dict.pop("宋朝庚")
dict.pop("陆佳琪")
dict.pop("吴美娟")
dict.pop("傅优")
dict_all = dict.copy()

for x in [x for x in dict]:
    list = [y for i, y in dict[x].items()]
    if list.count("请假") == 5:
        if x in leave_list:
            continue
        else:
            leave_list.append(x)
    elif list.count("缺卡") >= 1 or list.count("早退") >= 1 or list.count("旷工迟到")>=1 :
        if x in no_attendance:
            continue
        else:
            no_attendance.append(x)
    elif list.count("正常")+list.count("外勤")+list.count("补卡审批通过") == 5 or list.count("正常")==5:
        if x in attendance_list:
            continue
        else:
            attendance_list.append(x)
    else:
        no_attendance.append(x)

# for i in leave_list:
#     print("长期请假的有", i)
# for i in no_attendance:
#     print("没有全勤的有", i)
# for i in attendance_list:
#     print("全勤的有", i)
# print("长期", len(leave_list), "no", len(no_attendance), "ok", len(attendance_list))
# print(len(leave_list) + len(no_attendance) + len(attendance_list))

# print("总人数为", len(dict_all))
#写入表格


work = xlrd.open_workbook("人员统计.xlsx")
workbook = xlwt.Workbook(encoding="utf-8")  # 使用xlwt写入excel表格数据
copywb = copy(work)  # 把原来的工作簿拷贝一份
target = copywb.get_sheet(0)  # 获取第一个选项卡
for i in range(3,30,13):
    target.write(i,1,len(dict_all))#在职总人数写入
    target.write(i+1,1,len(leave_list))#长期请假人数写入
    target.write(i+2,1,len(normal_leave))#正常请假写入
    target.write(i+3,1,len(attendance_list))#实到人数
    target.write(i+4,1,len(no_attendance))#写入迟到人数

copywb.save(f"{t.tm_year}-{t.tm_mon}-{t.tm_mday}统计.xlsx")
# sheet.write(34,0,'Hello Word')
#写入表格
print("ok")

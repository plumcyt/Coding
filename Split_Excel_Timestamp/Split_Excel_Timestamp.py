import pandas as pd
import os
import time
import traceback
# import xlrd
import xlsxwriter

f = open("Split_Excel Conf.txt") 
lines = f.readlines()
#file_path = str(lines[0].rstrip()) #第一行是file path
#column = str(lines[1].rstrip()) #第二行是column name

def Split_Excel():
    try:
        raw_path = str(lines[0].rstrip())  #第一行是file path
        if not os.path.exists(raw_path):  # 判断文件是否存在
            print('No such file!')
        else:
            split_x = str(lines[1].rstrip())  #第二行是column name拆分标准
            raw_excel = pd.read_excel(raw_path)  # 读取文件
            header = raw_excel.columns
            if split_x not in header:
                print('No such column name!')
            else:
                timestamp = input('Timestamp Y/N?: ').lower()  # 时间戳
                split_list = list(set(raw_excel[split_x]))  # 提取拆分清单，去重
                row = len(split_list)  # 拆分数量
                end_path = os.path.dirname(raw_path)  # 输出文件夹
                file_date = time.strftime("%Y-%m-%d", time.localtime(time.time()))  # 日期戳
                count = 0
                # 遍历
                for line in split_list:
                    excel_select = raw_excel[raw_excel[split_x] == line]  # 根据拆分元素提取数据
                    if timestamp == "y":
                        writer = pd.ExcelWriter(f'{str(end_path)}\{str(line)} {str(file_date)}.xlsx',
                                                datetime_format="DD/MM/YYYY")  # 生成新文件名带时间戳
                    else:
                        writer = pd.ExcelWriter(f'{str(end_path)}\{str(line)}.xlsx',
                                                datetime_format="DD/MM/YYYY")  # 生成新文件名无时间戳
                    excel_select.to_excel(writer, sheet_name='Sheet1', index=False)  # 数据保存到新文件
                    worksheet = writer.sheets['Sheet1']  # 工作表名字
                    worksheet.set_column("A:Z", 15)  # 列宽15
                    writer.save()
                    count += 1
                    print('\rTotal line(s) {}, splitting {}'.format(row, count), end='')
                print("\n" + "Excel Split by " + '"' + str(split_x) + '"')
    except Exception:
        tb = traceback.format_exc()
        print('Error!!!!!!!!!:\n', tb)


# 默认执行一遍
Split_Excel()

while True:
    repeat = input("Do you want to repeat the split excel script? Y/N: ").lower()
    if repeat == "y":
        Split_Excel()
    else:
        break

os.system("pause")  # 结束后不退出，win os only
# 或者
# input("Press any key to exit!")

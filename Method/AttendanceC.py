# 这是一个考勤排查程序系统
#by relee_xrl
import pandas as pd

'''循环按照员工，拉出来所有该员工打卡信息
                在该员工打卡信息中上班时间循环遍历：
                    如xx天没有打卡：printxx天缺勤
                    如打卡：列出当天所有打卡记录
                        如少于两次，列出仅有的打卡时间
                        如大于等于两次，循环遍历打卡时间
                            如在上午规定9：00前打卡：上午打卡记录+1
                            如在下午规定5：30后打卡：下午打卡记录+1
                            如上午打卡记录=0但下午打卡记录！=0：
                                print上午没有在规定时间打卡
                            如下午打卡记录=0但下午打卡记录！=0：
                                print下午没有在规定时间打卡
                            如上午下午都=0：
                                print上午下午均没有在规定时间打卡       
'''


def readTXTFile(file):
    try:
        global defined_days
        with open(file, 'r') as file:
            content = file.read()
            defined_days = content.split()
            return defined_days
    except FileNotFoundError:
        print("时间表文件未找到，请检查文件路径和名称是否正确。")


def readExcelFile(file):
    try:
        df = pd.read_excel(file)
        return df
    except FileNotFoundError:
        print("考勤表文件未找到，请检查文件路径和名称是否正确。")


def groupedAteendanceByName(df):
    print(" ")
    print("----------------考勤表按照姓名分类--------------------------------")
    print("- 输入'1'显示所有人员并按照姓名分类。")
    print("- 输入'姓名'显示该'姓名'考勤结果。（孙岩，李婧，李新蕊，林华，聂兰勇，黄岩）")
    name = input("- 输入 'q' 退出。")

    try:
        while name != 'q':
            if name == '1':
                employeeGrouped = df.groupby('姓名')
                for employee, employee_info in employeeGrouped:
                    print(employee_info)
                name = input("- 输入 '1' 或 '姓名' 或 'q' 退出。")
            elif name == '孙岩' or '李婧' or '李新蕊' or '林华' or '聂兰勇' or '黄岩':
                filtered_df = df[df['姓名'] == name]
                print(filtered_df)
                name = input("- 输入 '1' 或 '姓名' 或 'q' 退出。")
            elif name == 'q':
                quit()
    except:
        print("Err, try again.")



def AttendanceC(df, defined_days, defined_work_start_time, defined_work_end_time):
    employeeGrouped = df.groupby('姓名')
    for employee, employee_info in employeeGrouped:
        employee_info['时间'] = pd.to_datetime(employee_info['时间'])

        for d in defined_days:
            if pd.to_datetime(d).date() not in employee_info['时间'].dt.date.values:
                print(employee + " " + d + "号缺勤，没有打卡记录。")
            else:
                date_records = employee_info[employee_info['时间'].dt.date == pd.to_datetime(d).date()]

                am_records = 0
                pm_records = 0
                # 遍历每次打卡记录
                for _, record in date_records.iterrows():
                    work_time = record['时间'].time()
                    if work_time < defined_work_start_time:
                        am_records += 1
                    elif work_time > defined_work_end_time:
                        pm_records += 1

                if am_records == 0 and pm_records == 0:
                    print(employee + " " + d + " 上午下午均没有在规定时间打卡")
                    # 输出没有在规定时间打卡的记录时间
                    for _, record in date_records.iterrows():
                        work_time = record['时间'].time()
                        work_date = record['时间'].date()
                        if defined_work_start_time < work_time < defined_work_end_time:
                            print("当天上班时间打卡记录：" + str(work_date) + " " + str(work_time))
                elif am_records == 0 and pm_records != 0:
                    print(employee + " " + d + " 上午没有在规定时间打卡")
                    # 输出没有在规定时间打卡的记录时间
                    for _, record in date_records.iterrows():
                        work_time = record['时间'].time()
                        work_date = record['时间'].date()
                        if defined_work_start_time < work_time < defined_work_end_time:
                            print("当天上班时间打卡记录：" + str(work_date) + " " + str(work_time))
                elif am_records != 0 and pm_records == 0:
                    print(employee + " " + d + " 下午没有在规定时间打卡")
                    # 输出没有在规定时间打卡的记录时间
                    for _, record in date_records.iterrows():
                        work_time = record['时间'].time()
                        work_date = record['时间'].date()
                        if defined_work_start_time < work_time < defined_work_end_time:
                            print("当天上班时间打卡记录：" + str(work_date) + " " + str(work_time))

                # if am_records == 0 and pm_records == 0:
                #     print(employee + " " + d + " :上午下午均没有在规定时间打卡")
                # elif am_records == 0 and pm_records != 0:
                #     print(employee + " " + d + " :上午没有在规定时间打卡")
                # elif am_records != 0 and pm_records == 0:
                #     print(employee + " " + d + " :下午没有在规定时间打卡")


# 输出结果写入文本文件
# output_file = '1.txt'
# with open(output_file, 'w') as file:
#     for entry in output:
#         file.write(entry + '\n')
#
# print(f'输出结果已保存到 "{output_file}" 文件中。')


def main():
    defined_work_start_time = pd.to_datetime('09:00:00').time()
    defined_work_end_time = pd.to_datetime('17:30:00').time()

    try:
        # defined_days = 1.txt 格式['2023-08-25', '2023-08-28', '2023-08-29']
        defined_days_file = input("请输入时间表名(.txt)：")
        defined_days = readTXTFile(defined_days_file)
        # print(defined_days)

        # defined_file = '2.xlsx'
        defined_file_input = input("请输入考勤表名(.xlsx)：")
        df = readExcelFile(defined_file_input)
        # print(df)

        print("----------------考勤问题人员如下---------------------------------")
        print(" ")
        AttendanceC(df, defined_days, defined_work_start_time, defined_work_end_time)
        groupedAteendanceByName(df)
    except:
        print("EXIT!")

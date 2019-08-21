import os
import re
import time

# from library import yaml
from library.R_r_excel import ReadExcel

ALL_PAT = r"(\n\n\d\d\d\d-\d\d-\d\d \d*?:\d\d:\d\d .*?\n打卡\d*?次.*?\n\n)"
CONT_PAT = r"(?<=[\>|\)])\n(.*)"
SIGN_PAT = r"\n打卡(.*)次(.*)"
NICK_PAT = r":\d{2} (.*)[\(|\<]([^\)|\>]*)"
DATE_PAT = r"(\d{4}-\d{2}-\d{2}) (\d{1,2}:\d{1,2}:\d{1,2})"

# YAML_PATH = "./config.yaml"

num_dict = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",
            "十一": "11", "十二": "12", "十三": "13", "十四": "14", "十五": "15"}
TXT_PATH = r"./txt/『联盟群』花落闪暖，国服安一_2.txt"
EXCEL_PATH = "./All_3.xlsx"
ro = ReadExcel(EXCEL_PATH, "Sheet1")
NESS_WORDS = ["打卡", '次']
FORB_WORDS = []
FRESHTIME = "05:00:00"
STARTCOLNO = 3  # 开始记录的列
TITLELINE = 3  # title所在行
LastUpdateTime = ro.read_one_cell(2, 3)

QQ_dict = {i.QQ: i.row for i in ro.r_data_obj_from_column([1, 2], title_line=TITLELINE)}
nick_dict = {i.nickname: i.row for i in ro.r_data_obj_from_column([1, 2], title_line=TITLELINE)}

Origin_Date = ro.read_one_cell(TITLELINE - 1, STARTCOLNO)

Date = str(Origin_Date.year) + "-" + str(Origin_Date.month) + "-" + str(Origin_Date.day)

Trun_Date = lambda string: time.strptime(string, "%Y-%m-%d").tm_yday
ydate = Trun_Date(Date)
miss_list = []
susp_list = []

# with open(YAML_PATH, 'r', encoding='utf-8') as f:
#     conf = yaml.safe_load(f)
#     TXT_NAME = conf['txt']['name']
#     TXT_PATH = conf['txt']['path'] + TXT_NAME
#     EXCEL_NAME = conf['excel']['name']
#     EXCEL_PATH = conf['excel']['path'] + EXCEL_NAME
#     NESS_WORDS = conf['necessaryWords']
#     FORB_WORDS = conf['forbiddenWords']


with open(TXT_PATH, "r", encoding="utf-8") as fo:
    txt_file = fo.read()


def add_zero(strr, num):
    while len(strr) < num:
        strr = "0" + strr
    return strr


LastUpdateTime = str(LastUpdateTime.year) + "-" + add_zero(str(LastUpdateTime.month), 2) + "-" + add_zero(
    str(LastUpdateTime.day), 2) + " " + add_zero(str(LastUpdateTime.hour), 2) + ":" + add_zero(
    str(LastUpdateTime.minute), 2) + ":" + add_zero(str(LastUpdateTime.second), 2)
print(LastUpdateTime)
BeginTime = LastUpdateTime
currentmodifytime = LastUpdateTime


def char_trans(char):
    """
    中文大写的数字数值化
    :param char: string  待处理的数据
    :return: int    对应数据
    """
    if char in num_dict.keys():
        return num_dict[char]
    if char in num_dict.values():
        return char
    raise AttributeError


#    raise AttributeError


def inter_trans(inter):
    """
    数据化
    :param inter: string  待处理数据
    :return: double  计算完后的数据
    """
    inter = inter.lower()
    inter = inter.strip("分")  # 移除句尾的分
    inter = char_trans(inter)
    special_symbol = ["w", "万"]
    for symbol in special_symbol:
        if symbol in inter:
            raw_int = inter.split(symbol)[0]
            raw_int = char_trans(raw_int)
            if "." in raw_int:
                return float(raw_int) * 10000
            raw_float = inter.split(symbol)[1]
            raw_float = char_trans(raw_float)
            return float(raw_int + "." + raw_float) * 10000
    else:
        return float(inter)


def cal_day_offset(time):
    if time > FRESHTIME:
        return 0
    else:
        return -1


def table_init():
    """
    表格初始化
    :return: None
    """
    # 如果表格已存在，则清空除第一行的所有数据
    if os.path.exists(EXCEL_PATH):
        print("读取表格中".center(24, "-"))
        fo = ReadExcel(EXCEL_PATH, "Sheet1")
        return fo
        # 如果表格不存在，新建表格、写表头
    # else:
    #     print("正在新建表格".center(24, "-"))
    #     ReadExcel.create_new_workbook(EXCEL_NAME)
    #     fo = ReadExcel(EXCEL_PATH, "Sheet1")
    #     fo.clear_sheet()
    #     column = 1
    #     titles = ("日期", "时间", "昵称", "单次分数", "总次数", "总分", "原文本", "部分未解析数据")
    #     widths = ("11", "9", "40", "9", "7", "8", "24", "68")
    #     for (title, width) in zip(titles, widths):
    #         fo.w_data(1, column, title)
    #         fo.set_column_width(chr(ord("A") + column - 1), width)
    #         fo.set_font(1, column)
    #         column += 1
    #     return fo


def run(Excel_File):
    """
    主逻辑处理
    :param Excel_File: Excel_File
    :return: None
    """
    print("Current First Date:" + Date)
    li = re.findall(ALL_PAT, txt_file)
    for content in li:
        filter_flag = False
        for words in FORB_WORDS:
            if words in content:
                filter_flag = True
        for words in NESS_WORDS:
            if words not in content:
                filter_flag = True
        if filter_flag: continue
        try:
            nickname = re.search(NICK_PAT, content).group(1)
            QQ = re.search(NICK_PAT, content).group(2)
            date = re.search(DATE_PAT, content).group(1)
            time = re.search(DATE_PAT, content).group(2)
            if len(time) == 7:
                time = "0" + time
            joinedtime = date + " " + time
            if joinedtime <= BeginTime:
                continue
            currentmodifytime = joinedtime
            column_no = STARTCOLNO + (Trun_Date(date) - ydate + cal_day_offset(time)) * 6
            line_no = QQ_dict.get(QQ) if QQ_dict.get(QQ) is not None else nick_dict.get(nickname)
            if line_no is None:  # 说明没找到
                QQ_dict.update({QQ: len(QQ_dict) + TITLELINE + 1})
                print(QQ)
                line_no = QQ_dict.get(QQ)
                print(line_no)
                Excel_File.w_data(line_no, 1, QQ)
                Excel_File.w_data(line_no, 2, nickname)
            words = re.search(CONT_PAT, content).group(1)
            times = re.search(SIGN_PAT, content).group(1)
            Excel_File.w_data(line_no, column_no, char_trans(times))
            average = re.search(SIGN_PAT, content).group(2)
            total = inter_trans(times) * inter_trans(average)

            Excel_File.w_data(line_no, column_no + 1, char_trans(average))
            Excel_File.w_data(line_no, column_no + 2, total)
            Excel_File.w_data(line_no, column_no + 3, date + " " + time)
            print(nickname + "\n\t\t" + date + "\t" + words + "\n")
        except AttributeError:
            susp_list.append(content)
        except ValueError:
            susp_list.append(content)

    with open("./Error_list.txt", "w", encoding="utf-8") as f:
        f.write("以下为未解析的成功记录数据".center(60, "*") + "\n")
        for member in susp_list:
            f.write(member)
            f.write("\n" + "".center(30, "-") + "\n")
        f.write("\n\n" + "以下为未找到该成员的消息记录".center(60, "*"))
        for member in miss_list:
            f.write(member)
            f.write("\n" + "".center(30, "-") + "\n")

    # Excel_File.w_data(2, 3, currentmodifytime)


if __name__ == '__main__':
    File = table_init()
    run(File)
    # os.system("pause")

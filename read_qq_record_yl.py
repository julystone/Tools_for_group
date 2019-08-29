import sys
import re
import time

# from library import yaml
from library.R_r_excel import ReadExcel

ALL_PAT = r"(\n\n\d\d\d\d-\d\d-\d\d \d*?:\d\d:\d\d .*?\n打卡.*?次.*?\n\n)"
# CONT_PAT = r"(?<=[\>|\)])\n(.*)"
SIGN_PAT = r"\n打卡(.*)次(.*)"
NICK_PAT = r":\d{2} (.*)[\(|\<]([^\)|\>]*)"
DATE_PAT = r"(\d{4}-\d{2}-\d{2}) (\d{1,2}:\d{1,2}:\d{1,2})"

# YAML_PATH = "./config.yaml"
config_dict = {"落花": ["./『联盟群』花落闪暖，国服安一.txt", "./落花机密表.xlsx"], "月见": ["./『联盟群』月见闪暖，国服果二.txt", "./月见机密表.xlsx"],
               "月华": ["./『台服篇』愿逐月华，闪耀暖暖.txt", "./月华机密表.xlsx"]}
num_dict = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",
            "十一": "11", "十二": "12", "十三": "13", "十四": "14", "十五": "15"}
GameStartMon = "2019-8-5"
txt_file = ""
Excel_File = None
FRESHTIME = "05:00:00"
STARTCOLNO = 3  # 开始记录的列
TITLELINE = 4  # title所在行
ydate = ""
susp_list = []
BeginTime = ""
QQ_dict = {}
nick_dict = {}
Trun_Date = lambda string: time.strptime(string, "%Y-%m-%d").tm_yday


def add_zero(strr, num):
    while len(strr) < num:
        strr = "0" + strr
    return strr


def char_trans(char):
    """
    中文大写的数字数值化
    :param char: string  待处理的数据
    :return: int    对应数据
    """
    if char in num_dict.keys():
        return num_dict[char]
    return char


def trans_if_times_valid(char):
    """
    判断次数是否合法
    :param char: string  待处理的数据
    :return: int    对应数据
    """
    char = char.strip()
    if char in num_dict.keys():
        return int(num_dict[char])
    if char in num_dict.values():
        return int(char)
    raise AttributeError


def trans_if_score_valid(inter):
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
    return float(inter)


def cal_day_offset(time):
    if time > FRESHTIME:
        return 0
    else:
        return -1


def init_setpara():
    """
    表格初始化
    :return: None
    """
    global txt_file, BeginTime, ydate, Excel_File
    global GameStartMon
    choice = input("请输入联盟名：（如月华or月见or落花）")
    TXT_PATH = config_dict.get(choice)[0]
    EXCEL_PATH = config_dict.get(choice)[1]

    if choice == "月华":
        GameStartMon = "2019-04-08"

    currentdate = time.strftime("%Y-%m-%d")
    num_week = int((Trun_Date(currentdate) - Trun_Date(GameStartMon)) / 7) + 1
    sheet = "第" + str(num_week) + "周"

    #    TXT_PATH = "temp.txt"

    #    try:
    with open(TXT_PATH, "r", encoding="utf-8") as fo:
        txt_file = fo.read()
    #    except :
    #        print("无法找到文件"+TXT_PATH+":")
    #        sys.exit(1)

    #    try:
    Excel_File = ReadExcel(EXCEL_PATH, sheet)
    #    except:
    #        print("无法找到文件"+EXCEL_PATH+" "+sheet+":")
    #        sys.exit(1)

    QQ_dict.update({i.QQ: i.row for i in Excel_File.r_data_obj_from_column([1, 2], title_line=TITLELINE)})
    #    print("QQ_dict:"+str(len(QQ_dict)))
    nick_dict.update({i.nickname: i.row for i in Excel_File.r_data_obj_from_column([1, 2], title_line=TITLELINE)})
    Origin_Date = Excel_File.read_one_cell(TITLELINE - 1, STARTCOLNO)
    Date = str(Origin_Date.year) + "-" + str(Origin_Date.month) + "-" + str(Origin_Date.day)
    ydate = Trun_Date(Date)

    LastUpdateTime = Excel_File.read_one_cell(2, 3)
    if type(LastUpdateTime) != str:
        LastUpdateTime = str(LastUpdateTime.year) + "-" + add_zero(str(LastUpdateTime.month), 2) + "-" + add_zero(
            str(LastUpdateTime.day), 2) + " " + add_zero(str(LastUpdateTime.hour), 2) + ":" + add_zero(
            str(LastUpdateTime.minute), 2) + ":" + add_zero(str(LastUpdateTime.second), 2)
    BeginTime = LastUpdateTime


def run():
    """
    主逻辑处理
    :param Excel_File: Excel_File
    :return: None
    """
    currentmodifytime = BeginTime
    li = re.findall(ALL_PAT, txt_file)
    for content in li:
        #        print(content)
        try:
            nickname = re.search(NICK_PAT, content).group(1)
            QQ = re.search(NICK_PAT, content).group(2)
            date = re.search(DATE_PAT, content).group(1)
            time = re.search(DATE_PAT, content).group(2)
            if len(time) == 7:
                time = "0" + time
            joinedtime = date + " " + time
            if joinedtime <= BeginTime or Trun_Date(date) - ydate + cal_day_offset(time) < 0:  # 在上次更新时间之前的信息或者第一列之前的信息
                continue
            currentmodifytime = joinedtime

            column_no = STARTCOLNO + (Trun_Date(date) - ydate + cal_day_offset(time)) * 6
            if QQ_dict.get(QQ) is not None:
                line_no = QQ_dict.get(QQ)
            elif nick_dict.get(nickname) is not None:
                print("update ID:" + QQ)
                line_no = nick_dict.get(nickname)
                Excel_File.w_data(line_no, 1, QQ)
            else:  # 说明没找到
                line_no = len(QQ_dict) + TITLELINE + 1
                print("Add nickname:" + nickname + ",line_no:" + str(line_no))
                QQ_dict.update({QQ: line_no})
                Excel_File.w_data(line_no, 1, QQ)
                Excel_File.w_data(line_no, 2, nickname)

            times = re.search(SIGN_PAT, content).group(1)
            times = trans_if_times_valid(times)

            Excel_File.w_data(line_no, column_no, times)
            Excel_File.w_data(line_no, column_no + 3, time)
            average = re.search(SIGN_PAT, content).group(2)
            if average == "\n\n":
                continue
            average = trans_if_score_valid(average)
            total = times * average / 10000
            Excel_File.w_data(line_no, column_no + 1, average)
            Excel_File.w_data(line_no, column_no + 2, total)
            print("写入：" + nickname + "\t\t" + joinedtime + "\t" + str(total) + "w分")
        except AttributeError as e:
            susp_list.append(content)
        except ValueError as e:
            susp_list.append(content)

    with open("./log.txt", "w", encoding="utf-8") as f:
        f.write("以下为可能需要人工核对的记录".center(60, "*") + "\n")
        for member in susp_list:
            f.write(member)
            f.write("\n" + "".center(30, "-") + "\n")

    Excel_File.w_data(2, 3, currentmodifytime)
    Excel_File.close()


if __name__ == '__main__':
    init_setpara()
    run()
    input("输入任意键结束")
    # os.system("pause")

import os
import re

from library import yaml
from library.R_r_excel import ReadExcel

CONT_PAT = r"(?<=[\>|\)])\n(.*)"
SIGN_PAT = r"打卡(\d*)次(.*)"
NICK_PAT = r":\d{2} (落花.*)"
DATE_PAT = r"(\d{4}-\d{2}-\d{2}) (\d{1,2}:\d{1,2}:\d{1,2})"

YAML_PATH = "./config.yaml"

num_dict = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9", "十": "10"}

with open(YAML_PATH, 'r', encoding='utf-8') as f:
    conf = yaml.safe_load(f)
    TXT_NAME = conf['txt']['name']
    TXT_PATH = conf['txt']['path'] + TXT_NAME
    EXCEL_NAME = conf['excel']['name']
    EXCEL_PATH = conf['excel']['path'] + EXCEL_NAME

with open(TXT_PATH, "r", encoding="utf-8") as fo:
    txt_file = fo.read()


def char_trans(char):
    """
    中文大写的数字数值化
    :param char: string  待处理的数据
    :return: int    对应数据
    """
    if char in num_dict.keys():
        return num_dict[char]
    return char


def inter_trans(inter):
    """
    数据化
    :param inter: string  待处理数据
    :return: double  计算完后的数据
    """
    inter = inter.lower()
    inter = inter.strip("分")
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


def table_init():
    """
    表格初始化
    :return: None
    """
    # 如果表格已存在，则清空除第一行的所有数据
    if os.path.exists(EXCEL_PATH):
        print("读取表格中".center(24, "-"))
        fo = ReadExcel(EXCEL_PATH, "Sheet1")
        # fo.clear_sheet_except_title(3)
        return fo
        # 如果表格不存在，新建表格、写表头
    else:
        print("正在新建表格".center(24, "-"))
        ReadExcel.create_new_workbook(EXCEL_NAME)
        fo = ReadExcel(EXCEL_PATH, "Sheet1")
        fo.clear_sheet()
        column = 1
        titles = ("日期", "时间", "昵称", "单次分数", "总次数", "总分", "原文本", "部分未解析数据")
        widths = ("11", "9", "40", "9", "7", "8", "24", "68")
        for (title, width) in zip(titles, widths):
            fo.w_data(1, column, title)
            fo.set_column_width(chr(ord("A") + column - 1), width)
            fo.set_font(1, column)
            column += 1
        return fo


def run(Excel_File):
    """
    主逻辑处理
    :param Excel_File: Excel_File
    :return: None
    """
    li = txt_file.split("\n\n")
    line_no = 2
    for content in li:
        if "打卡" not in content:
            continue
        if "格式" in content or "是" in content:
            continue
        if "打卡10次4w5" in content:
            print(1)
        try:
            nickname = re.search(NICK_PAT, content).group(1)
            date = re.search(DATE_PAT, content).group(1)
            time = re.search(DATE_PAT, content).group(2)
            Excel_File.w_data(line_no, 1, date)
            Excel_File.w_data(line_no, 2, time)
            Excel_File.w_data(line_no, 3, nickname)

            words = re.search(CONT_PAT, content).group(1)
            times = re.search(SIGN_PAT, content).group(1)
            average = re.search(SIGN_PAT, content).group(2)
            total = int(times) * inter_trans(average)

            Excel_File.w_data(line_no, 4, average)
            Excel_File.w_data(line_no, 5, times)
            Excel_File.w_data(line_no, 6, total)
            print(f"{nickname}\n\t\t{date} {time}\t{words}\n")
        except AttributeError:
            line_no -= 1
        except ValueError:
            Excel_File.w_data(line_no, 8, content)
        else:
            Excel_File.w_data(line_no, 7, words)
        finally:
            line_no += 1


if __name__ == '__main__':
    File = table_init()
    run(File)
    os.system("pause")

import osimport reimport timefrom library import yamlfrom library.R_r_excel import ReadExcelCONT_PAT = r"(?<=[\>|\)])\n(.*)"SIGN_PAT = r"打卡(.*)次(.*)"NICK_PAT = r":\d{2} (.*)[\(|\<]([^\)|\>]*)"DATE_PAT = r"(\d{4}-\d{2}-\d{2}) (\d{1,2}:\d{1,2}:\d{1,2})"YAML_PATH = "./config.yaml"num_dict = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5", "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",            "十一": "11", "十二": "12", "十三": "13", "十四": "14", "十五": "15"}miss_list = []susp_list = []with open(YAML_PATH, 'r', encoding='utf-8') as f:    conf = yaml.safe_load(f)    TXT_NAME = conf['txt']['name']    TXT_PATH = conf['txt']['path'] + TXT_NAME    EXCEL_NAME = conf['excel']['name']    SHEET_NAME = conf['excel']['sheet']    EXCEL_PATH = conf['excel']['path'] + EXCEL_NAME    NESS_WORDS = conf['necessaryWords']    FORB_WORDS = conf['forbiddenWords']with open(TXT_PATH, "r", encoding="utf-8") as fo:    txt_file = fo.read()ro = ReadExcel(EXCEL_PATH, SHEET_NAME)QQ_dict = {i.QQ: i.row for i in ro.r_data_obj_from_column([1, 2], title_line=2)}nick_dict = {i.nickname: i.row for i in ro.r_data_obj_from_column([1, 2], title_line=2)}Origin_Date = ro.read_one_cell(1, 3)Date = f"{Origin_Date.year}-{Origin_Date.month}-{Origin_Date.day}"Trun_Date = lambda string: time.strptime(string, "%Y-%m-%d").tm_ydayydate = Trun_Date(Date)def char_trans(char):    """    中文大写的数字数值化    :param char: string  待处理的数据    :return: int    对应数据    """    if char in num_dict.keys():        return num_dict[char]    return chardef inter_trans(inter):    """    数据化    :param inter: string  待处理数据    :return: double  计算完后的数据    """    inter = inter.lower()    inter = inter.lower()    inter = inter.strip("分")    inter = char_trans(inter)    special_symbol = ["w", "万"]    for symbol in special_symbol:        if symbol in inter:            raw_int = inter.split(symbol)[0]            raw_int = char_trans(raw_int)            if "." in raw_int:                return float(raw_int) * 10000            raw_float = inter.split(symbol)[1]            raw_float = char_trans(raw_float)            return float(raw_int + "." + raw_float) * 10000    else:        return float(inter)def table_init():    """    表格初始化    :return: None    """    # 如果表格已存在，则清空除第一行的所有数据    if os.path.exists(EXCEL_PATH):        print("读取表格中".center(24, "-"))        fo = ReadExcel(EXCEL_PATH, SHEET_NAME)        return fodef run(Excel_File):    """    主逻辑处理    :param Excel_File: Excel_File    :return: None    """    print(f"Current First Date: {Date}")    li = txt_file.split("\n\n")    for content in li:        filter_flag = False        for words in FORB_WORDS:            if words in content:                filter_flag = True                break        for words in NESS_WORDS:            if words not in content:                filter_flag = True                break        if filter_flag: continue        try:            column_no = 3            date = re.search(DATE_PAT, content).group(1)            column_no += (Trun_Date(date) - ydate) * 6            if column_no < 0:                continue            nickname = re.search(NICK_PAT, content).group(1)            QQ = re.search(NICK_PAT, content).group(2)            time = re.search(DATE_PAT, content).group(2)            time = "0" + time if len(time) == 7 else time            line_no = QQ_dict.get(QQ) if QQ_dict.get(QQ) is not None else nick_dict.get(nickname)            if line_no is None:                miss_list.append(content)                continue            Excel_File.w_data(line_no, column_no + 3, time)            words = re.search(CONT_PAT, content).group(1)            times = inter_trans(re.search(SIGN_PAT, content).group(1))            average = inter_trans(re.search(SIGN_PAT, content).group(2))            total = times * average            Excel_File.w_data(line_no, column_no, times)            Excel_File.w_data(line_no, column_no + 1, average)            Excel_File.w_data(line_no, column_no + 2, total)            print(f"{nickname}\n\t\t{date} {time}\t{words}\n")        except AttributeError:            susp_list.append(content)        except ValueError:            susp_list.append(content)    with open(f"./Error_list_{SHEET_NAME}.txt", "w", encoding="utf-8") as f:        f.write("以下为未解析的成功记录数据".center(60, "*") + "\n")        for member in susp_list:            f.write(member)            f.write("\n" + "".center(30, "-") + "\n")        f.write("\n\n" + "以下为未找到该成员的消息记录".center(60, "*") + "\n")        for member in miss_list:            f.write(member)            f.write("\n" + "".center(30, "-") + "\n")if __name__ == '__main__':    File = table_init()    run(File)
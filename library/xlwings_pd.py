# -*- coding: utf-8 -*-
# @File   :   xlwings_pd.py
# @Author :   julystone
# @Date   :   2019/8/29 12:16
# @Email  :   july401@qq.com

import pandas as pd
from library.xlwings_excel import FormattingExcel


class OperatingExcelPd:
    def __init__(self, excel_path, sheet_name):
        self.dp = pd.read_excel(excel_path, sheet_name, skiprows=0, index_col=0)
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        # self.writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
        # self.app = xw.App(visible=False, add_book=False)
        # self.app.display_alerts = False
        # self.app.screen_updating = False
        # self.wb = self.app.books.open(excel_path)
        # self.sheet = self.wb.sheets[sheet_name]
        self.max_row = self.dp.shape[0]
        self.max_column = self.dp.shape[1]

    def write_in(self, pattern, content):
        string = f"self.dp.loc[dp.{pattern[0]} == '{pattern[1]}', '{content[0]}'] = {content[1]}"
        exec(string)

    def read_one_cell(self, row, column):
        return self.dp.iloc[row - 1, column - 1]

    def return_dp(self):
        return self.dp

    def save_to_excel(self):
        self.dp.to_excel(self.excel_path, sheet_name=self.sheet_name)
        formatter = FormattingExcel(self.excel_path, self.sheet_name)
        formatter.autofit()
        # formatter.

    def w_data(self, row, column, data):
        self.dp.iloc[row - 1, column - 1] = data


if __name__ == '__main__':
    o1 = OperatingExcelPd("./r.xlsx", "8.19-8.25")
    # o1 = OperatingExcel_Pd("./落花-统计打卡.xlsx", "8.19-8.25", [])
    # dp = o1.return_dp()
    # print(dp.QQ)
    # o1.w_data(2, 2, "123")
    # dp = o1.return_dp()
    # print(dp)
    # print(dp.nickname)
    # dp1 = dp[["QQ", "nickname"]]
    # # dp_1 = dp[["QQ"]]
    # di = dp1.to_dict("dict")
    # di = dp1.to_dict('records')
    # print(di.keys())
    # print(di['QQ'])
    # dic = {i['QQ']: i['nickname'] for i in di}
    # print(dic)
    # res = o1.read_one_cell(8, 4)
    # print(res)
    # print(dp)
    # print(dp.columns)
    # dp.loc[dp["info"]["QQ"] == "1223871051"].Monday.总分 = 1
    # dp.loc[dp["info"]["QQ"] == "1223871051"].loc[dp["Monday"]["总分"]] = 1
    # line["Monday"]["总分"] = 1
    # print(dp)
    # o1.write_in(("QQ", "1223871051"), ("总分.1", "3"))
    # print(dp.loc[dp.QQ == '1223871051', "总分.1"])
    # dp.to_excel("./text.xlsx")
    # o1.to_excel(dp)
    # dp.loc[dp["info"]["QQ"] == "1223871051", dp["Monday"]["总分"]] = 1
    # print(dp.loc[dp.QQ == '1223871051', "总分"])
    # print(dp.at[1, 2])
    # print(dp.iloc[0].loc[dp.QQ == '1223871051'])
    # print(dp["总分"])

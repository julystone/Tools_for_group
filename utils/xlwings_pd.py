# -*- coding: utf-8 -*-
# @File   :   xlwings_pd.py
# @Author :   julystone
# @Date   :   2019/8/29 12:16
# @Email  :   july401@qq.com

import pandas as pd
from library.xlwings_excel import FormattingExcel


class OperatingExcel_Pd:
    def __init__(self, excel_path, sheet_name):
        self.dp = pd.read_excel(excel_path, sheet_name, skiprows=1)
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        self.app.screen_updating = False
        self.wb = self.app.books.open(excel_path)
        self.sheet = self.wb.sheets[sheet_name]
        self.max_row = self.sheet.used_range.last_cell.row
        self.max_column = self.sheet.used_range.last_cell.column

    def write_in(self, pattern, content):
        string = f"self.dp.loc[dp.{pattern[0]} == '{pattern[1]}', '{content[0]}'] = {content[1]}"
        exec(string)

    def __del__(self):
        self.wb.save()
        pass
        self.app.quit()

    def ret_dp(self):
        return self.dp

    def save_to_excel(self):
        self.sheet.range('A1').value = self.dp


if __name__ == '__main__':
    o1 = OperatingExcel_Pd("./落花-统计打卡_3.xlsx", "8.19-8.25", [0, 1])
    # o1 = OperatingExcel_Pd("./落花-统计打卡.xlsx", "8.19-8.25", [])
    dp = o1.ret_dp()
    # print(dp)
    # print(dp.columns)
    # dp.loc[dp["info"]["QQ"] == "1223871051"].Monday.总分 = 1
    # dp.loc[dp["info"]["QQ"] == "1223871051"].loc[dp["Monday"]["总分"]] = 1
    # line["Monday"]["总分"] = 1
    print(dp)
    o1.write_in(("QQ", "1223871051"), ("总分.1", "3"))
    print(dp.loc[dp.QQ == '1223871051', "总分.1"])
    # dp.to_excel("./text.xlsx")
    # o1.to_excel(dp)
    # dp.loc[dp["info"]["QQ"] == "1223871051", dp["Monday"]["总分"]] = 1
    # print(dp.loc[dp.QQ == '1223871051', "总分"])
    # print(dp.at[1, 2])
    # print(dp.iloc[0].loc[dp.QQ == '1223871051'])
    # print(dp["总分"])

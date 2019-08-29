# -*- coding: utf-8 -*-
# @File   :   utils_op.py
# @Author :   julystone
# @Date   :   2019/8/29 9:28
# @Email  :   july401@qq.com

import pandas as pd

# class MyPanda:
#     def __init__(self, EXCEL_PATH, SHEET_NAME, title_list):
#         self.data_init = pd.read_excel(EXCEL_PATH, SHEET_NAME, header=title_list)
#
#     def __del__(self):


# data_init.reset_index()
data_init2 = pd.read_excel("./r.xlsx", "8.19-8.25", header=[0, 1])
# data_init2.reset_index()
# res = pd.MultiIndex.from_frame(data_init)
# data_init = pd.DataFrame(data_init)
# print(data_init)
print(data_init2)
# print(data_init2['info']['QQ'])
# print(data_init2['总分']['单次'])
data_init2.to_excel("./r2.xlsx", sheet_name="8.19-8.25", encoding='utf-8')
# data_init2.to_html("./test.html")
# data_df_transponse = data_df.T
# print(data_df_transponse)
# print(data_df.iloc[0,0])
# print(data_df.iloc[0,0])
# data_init.fillna(method="pad")
# sr
# print(data_init)
# print(res)
# print(data_init.columns)
# print("".center(25, "-"))
# print(data_init2.columns)
# data_init.
# print(data_init)

"""
群成员列表生成器
"""


import re

from library.R_r_excel import ReadExcel


def delblankline(infile, outfile):
    infopen = open(infile, 'r', encoding="utf-8")
    outfopen = open(outfile, 'w', encoding="utf-8")

    lines = infopen.readlines()
    for line in lines:
        if line.split():
            outfopen.writelines(line)
        else:
            outfopen.writelines("")

    infopen.close()
    outfopen.close()


def sheet_ge(excel, sheet_name, txt):
    ro = ReadExcel(excel, sheet_name)
    ro.clear_sheet_except_title()
    delblankline(txt, "temp.txt")

    with open("temp.txt", "r", encoding="utf-8") as fo:
        txt = fo.read()

    split_pattern = r" \n\d{1,3}\n"
    group_pattern = r"\n"

    res = re.split(split_pattern, txt)

    i = 2
    for j in res:
        print("".center(24, "-"))
        print(j)
        out = re.split(group_pattern, j)
        if len(out) < 7:
            out.insert(0, "无昵称")
        if sheet_name[:2] not in out[1]:
            continue
        QQ = out[2].strip()
        nick = out[1].strip()
        Enter = out[-1].strip()
        ro.w_data_origin(i, 1, QQ)
        ro.w_data_origin(i, 2, nick)
        ro.w_data_origin(i, 3, Enter)
        i += 1
    ro.save()


sheet_ge("./All.xlsx", "落花成员资料", "./txt/originQQmem_And.txt")
sheet_ge("./All.xlsx", "月见成员资料", "./txt/originQQmem_iOS.txt")

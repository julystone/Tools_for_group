from library.R_r_excel import ReadExcel

ro = ReadExcel("./all_2.xlsx", "Sheet1")

for i in ro.read_data_obj():
    print(i)
import xlwings as xw


# app = xw.App(visible=True, add_book=False)

# app.display_alerts = False
# app.screen_updating = False

# wb = xw.Book("落花-统计打卡.xlsx")
# wb = xw.Book()
#
# wb.save("temp.xlsx")
#
# wb.close()


# app.quit()
#
class Case:
    pass


class OperatingExcel:
    def __init__(self, filepath, sheetname):
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        # self.app = xw.App(visible=True, add_book=False)
        self.wb = self.app.books.open(filepath)
        self.sheet = self.wb.sheets[sheetname]

    def __del__(self):
        self.wb.save()
        pass
        self.app.quit()

    def w_data(self, start_rng, data_list):
        self.sheet.range(start_rng).value = data_list
        # self.sheet.range(start_rng).color = (0, 254, 254)
        # self.sheet.range("B1").api.Font.Color = 0x0000ff
        # self.sheet.range("B1").api.Interior.Color = 0x0000ff
        self.sheet.range("B1:B3").api.Borders.LineStyle = 9
        # self.sheet.range("B1:B3").api.Borders.Weight = 4
        self.sheet.range("B1:B3").api.Borders.Color = 0x0000ee
        print(self.sheet.range("B1").api.Borders.Color)
        print(self.sheet.range("B1").api.Borders.LineStyle)
        print(self.sheet.range("B1").api.Font.Color)
        self.sheet.range("B1").api.Font.Size = 20
        print(self.sheet.range("B1").api.Font.Size)
        # self.wb.save()

    def clear_sheet(self):
        self.sheet.clear()

    def autofit(self, type='c'):
        self.sheet.autofit(type)
        # self.sheet.shapes[0]

    def read_obj(self):
        titles = []
        temp_value = None
        for rng in self.sheet.range("1:1"):
            if not rng.value and not rng.api.MergeCells:
                break
            # if rng.value:
            #     temp_value = rng.value
            # titles.append(temp_value)
            titles.append(rng.value)
        print(titles)
        second_titles = []
        for rng in self.sheet.range("2:2"):
            if not rng.value and not rng.api.MergeCells:
                break
            if rng.value:
                temp_value = rng.value
            second_titles.append(temp_value)
        print(second_titles)

        # rows_data = self.sheet.last_cell.value
        # print(rows_data.value)
        # titles = []
        # for title in rows_data[0]:
        #     titles.append(title)
        # cases = []
        # for case in rows_data[1:]:
        #     # 创建一个Cases类的对象，用来保存用例数据，
        #     case_obj = Case()
        #     # data用例临时存放用例数据
        #     data = []
        #     for cell in case:
        #         data.append(cell)
        #     case_data = list(zip(titles, data))
        #     for i in case_data:
        #         if i[0] == 'result' or i[0] is None:
        #             continue
        #         setattr(case_obj, i[0], i[1])
        #     setattr(case_obj, 'row', case[0].row)
        #     cases.append(case_obj)

        # print(titles)
        # print(cases[0].日期)
        # return cases

    def read_excel(self):
        rows_data = self.sheet.range("A1").expand().value
        print(rows_data)


o1 = OperatingExcel("./ALL.xlsx", "Sheet1")
# o1.clear_sheet()
# o1.w_data("A1:A10", ["aaaaaaaaaaaaaaa", "bbb", "ccc", "aaaaaaaaaaaaaaa", "bbb", "ccc", "aaaaaaaaaaaaaaa", "bbb", "ccc"])
# o1.sheet_autofit()
o1.read_obj()
# o1.read_excel()

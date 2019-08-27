import xlwings as xw


class MyDict(dict):
    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__

    def __getitem__(self, key):
        if key in self.keys():
            dict.__getitem__(self, key)
        else:
            r = MyDict()
            # self.__setattr__[key] = r
            self.key = r
            # self.__setitem__(key, r)
            return r


def dict_to_object(dict_obj):
    if isinstance(dict_obj, dict) is False:
        return dict_obj
    inst = MyDict()
    for k, v in dict_obj.items():
        inst[k] = dict_to_object(v)
    return inst


class OperatingExcel:
    def __init__(self, filepath, sheetname):
        self.app = xw.App(visible=False, add_book=False)
        self.app.display_alerts = False
        # self.app = xw.App(visible=True, add_book=False)
        self.wb = self.app.books.open(filepath)
        self.sheet = self.wb.sheets[sheetname]
        self.max_row = self.sheet.used_range.last_cell.row
        self.max_column = self.sheet.used_range.last_cell.column

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

    def read_titles(self, deep_sides):
        all_title = []
        temp_value = None
        for title_row in range(1, deep_sides + 1):
            contents = []
            for rng in self.sheet.range(f"{title_row}:{title_row}"):
                if rng.column > self.max_column:
                    break
                if rng.value:
                    temp_value = rng.value
                contents.append(temp_value)
            all_title.append(contents)
        return all_title

    def read_one_line(self, row):
        contents = []
        for rng in self.sheet.range(f"{row}:{row}"):
            if rng.column > self.max_column:
                break
            contents.append(rng.value)
        return contents

    # dict_all.ff(123).ff(456).ff(456)

    def read_data(self):
        title_deep_sides = 2
        all_cases = []
        all_title = self.read_titles(title_deep_sides)
        for data_row in range(title_deep_sides + 1, self.max_row + 1):
            data = self.read_one_line(data_row)
            dict_all = MyDict()
            zip_tuple = list(zip(*all_title, data))
            string = "dict_all"
            for _ in range(title_deep_sides):
                string += f".ff(item[{_}])"
            string += f" = item[{_ + 1}]"
            for item in zip_tuple:
                try:
                    print(dict_all[item[0]])
                    print(dict_all[item[0]][item[1]])
                    dict_all[item[0]][item[1]] = item[2]
                    # exec(string)
                except KeyError:
                    dict_all[item[0]] = {}
                    exec(string)
            res = dict_to_object(dict_all)
            all_cases.append(res)
        for index, value in enumerate(all_cases):
            value.row = title_deep_sides + 1
        print(all_cases)

    def read_excel(self):
        rows_data = self.sheet.range("A1").expand().value
        print(rows_data)

    @staticmethod
    def get_value(di, ke):
        return di[ke]


o1 = OperatingExcel("./ALL_4.xlsx", "Sheet1")
# o1.clear_sheet()
# o1.sheet_autofit()
o1.read_data()
# fff = MyDict()
# fff["1"]
# print(fff)

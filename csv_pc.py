li = [('time', '日期', '2019-08-13'), ('time', '时间', '13:39:02'), ('grades', '昵称', '落花ʚ阿喵王子他父王ɞ'),
      ('grades', '单次', '2w7'), ('grades', '次数', '10'), ('grades', '总分', 270000.0), ('exception', '异常', '打卡10次，单次4万六\n')]


# class WPUnit(dict):
# def __init__(self):
#     self._res = {}
#
# def __setitem__(self, key, val):
#     self._res[key] = val
#
# def __getitem__(self, key):
#     # if self._res.has_key(key):
#     if key in self._res.keys():
#         return self._res[key]
#     else:
#         r = WPUnit()
#         self.__setitem__(key, r)
#         return r


# a = WPUnit()
# a['a']['b']['c']['d']['e']['f']['g'] = 5
# print(a['a']['b']['c']['d']['e']['f']['g'])


class MyDict(dict):
    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__

    def __getitem__(self, k):
        try:
            return super().__getitem__(k)
        except KeyError:
            r = MyDict()
            dict.__setitem__(self, k, r)


a = MyDict()
# a.ff(123).ff(456).ff(456).ff(456)
a[123] = 1
a[123] = 56
a[12354]
a[12354][5]
# print(type(a[12354]))
print(a[12354][1])
a[12354][1][000] = 156
# a[123] = 1
# a[123][456] = 3
# print(a[123][456])
print(a[12354][1][000])
print(type(a[123]))
# print(type(a[123][456]))
# print(a[123][456])
# print(a[123][456])
# print(a[123][1][2][3])
print(a)
# print(type(a))

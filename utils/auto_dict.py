# -*- coding: utf-8 -*-
# @File   :   good_dict.py
# @Author :   julystone
# @Date   :   2019/8/28 9:59
# @Email  :   july401@qq.com


class AutoDict(dict):
    def __init__(self):
        self._res = {}

    def __setitem__(self, key, val):
        self._res[key] = val
        dict.__setitem__(self, key, val)

    def __getitem__(self, key):
        if key in self._res.keys():
            return self._res[key]
        else:
            r = AutoDict()
            self.__setitem__(key, r)
            return r


class _DictObj(dict):
    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__


def dict_to_object(dic):
    if isinstance(dic, dict) is False:
        return dic
    inst = _DictObj()
    for key, value in dic.items():
        inst[key] = dict_to_object(value)
    return inst

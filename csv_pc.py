s0 = "abcacde"
s1 = "abcacdeb"
s2 = "abcacdee"

set_temp = set()

li = []

for index, value in enumerate(s2):
    if value not in set_temp:
        li.append(value)
    else:
        li.remove(value)
    set_temp.add(value)

print(li[0], s2.index(li[0]))

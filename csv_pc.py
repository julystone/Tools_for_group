from dateutil import rrule, parser

# Origin_Date = "2019-08-06 23:49:00"
# Test_Date = "2019-08-08 00:00:00"
#
# Trun_Date = lambda string: time.strptime(string, "%Y-%m-%d %H:%M:%S").tm_yday
# ydate = Trun_Date(str(Origin_Date))
#
# print(Trun_Date(str(Test_Date)) - Trun_Date(str(Origin_Date)))

res = rrule.rrule(rrule.DAILY, dtstart=parser.parse('2019-08-06 23:49:00'),
                  until=parser.parse('2019-08-07 23:48:59')).count()
print(res)

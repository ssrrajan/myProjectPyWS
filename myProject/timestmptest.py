from datetime import datetime
dateTimeObj = datetime.now()
print(dateTimeObj)

print(dateTimeObj.year, dateTimeObj.month, dateTimeObj.day)


from datetime import datetime

now = datetime.now() # current date and time

year = now.strftime("%Y")
print("year:", year)

month = now.strftime("%m")
print("month:", month)

day = now.strftime("%d")
print("day:", day)

time = now.strftime("%Y%m%d%H%M%S")
print("time:", time)

date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
print("date and time:",date_time)
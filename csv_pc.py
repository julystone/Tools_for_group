import csv

with open("./test11.csv", "r") as csvFile:
    reader = csv.reader(csvFile)
    for item in reader:
        print(item)

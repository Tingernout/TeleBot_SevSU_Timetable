import datetime
from openpyxl import load_workbook
file = load_workbook(filename='iit_23_b_o28.03.xlsx')
day = datetime.date.today()
today = day.strftime("%d.%m.%Y")

for a in range(23, 35):
    list = file['уч.н. ' + str(a)]
    for i in range(7, 56, 8):
        if list['B' + str(i)].value == today:
            for n in range(i, i + 8):
                noneF = 'F'
                noneG = 'G'
                if list['F' + str(n)].value == None:
                    noneF = "I"
                    noneG = "J"
                print(str(list['C' + str(n)].value) + " " + str(list['D' + str(n)].value) + " " + str(list['E' + str(n)].value) + " " + str(list[noneF + str(n)].value) + " " + str(list[noneG + str(n)].value) + "\n\n")




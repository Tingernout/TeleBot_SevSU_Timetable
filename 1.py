from openpyxl import load_workbook
file = load_workbook(filename='iit_23_b_o28.05.xlsx')
data = [[[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []]]
a = 23
list = file['уч.н. ' + str(a)]
stroca = 6
dd = 7
for data_day_element in range(6):
    data[data_day_element].append(list['A' + str(dd)].value)
    dd += 8
    for data_para_number_element in range(0, 8):
        stroca += 1
        for stolbec in "CDEFG":
            data[data_day_element][data_para_number_element].append(list[stolbec + str(stroca)].value)






print(data)






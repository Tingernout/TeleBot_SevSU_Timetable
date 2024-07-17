from openpyxl import load_workbook
file = load_workbook(filename='iit_23_b_o28.05.xlsx')
data = [[[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []]]
a = 23 #номер недели
list = file['уч.н. ' + str(a)]
stroca = 7 #строка самой первой пары в таблице эксель
dd = 7 #строка дня недели понедельника
for data_day_element in range(6):
    for data_para_number_element in range(0, 8):

        for stolbec in "CDEFG": #"ходьба" между столбцов - там находится основная инфа (кто где когда) (столбец С - номер пары, D - время и тд)
            if list['E' + str(stroca)].value == None:
                break
            data[data_day_element][data_para_number_element].append(list[stolbec + str(stroca)].value) #добавление информации о паре (кто где когда)
        stroca += 1  # перевод на следующую пару
    data[data_day_element].insert(0, list['A' + str(dd)].value)  # добавление дня недели
    dd += 8  # перевод на следущую ячейку, где содержиться день недели


for i in range(6):
    for a in range(9):
        print(data[i][a])






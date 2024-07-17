from openpyxl import load_workbook



file = load_workbook(filename='iit_23_b_o28.05.xlsx')
data = [[[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []], [[], [], [], [], [], [], [], []]]
a = 24 #номер недели
list = file['уч.н. ' + str(a)]
stroca = 7 #строка самой первой пары в таблице эксель
dd = 7 #строка дня недели понедельника
t = 0 #переменная для подсчета количества добавлений, если в расписании стоит общая пара



for data_day_element in range(6):
    for data_para_number_element in range(0, 8):
        for stolbec in "CDEFGHIJ": #"ходьба" между столбцов - там находится основная инфа (кто где когда) (столбец С - номер пары, D - время и тд)
            if list['E' + str(stroca)].value == None: #если пусто то забить
                break
            #if "\n" in list['E' + str(stroca)].value:  ----- придумать что-то

            if ',' in str(list[str(stolbec) + str(stroca)].value): #если есть запятая то разделить строку на несколько частей и добавть эти части
                k = str(list[stolbec + str(stroca)].value).split(',')
                for i in range(len(k)):#добавление частей
                    data[data_day_element][data_para_number_element].append(k[i])
                continue #переход на слеующий столбец

            if list[str(stolbec) + str(stroca)].value == None: #для отображения аудитории и типа занятия если пара стоит общая
                if t == 2: #без этого условия будетдобавлено в массив тип занятия, кабинет, 3 штуки None и опять тип и аудиория, поэтому я ограничил добавление до 2-х
                    break
                stolbec = str(chr(ord(stolbec) + 3)) #перевод ячейки на 3 столбца вперед, где она НЕ пустая
                data[data_day_element][data_para_number_element].append(list[stolbec + str(stroca)].value)
                t += 1
            else:
                data[data_day_element][data_para_number_element].append(list[stolbec + str(stroca)].value) #добавление информации о паре (кто где когда)
        stroca += 1  # перевод на следующую пару
    data[data_day_element].insert(0, list['A' + str(dd)].value)  # добавление дня недели
    dd += 8  # перевод на следущую ячейку, где содержиться день недели


for i in range(6):
    for a in range(9):
        print(data[i][a])







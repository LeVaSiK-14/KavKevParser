import openpyxl

zp = openpyxl.open("./test.xlsx", read_only = True)

sheet = zp.active

mainArr = []

for row in range(12, 150):
    wokrArr = []
    splitPep = str(sheet[row][0].value).split(' ')# имена сотрудников
    if len(splitPep) > 1:
        parsRow = (str(sheet[row][0])).split('.') # нужные строки
        parsRowStrip = parsRow[1].lstrip('A').rstrip('>') # нужные числа из строк
        for i in range(0, 29):
            if sheet[parsRowStrip][i].value == None: # Если значение пустое то вместо None выводить 0
                wokrArr.append(0)
            else:
                wokrArr.append(sheet[parsRowStrip][i].value)
        mainArr.append(wokrArr)
    else:
        pass


finArrZp = []
for i in range(0, len(mainArr)):
    item = mainArr[i]

    workDay = item[3] - 26
    oklad = float(8000 + (workDay * 258))
    # if item[4] == 8000:
    #     oklad = float(8000 + (workDay * 258))
    # elif item[4] == 11000:
    #     oklad = float(1000 + (workDay * 354))
    print((item[20] - item[21]) * 2)
    finalZP = (item[0], (oklad + item[6] * item[7]) + ((item[20] - item[21]) * 2) + item[17] - item[11] - item[12] - item[13] - item[14] - item[15] - item[16] - item[18] - item[19] - item[24] - item[25] - item[26] - item[27] - item[28])
    finArrZp.append(list(finalZP))


print(finArrZp)
import os
import pandas as pd
import gc

path = os.getcwd() #Текущая папка

allFiles = os.listdir(path) #Список файлов и директорий в папке path

# В цикле идем по списку файлов при условии что нам нужны только .xlsx

for file in allFiles:
    if file[-4:] == 'xlsx':
        print("Нашли xlsx: " + file + "\n")
        dataFile = pd.ExcelFile(file)
        dataSheet = dataFile.sheet_names
        for sheet in dataSheet:
            if sheet[-2:] == '_O':
                print("Нашли лист с _O: " + sheet + "\n")
                df1 = pd.read_excel(file, sheet)
                print("Записали лист в df1\n")
                df1.columns = df1.iloc[0,:]
                df1.drop([0], inplace = True)
                df1_tb = df1.loc[df1['Демография_1'] == '2. ТБ']
                beg = df1_tb['Демография_2'].unique()
                print("Обработали лист по правилам\n")
                for b in beg:
                    df1_tb1 = df1_tb.loc[df1_tb['Демография_2'] == b ]
                    df1_tb1.to_excel(b + '_' + sheet + '.xlsx')
                    print("Cоздали новый файл\n")
        gc.collect()
        print("Очистили память после цикла с поиском листов\n")

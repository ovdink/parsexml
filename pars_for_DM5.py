import os
import pandas as pd
import gc

path = os.getcwd()
allFiles = os.listdir(path)

#arrSheets = []

for file in allFiles:
    if file[-4:] == 'xlsx':
        print('\nНашли xlsx: ' + file + '\n')
        dataFile = pd.ExcelFile(file)
        dataSheet = dataFile.sheet_names
        for sheet in dataSheet:
            if sheet[-3:] != '_ФВ' and sheet[-2:] != '_O' and sheet != 'Свод' and sheet[-2:] != 'ы)' and sheet[-7:] != '_ФВ (3)' and sheet[-7:] != '_ФВ (2)' and sheet[-7:] != '_ФВ (1)' and sheet[-7:] != '_ФВ (4)' and sheet[-7:] != '_ФВ (5)':
                print('Нашли лист: ' + sheet + '\n')
                df = pd.read_excel(file, sheet)
                df = df.iloc[13:]
                df.columns = df.iloc[0,:]
                df.drop([13], inplace = True)
                df_d5 = df.loc[df['Демография_5'] == 'Подразделение прямых продаж']
                df_d5.reset_index()
                df_d5.to_excel(file[:-4] + sheet + '_ППП' + '.xlsx')
                print('Создали файл\n')
                gc.collect()
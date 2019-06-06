import os
import pandas as pd

path = os.getcwd()
allFiles = os.listdir(path)

arrSheets = []

for file in allFiles:
    if file[-4:] == 'xlsx':
        dataFile = pd.ExcelFile(file)
        dataSheet = dataFile.sheet_names
        for sheet in dataSheet:
            if sheet[-2:] == '_O':
                  df1 = pd.read_excel(file, sheet)
                  df1.columns = df1.iloc[0,:]
                  df1.drop([0], inplace = True)
                  df1_tb = df1.loc[df1['Демография_1'] == '2. ТБ']
                  #df1_tb['Демография_2'].value_counts()
                  
                  beg = df1_tb['Демография_2'].unique()
                  for b in beg:
                      df1_tb1 = df1_tb.loc[df1_tb['Демография_2'] == b ]
                      df1_tb1.to_excel(b + '_' + sheet + '.xlsx')
                  # summ = df1.append(df1)
                  #print(summ + '\n')

        
 #print(file_xlsx)
        
#files_xlsx = [f for f in files if f[-4:] == 'xlsx']

#df = pd.DataFrame()


#for f in files_xlsx:
#    lists = pd.read_excel(f)
#    dl = df.append(lists)


#list_xlsx = [l for l in lists if l[-2:] == '_O']

#data = pd.read_excel(f)

#for f in files_xlsx:
#    data = pd.read_excel(f)
#    df = df.append(data)

#dd = df
#data.parse(sheet_name)


# allSheets = []
# for f in files_xlsx:
#     data = pd.ExcelFile(f)
#     data1 = data.sheet_names
#     allSheets.append(data1)
#     #print(allSheets)
    
# #print("____________\n")

# withO = []
# for el in allSheets:
#     if el[-2:] == '_O':
#         withO.append(el)
        
# print(withO)
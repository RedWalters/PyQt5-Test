import pandas as pd
import os


def saveToCsv(dataFrame):
    try:
        print(f'      Saving: ' + os.path.splitext(dataFrame['File Name'].iloc[1])[0] + '_' + os.path.splitext(dataFrame['Sheet Name'].iloc[1])[0])
        dataFrame.to_csv(os.path.splitext(dataFrame['File Name'].iloc[1])[0] + '_' + os.path.splitext(dataFrame['Sheet Name'].iloc[0])[0] +'.csv', sep = '|', index = False, chunksize = 100000, encoding='utf-8')
    except:
        try:
            print(f'      Saving: ' + os.path.splitext(dataFrame['File Name'].iloc[1])[0])
            dataFrame.to_csv(os.path.splitext(dataFrame['File Name'].iloc[1])[0] +'.csv', sep = '/', index = False, chunksize = 100000, encoding='utf-8')
        except Exception as e:
            print(e)
from ntpath import basename
import pandas as pd
import sys
import os



def read_sheets(filename):
    try:
        print(f'Importing: ' + basename(filename))
        file = pd.ExcelFile(filename)
        sheets = file.sheet_names
        dfs = []
        df = pd.DataFrame()
        for i in sheets:
            print(f'     Sheet: ' + i)
            df = file.parse(i, header = None, encoding="utf-8-sig").fillna('')
            df['filename'] = basename(filename)
            df['sheetname'] = i
            dfs.append(df)
        return dfs
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(str(e))
    
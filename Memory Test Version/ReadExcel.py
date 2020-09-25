from ntpath import basename
import pandas as pd
import sys
import os


def read_excel(filename):
    try:
        try:
            print(f'Importing: ' + basename(filename))
            dfBeta = pd.read_excel(filename, sheet_name = None, header = None, encoding="utf-8-sig")
            dfBeta = pd.concat([dfBeta[frame] for frame in dfBeta.keys()], ignore_index = True).fillna('')
            dfBeta['filename'] = basename(filename)
            return dfBeta
        except:
            try:
                print(f'Importing: ' + basename(filename))
                dfBeta = pd.read_csv(filename, header = None, encoding = "utf-8-sig", low_memory = False).fillna('')
                dfBeta['filename'] = basename(filename)
                return dfBeta
            except:
                print(f'Importing: ' + basename(filename))
                dfBeta = pd.read_csv(filename, header = None, encoding = "ISO-8859-1", low_memory = False).fillna('')
                dfBeta['filename'] = basename(filename)
                return dfBeta
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(str(e))
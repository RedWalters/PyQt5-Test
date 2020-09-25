import pandas as pd
import sys
import os
"""
0 1 2 3 4 5 6 7 8 9 10 11 12 13 14 15 16 17 18 19 20 21 22 23 24 25 26 27 28 29 30 31 32 33 34 35 36 37 38 39 40
A B C D E F G H I J K  L  M  N  O  P  Q  R  S  T  U  V  W  X  Y  Z  AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO
"""
def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

def formatting(samplePD):
    rowCount =  len(samplePD)
    rawRow = 0
    table = {}
    try:
        while rawRow < rowCount:
            
            if mid(str(samplePD.loc[rawRow, 0]), 4, 1) != '-' and str(samplePD.loc[rawRow, 0]) != '':
                account = samplePD.loc[rawRow,0]
            elif mid(str(samplePD.loc[rawRow, 0]), 4, 1) == '-' and str(samplePD.loc[rawRow, 2]) != 'Opening balance':
                table[rawRow] = {  'Account':         account,
                                   'Date':            str(samplePD.loc[rawRow,0]),
                                   'Type':            str(samplePD.loc[rawRow,1]),
                                   'Transaction':     str(samplePD.loc[rawRow,2]),
                                   'Reference':       str(samplePD.loc[rawRow,3]),
                                   'Debit':           str(samplePD.loc[rawRow,4]),
                                   'Credit':          str(samplePD.loc[rawRow,5]),
                                   'VAT':             str(samplePD.loc[rawRow,6])}
            rawRow +=1 
        df = pd.DataFrame.from_dict(table, "index")
    
        return df
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(str(e))
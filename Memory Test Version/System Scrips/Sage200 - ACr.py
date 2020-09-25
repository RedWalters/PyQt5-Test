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
    rawRow = 1
    table = {}
    print(f'Formatting: ' + samplePD.loc[1, 'File Name'])
    try:
        while rawRow < rowCount:
            if str(samplePD.loc[rawRow, 2]) == 'A/C No.':
                accountNumber = samplePD.loc[rawRow,3]
                accountDescription = samplePD.loc[rawRow,10]
            elif mid(str(samplePD.loc[rawRow, 6]), 4, 1) == '-':
                table[rawRow] = {  'Seq':               str(samplePD.loc[rawRow,2]),
                                   'URN':               str(samplePD.loc[rawRow,2]),
                                   'TranDate':          str(samplePD.loc[rawRow,6]),
                                   'Period':            str(samplePD.loc[rawRow,8]),
                                   'Reference':         str(samplePD.loc[rawRow,10]),
                                   'Narrative':         str(samplePD.loc[rawRow,11]),
                                   'Debit':             str(samplePD.loc[rawRow,17]),
                                   'Credit':            str(samplePD.loc[rawRow,19]),
                                   'NomTranAnalysis1':  str(samplePD.loc[rawRow,21]),
                                   'A/C No.':           accountNumber,
                                   'A/C Name':          accountDescription,
                                   'File Name':         samplePD.loc[rawRow, 'File Name']}
            rawRow +=1 
        df = pd.DataFrame.from_dict(table, "index")
    
        return df     
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(str(e))
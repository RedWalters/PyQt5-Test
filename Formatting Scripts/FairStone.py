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
    #print(f'Formatting: ' + samplePD.loc[1, 'File Name'])
    try:
        while rawRow < rowCount:
            if str.isnumeric(left(str(samplePD.loc[rawRow,0]), 1)) == False:
                accountNumber = samplePD.loc[rawRow,0]
                accountDesc = samplePD.loc[rawRow,1]
            elif str(samplePD.loc[rawRow,0]) != '':
                table[rawRow] = {  'AccountNo':              accountNumber,
                                   'AccountDesc':            accountDesc,
                                   'Posting Date':           str(samplePD.loc[rawRow,0]),
                                   'Due Date':               str(samplePD.loc[rawRow,1]),
                                   'Series':                 str(samplePD.loc[rawRow,2]),
                                   'Doc No':                 str(samplePD.loc[rawRow,3]),
                                   'Trans No':               str(samplePD.loc[rawRow,4]),
                                   'Remarks':                str(samplePD.loc[rawRow,5]),
                                   'Offset Acct':            str(samplePD.loc[rawRow,6]),
                                   'Offset Acct Name':       str(samplePD.loc[rawRow,7]),
                                   'Deb/Cred (LC)':          str(samplePD.loc[rawRow,8]),
                                   'Cumulative Balance (LC)':str(samplePD.loc[rawRow,9]),
                                   'Blanket Agreement':      str(samplePD.loc[rawRow,10]),
                                   'Seq No':                 str(samplePD.loc[rawRow,11])}
            rawRow +=1 
        df = pd.DataFrame.from_dict(table, "index")
    
        return df
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(str(e))
            
    

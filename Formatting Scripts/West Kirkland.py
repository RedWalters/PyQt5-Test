import pandas as pd
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
    #print(f'Formatting: ' + samplePD.loc[1, 'File Name'])
    try:
        while rawRow < rowCount:
            if str.isnumeric(left(str(samplePD.loc[rawRow,1]), 4)) == True and samplePD.loc[rawRow, 11] != '' and len(str(samplePD.loc[rawRow,1])) > 2:
                accountNumber = samplePD.loc[rawRow,1]
                accountDesc = samplePD.loc[rawRow,11]
            elif mid(str(samplePD.loc[rawRow,7]), 4, 1) == '-':
                if samplePD.loc[rawRow, 26] == '':
                    debit = 0
                    credit = samplePD.loc[rawRow, 30]
                    amount = samplePD.loc[rawRow, 30] * -1
                else: 
                    credit = 0
                    debit = samplePD.loc[rawRow, 26]
                    amount = samplePD.loc[rawRow, 26]
                table[rawRow] = {  'AccountNo':     accountNumber,
                                   'AccountDesc':   accountDesc,
                                   'Period':        str(samplePD.loc[rawRow,1]),
                                   'Source':        str(samplePD.loc[rawRow,5]),
                                   'Doc. Date':     str(samplePD.loc[rawRow,7]),
                                   'Description':   str(samplePD.loc[rawRow,11]),
                                   'Posting Sequence':   str(samplePD.loc[rawRow,21]),
                                   'Batch-Entry':     str(samplePD.loc[rawRow,23]),
                                   'Debit':         debit,
                                   'Credit':        credit,
                                   'Amount':        amount}
            rawRow +=1 
        df = pd.DataFrame.from_dict(table, "index")
    
        return df
    except Exception as e:
        print(e)
            
    

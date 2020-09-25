import pandas as pd
import os

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

    while rawRow < rowCount:
        if left(samplePD.loc[rawRow,1], 13) == 'Account Code:':
            accountCode = right(samplePD.loc[rawRow,1], 3)
        elif str.isnumeric(left(samplePD.loc[rawRow,1], 7)) == True and samplePD.loc[rawRow,1] != 'Entry Date':
            accountNumber = left(samplePD.loc[rawRow,1], 7)
            accountDesc = mid(samplePD.loc[rawRow,1], 9, len(samplePD.loc[rawRow,1])).strip()
        elif str.isnumeric(left(samplePD.loc[rawRow,1],2)) == False and '20' in samplePD.loc[rawRow,1]:
            period = samplePD.loc[rawRow,1]
        elif ' 20' not in samplePD.loc[rawRow,1] and len(samplePD.loc[rawRow,1]) == 9:
            date = str(samplePD.loc[rawRow,1])
        elif samplePD.loc[rawRow,2] != '' and samplePD.loc[rawRow,1] != 'Entry Date': 
            table[rawRow] = {  'Entity':        accountCode,
                               'AccountNo':     accountNumber,
                               'AccountDesc':   accountDesc,
                               'Month':         period,
                               'Date':          date,
                               'JournalNo':     samplePD.loc[rawRow,2],
                               'Description':   samplePD.loc[rawRow,4],
                               'Amount':        samplePD.loc[rawRow,8],
                               'Balance':       samplePD.loc[rawRow,11]}
        rawRow +=1        
    
    df = pd.DataFrame.from_dict(table, "index")
    
    return df
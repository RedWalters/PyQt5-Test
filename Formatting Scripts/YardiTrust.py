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
    rawRow = 1
    table = {}
    
    #print(f'Formatting: ' + samplePD.loc[1,'File Name'] +' - ' + samplePD.loc[1,'Sheet Name'])
    
    while rawRow < rowCount:
        if str.isnumeric(left(samplePD.loc[rawRow,0], 5)) == True and samplePD.loc[rawRow,1] == '':
            accountCode = samplePD.loc[rawRow,0]
            accountDesc = samplePD.loc[rawRow,4]
        elif '-' in str(samplePD.loc[rawRow,2]) and '-' in str(samplePD.loc[rawRow,3]):
            table[rawRow] = {  'Account Number':        accountCode,
                               'Account Description':   accountDesc,
                               'Property':              samplePD.loc[rawRow,0],
                               'Property Name':         samplePD.loc[rawRow,1],
                               'Date':                  samplePD.loc[rawRow,2],
                               'Period':                samplePD.loc[rawRow,3],
                               'Person/Description':    samplePD.loc[rawRow,4],
                               'Control':               samplePD.loc[rawRow,5],
                               'Reference':             samplePD.loc[rawRow,6],
                               'Debit':                 samplePD.loc[rawRow,7],
                               'Credit':                samplePD.loc[rawRow,8],
                               'Balance':               samplePD.loc[rawRow,9],
                               'Remarks':               samplePD.loc[rawRow,10].replace('\"', '')}
        rawRow +=1 
        
    df = pd.DataFrame.from_dict(table, "index")
    
    return df

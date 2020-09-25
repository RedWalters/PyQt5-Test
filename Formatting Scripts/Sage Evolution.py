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
    
    #print(f'Formatting: ' + samplePD.loc[1, 'File Name'])

    while rawRow < rowCount:
        if len(str(samplePD.loc[rawRow,0])) == 4 and str(samplePD.loc[rawRow, 0]).count('/') == 0 and samplePD.loc[rawRow, 0] != 'Date':
            accountNumber = samplePD.loc[rawRow,0]
        elif str(samplePD.loc[rawRow, 0]).count('/') > 2 and str.isnumeric(left(samplePD.loc[rawRow, 0], 4)) == False:
            journalDescription = samplePD.loc[rawRow, 0]
        elif str(samplePD.loc[rawRow, 0]).count('/') > 2 and str.isnumeric(left(samplePD.loc[rawRow, 0], 4)) == True:
            journalDetails = samplePD.loc[rawRow,0]
        elif str(samplePD.loc[rawRow, 0]).count('/') == 2 and len(str(samplePD.loc[rawRow,0])) == 10: 
            table[rawRow] = {  'Account Number':        accountNumber,
                               'Journal Description':   journalDescription,
                               'Journal Details':       journalDetails,
                               'Date':                  samplePD.loc[rawRow,0],
                               'Reference':             samplePD.loc[rawRow,1],
                               'Description':           samplePD.loc[rawRow,2],
                               'Debit':                 samplePD.loc[rawRow,4],
                               'Credit':                samplePD.loc[rawRow,5],
                               'Balance':               samplePD.loc[rawRow,6]}
        rawRow +=1        
    
    df = pd.DataFrame.from_dict(table, "index")
    
    return df
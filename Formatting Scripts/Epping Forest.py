import pandas as pd
import os

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

def formatting(samplePD):
    print(len(samplePD))
    rowCount =  len(samplePD)
    rawRow = 0
    table = {}
    
    print(samplePD.loc[1,'File Name'])

    while rawRow < rowCount:
        if samplePD.loc[rawRow,3] == 'Account Code':
            accountCode = samplePD.loc[rawRow +1, 3]
        elif '-' in str(samplePD.loc[rawRow,8]) or '/' in str(samplePD.loc[rawRow,8]): 
            table[rawRow] = {  'Cost Centre':                   samplePD.loc[rawRow,1],
                               'CC Name':                       samplePD.loc[rawRow,2],
                               'Account Code':                  accountCode,
                               'AC Name':                       samplePD.loc[rawRow,4],
                               'Job Code':                      samplePD.loc[rawRow,5],
                               'JC Name':                       samplePD.loc[rawRow,6],
                               'Period':                        samplePD.loc[rawRow,7],
                               'Entry Date':                    samplePD.loc[rawRow,8],
                               'Journal Name':                  samplePD.loc[rawRow,9],
                               'Tranaction Description':        samplePD.loc[rawRow,10],
                               'Transaction Type Description':  samplePD.loc[rawRow,11],
                               'Transaction Ref':               samplePD.loc[rawRow,12],
                               'Trans Sum':                     samplePD.loc[rawRow,13],
                               'Posting User':                  samplePD.loc[rawRow,14],
                               'FileName':                      samplePD.loc[rawRow,'File Name']}
        rawRow +=1        
    df = pd.DataFrame.from_dict(table, "index")
    
    return df
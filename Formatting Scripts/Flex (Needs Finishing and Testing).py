import pandas as pd

'a,b,c,d,e,f,g,h,i,j,k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, z'
'0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25'

def left(s, amount):
    return s[:amount]

def right(s, amount):
    return s[-amount:]

def mid(s, offset, amount):
    return s[offset:offset+amount]

def formatting(data):
    rowCount =  len(data)
    rawRow = 2
    table = {}

    while rawRow < rowCount:
        if data.loc[rawRow,4] == 'SOURCE BATCH ID':
            JournalNo = data.loc[rawRow,0]
            JournalDesc = data.loc[rawRow,7]
            Count = 0
        elif mid(data.loc[rawRow,1], 4, 1) == '-': 
            table[rawRow] = {  'Entity':        data.loc[rawRow,'File Name'],
                               'JournalNo':     JournalNo, 
                               "JournalDesc":   JournalDesc,
                               "LineNo":        data.loc[rawRow,0],
                               "AccountNo":     data.loc[rawRow,1],
                               "AccountName":   data.loc[rawRow,2],
                               "Currency":      data.loc[rawRow,3],
                               "DC Indicator":  data.loc[rawRow,5],
                               "FXrate":        data.loc[rawRow,6],
                               "Debit":         data.loc[rawRow,7],
                               "Credit":        data.loc[rawRow,8],
                               "Particulars":   data.loc[rawRow,9],
                               "DocNo":         data.loc[rawRow,10],
                               "T":             data.loc[rawRow,11],
                               "DocDate":       data.loc[rawRow,12],
                               "PT":            data.loc[rawRow,13],
                               "DueDate":       data.loc[rawRow,14],
                               "Ana1":          data.loc[rawRow,15],
                               "Ana2":          data.loc[rawRow,16],
                               "Ana3":          data.loc[rawRow,17],
                               "Ana4":          data.loc[rawRow,18],
                               "Ana5":          data.loc[rawRow,19],
                               "Amount":        data.loc[rawRow,7] - data.loc[rawRow,8],
                               "FCAmount":      data.loc[rawRow,4]} 
            Count = Count +1
        """ ????
          For Line = (RowInNew - Count) To RowInNew
              NewGL.Range("Y" & Line).Value = Worksheets(w).Range("B" & RowInRaw).Value 'UserID
              NewGL.Range("Z" & Line).Value = Worksheets(w).Range("C" & RowInRaw).Value 'PostDate
              Next Line
        """
        rawRow +=1        
    df = pd.DataFrame.from_dict(table, "index")
    
    return df 

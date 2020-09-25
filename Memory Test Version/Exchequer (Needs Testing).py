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
        if data.loc[rawRow,0].find(',') != -1:
            a = data.loc[rawRow].find(',')
            accountCode = left(data.loc[rawRow], a)
            accountDesc = right(data.loc[rawRow], a - 2)
        elif mid(data.loc[rawRow,6], 4, 1) == '/': 
            table[rawRow] = {  'File Name':     data.loc[rawRow,'filename'],
                               'Entity Name':   data.loc[0,3], #This needs testing
                               "Account":       accountCode,
                               "Account Desc":  accountDesc,
                               "Our Ref":       data.loc[rawRow,0],
                               "Acc No":        data.loc[rawRow,2],
                               "Your Ref":      data.loc[rawRow,3],
                               "Period":        left(data.loc[rawRow,6], 3),
                               "Date":          data.loc[rawRow,7],
                               "Due Date":      data.loc[rawRow,10],
                               "Description":   data.loc[rawRow,11],
                               "Debit":         data.loc[rawRow,14],
                               "Credit":        data.loc[rawRow,16]} #this needs testing
        rawRow +=1        
    return table  


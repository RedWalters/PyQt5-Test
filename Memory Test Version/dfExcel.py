import os
import pandas as pd
from import_file import import_file
from ntpath import basename
from urllib.parse import quote_plus
from sqlalchemy import create_engine, event
import time
import xlrd

HERE = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.abspath(os.path.join(HERE, '..', 'data'))

def chunks(l, n):
    for i in range(0, len(l), n):
         yield l.iloc[i:i+n]

#def write_df_to_sql(frame, table_name, engine, chunk_size):
    #i_chunk = 0
    #for idx, chunk in enumerate(chunks(frame, chunk_size)):
        #print(f'  - Writing File to {table_name} - Chunk {i_chunk} to DB')
        #i_chunk += 1
        #if idx == 0:
         #   if_exists_param = 'replace'
        #else:
         #   if_exists_param = 'append'
        #chunk.to_sql(con=engine, name=table_name, index = False, if_exists='append', schema = 'RAW')


def write_df_to_sql(table_name, engine, chunk_size, frame):

    #conn =  "DRIVER={ODBC Driver 13 for SQL Server};SERVER="+self.server+";DATABASE="+self.database+";Trusted_Connection=yes"
    quoted = quote_plus(engine)
    new_con = 'mssql+pyodbc:///?odbc_connect={}'.format(quoted)
    engine = create_engine(new_con)
            
    @event.listens_for(engine, 'before_cursor_execute')
    def receive_before_cursor_execute(conn, cursor, statement, params, context, executemany):
        if executemany:
            cursor.fast_executemany = True    
    
    i_chunk = 0
    for idx, chunk in enumerate(chunks(frame, chunk_size)):
        print(f'  - Writing File to {table_name} - Chunk {i_chunk} to DB')
        i_chunk += 1
        #if idx == 0:
         #   if_exists_param = 'replace'
        #else:
         #   if_exists_param = 'append'
        chunk.to_sql(con=engine, name=table_name, index = False, if_exists='append', schema = 'RAW')

def methodSelection(filename):

    if os.path.splitext(filename)[1] == '.csv' or os.path.splitext(filename)[1] == '.txt':
        return make_df_from_csv(filename)
    else:
        return make_df_from_excelNoForm(filename)

def saveToCsv(dataFrame, zipped):
    if zipped == True:
        print(f'Saving: ' + os.path.splitext(dataFrame['File Name'].iloc[1])[0])
        dataFrame.to_csv(os.path.splitext(dataFrame['File Name'].iloc[1])[0] +'.csv.gz', sep = '|', index = False, compression='gzip', chunksize = None)
       
    else:
        print(f'Saving: ' + os.path.splitext(dataFrame['File Name'].iloc[1])[0])
        dataFrame.to_csv(os.path.splitext(dataFrame['File Name'].iloc[1])[0] +'.csv', sep = '|', index = False, chunksize = None)        
  

def make_df_from_csv(file_name):
    
    with open(file_name) as f:
        encoding = f.encoding
        
    nrows = 10000
    
    filename = basename(file_name)
    
    if os.path.splitext(filename)[1] == '.csv':
        print(f"CSV File: {filename}")
    elif os.path.splitext(filename)[1] == '.txt':
        print(f"TXT File: {filename}")
    else:
        print(f"File: {filename}")
    
    try:
        chunks = []
        i_chunk = 0
        for chunk in pd.read_csv(file_name, names = list(range(0,25)), chunksize = nrows, encoding = 'utf-16'):
            chunks.append(chunk)
            print(f"  - File {filename} - Chunk {i_chunk} ({len(chunk)} rows)")
            i_chunk += 1
    except:
        chunks = []
        i_chunk = 0
        for chunk in pd.read_csv(file_name, names = list(range(0,25)), chunksize = nrows, encoding = encoding):
            chunks.append(chunk)
            print(f"  - File {filename} - Chunk {i_chunk} ({len(chunk)} rows)")
            i_chunk += 1
            
    print(f'  - File {filename} - Concatenating')
    df_chunks = pd.concat(chunks,  ignore_index = True).fillna('')
    df_chunks['File Name'] = filename
    chunks = []
    
    return df_chunks

def make_df_from_excel(file_name, formatting):
    nrows = 10000
    
    filename = basename(file_name)
    
    print(f"Excel file: {filename}")
    
    impModule = import_file(formatting)
    
    file_path = os.path.abspath(os.path.join(DATA_DIR, file_name))
    xl = pd.ExcelFile(file_path)

    sheets= []
    
    for i in xl.sheet_names:
        print(f"- File {filename} - Worksheet: {i}")
        chunks = []
        i_chunk = 0
        skiprows = 0
        while True:
            df_chunk = xl.parse(i,
                nrows=nrows, skiprows=skiprows, ignore_index = True, header=None)
            skiprows += nrows
            # When there is no data, we know we can break out of the loop.
            if not df_chunk.shape[0]:
                break
            else:
                print(f"  - File {filename} - Chunk {i_chunk} ({df_chunk.shape[0]} rows)")
                chunks.append(df_chunk)
            i_chunk += 1
        sheets.append(impModule.formatting(pd.concat(chunks, ignore_index = True).fillna('')))
        
    print(f'  - File {filename} - Concatenating')
    
    xl.close()
    
    df_chunks = pd.concat(sheets, ignore_index = True)
    df_chunks['File Name'] = filename
    sheets = []
    chunks = []
    
    return df_chunks

def make_df_from_excelDepen(file_name, formatting):
    
    nrows = 10000
    
    filename = basename(file_name)
    
    print(f"Excel file: {filename}")
    
    impModule = import_file(formatting)
    
    file_path = os.path.abspath(os.path.join(DATA_DIR, file_name))
    xl = pd.ExcelFile(file_path)

    sheets= []
    
    for i in xl.sheet_names:
        print(f"- File {filename} - Worksheet: {i}")
        chunks = []
        i_chunk = 0
        # The first row is the header. We have already read it, so we skip it.
        skiprows = 0
        while True:
            df_chunk = xl.parse(i,
                nrows=nrows, skiprows=skiprows, ignore_index = True, header=None)
            skiprows += nrows
            # When there is no data, we know we can break out of the loop.
            if not df_chunk.shape[0]:
                break
            else:
                print(f"  - File {filename} - Chunk {i_chunk} ({df_chunk.shape[0]} rows)")
                chunks.append(df_chunk)
            i_chunk += 1
        sheets.append(pd.concat(chunks, ignore_index = True).fillna(''))

    print(f'  - File {filename} - Concatenating')
    
    xl.close()
    
    df_chunks = impModule.formatting(pd.concat(sheets, ignore_index = True))
    df_chunks['File Name'] = filename
    sheets = []
    chunks = []
    
    return df_chunks

def make_df_from_excelNoForm(file_name):
    nrows = 10000
    
    filename = basename(file_name)
    
    print(f"Excel file: {filename}")
    #starts = time.time()
    #file_path = os.path.abspath(os.path.join(DATA_DIR, file_name))
    
    #xls = xlrd.open_workbook(file_path, on_demand=True)
    #print(xls.sheet_names())
    
    #xl = pd.ExcelFile(file_path, on_demand = 0)
    #print(time.time() - starts)
    # Read the header outside of the loop, so all chunk reads are
    # consistent across all loop iterations.
    
    df_header = pd.read_excel(file_name, nrows=1)

    sheets= []
    chunks = []
    i_chunk = 0
    # The first row is the header. We have already read it, so we skip it.
    skiprows = 1
    while True:
        df_chunk = pd.read_excel(file_name, sheet_name = None, nrows=nrows, skiprows=skiprows, ignore_index = True, header=None)
        skiprows += nrows
        # When there is no data, we know we can break out of the loop.
        if not df_chunk.shape[0]:
            break
        else:
            print(f"  - File {filename} - Chunk {i_chunk} ({df_chunk.shape[0]} rows)")
            chunks.append(df_chunk)
            i_chunk += 1
            sheets.append(pd.concat(chunks, ignore_index = True, sort = False).fillna(''))

    print(f'  - File {filename} - Concatenating')
    
    #xl.close()
    
    df_chunks = pd.concat(sheets, ignore_index = True, sort = False)
    df_chunks['File Name'] = filename
    sheets = []
    chunks = []
    # Rename the columns to concatenate the chunks with the header.
    columns = {i: col for i, col in enumerate(df_header.columns.tolist())}
    df_chunks.rename(columns=columns, inplace=True)
    df = pd.concat([df_header, df_chunks], sort = False)
    df_chunks = df_chunks[0:0]
    return df
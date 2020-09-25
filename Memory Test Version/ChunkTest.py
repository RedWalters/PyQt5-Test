import os
import pandas as pd
import time
from import_file import import_file
from ntpath import basename

HERE = os.path.abspath(os.path.dirname(__file__))
DATA_DIR = os.path.abspath(os.path.join(HERE, '..', 'data'))


def make_df_from_excel(file_name, nrows = 10000, formatting):
    
    impModule = import_file(formatting)
    
    """Read from an Excel file in chunks and make a single DataFrame.

    Parameters
    ----------
    file_name : str
    nrows : int
        Number of rows to read at a time. These Excel files are too big,
        so we can't read all rows in one go.
    """
    file_path = os.path.abspath(os.path.join(DATA_DIR, file_name))
    xl = pd.ExcelFile(file_path)

    # In this case, there was only a single Worksheet in the Workbook.
    sheetname = xl.sheet_names[0]

    # Read the header outside of the loop, so all chunk reads are
    # consistent across all loop iterations.
    #df_header = pd.read_excel(file_path, sheet_name=sheetname, nrows=1)
    print(f"Excel file: {file_name} (worksheet: {sheetname})")
    sheets= []
    
    for i in xl.sheet_names:
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
                print(f"  - chunk {i_chunk} ({df_chunk.shape[0]} rows)")
                chunks.append(df_chunk)
            i_chunk += 1
        sheets.append(impModule.formatting(pd.concat(chunks, ignore_index = True).fillna('')))

    df_chunks = pd.concat(sheets, ignore_index = True)
    df_chunks['File Name'] = basename(file_name)
    # Rename the columns to concatenate the chunks with the header.
    #columns = {i: col for i, col in enumerate(df_header.columns.tolist())}
    #df_chunks.rename(columns=columns, inplace=True)
    #df = pd.concat([df_header, df_chunks])
    return df_chunks


if __name__ == '__main__':
   
    file = 'K:\\A & A\\Cardiff\\Audit\\Clients\\Open\\S\\Spotlight\\2. Staff Folders\\JWalters\\__Python\\Test Data and VBA\\Baybridge Housing\\Oct 2018 GL details - Yardi Trust.xlsx'
    frmat = 'K:\\A & A\\Cardiff\\Audit\\Clients\\Open\\S\\Spotlight\\2. Staff Folders\\JWalters\\__Python\\Formatting Scripts\\YardiTrust.py'

    s = time.time()
    
    df = make_df_from_excel(file, nrows=10000, formatting = frmat)

    print(df)
    print(time.time() - s)
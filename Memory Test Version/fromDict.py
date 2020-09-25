import pandas as pd
def toDataFrame(dic):
    try:
        df = pd.DataFrame.from_dict(dic, "index")
        return df
    except:
        pass
    
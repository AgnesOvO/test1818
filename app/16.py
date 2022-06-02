# 指定歸0
import pandas as pd
from pandas import DataFrame

a = pd.read_excel('D:/Python/e1.xlsx')
df = pd.DataFrame(a)
List= df['總成績'].tolist()  
print(List)

n=-1
for i in List:
    n=n+1  
    df.at[n, "總成績"] = 0 
    df = DataFrame(df) 
    DataFrame(df).to_excel('D:/Python/e1.xlsx',sheet_name='Sheet1', index=False, header=True)

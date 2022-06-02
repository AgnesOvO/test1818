# 指定歸0
import pandas as pd
from pandas import DataFrame


s = str(input("輸入："))
a = pd.read_excel(s)
df = pd.DataFrame(a)

sp = str(input("請輸入要修改欄位："))

List= df[sp].tolist()  

n=-1
for i in List:
    n=n+1  
    df.at[n, sp] = 0 
    df = DataFrame(df) 
    DataFrame(df).to_excel('C:/Users/建良/Desktop/co/e1.xlsx',sheet_name='Sheet1', index=False, header=True)

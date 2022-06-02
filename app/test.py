import pandas as pd 
import numpy as np 
import matplotlib.pyplot as plt 
from pandas import DataFrame 
from collections import Counter

@app.route("/test", methods=["GET", "POST"]) #
def test():
    path = 't1.txt' 
    f = open(path, 'w') 

    a=pd.read_excel("e.xlsx",usecols="D") 
    df = pd.DataFrame(a) 
    List= df['總成績'].tolist()

    recounted = Counter(List) #統計資料出現次數
    #print(recounted) #檢查 
    b1=df.groupby('總成績').size() 
    #print(b1) #檢查 
    A=max(b1) 
    print("最大次數",A) #找出最大值的點為多少

    a=np.array(List) #將資料改成陣列 (分數) 
    a1=np.unique(a) #print(a1) 
    b=np.array(b1) #將資料轉成陣列 (次數) 
    c=np.vstack((b,a1)) #合併變成二維陣列 
    print(c) #檢查 
    n=-1 
    p=[] 
    for row in c: 
        for col in row: 
            n=n+1 
            if(col==A): #判斷次數如果是跟最大一樣 
                print(c[1][n]) 
                p.append(c[1][n])
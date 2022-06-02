from pymongo import MongoClient
import pandas as pd
import certifi #為了解決連線到SSL的問題
import sqlalchemy
import json
import csv
import codecs

conn = MongoClient("mongodb+srv://root:root159258@cluster0.oe4sl.mongodb.net/myFirstDatabase?retryWrites=true&w=majority", tlsCAFile=certifi.where())
db = conn.test #選擇操作 test 資料庫
collection = db.test2 #選擇操作 users 集合

# test if connection success
collection.stats  # 如果沒有error，你就連線成功了。

#dk = pd.read_excel("/app/app/e1.xlsx", index_col=0)
#data_csv = dk.to_csv('new.csv', encoding='utf-8-sig')

#df = pd.read_csv("new.csv")
#data = df.to_dict(orient = "records")
#collection.insert_many(data)

#print("You did it") 

# Importing the whole thing back again! 

#data_from_db = db.user_info_table.find({},{'_id':0})
#pd.DataFrame.from_dict(data_from_db).head()

# Query by FIRST_NAME = Alex, and Export from MongoDB to Python:

data_from_db_ = db.test2.find({'Name':'Alex'},{'_id':0})
name_alex=list(data_from_db_)
data = pd.DataFrame.from_dict(name_alex)
print(data)

print("You did it")
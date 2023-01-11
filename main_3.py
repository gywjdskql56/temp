import pandas as pd
from pymongo import MongoClient
import os
client = MongoClient('mongodb+srv://hyojeong_kim:5915@cluster0.nfjxm.mongodb.net/?retryWrites=true&w=majority')
db = client.mydb
col = db.members

files = sorted(os.listdir('crawl_data/'), reverse=True)
for file in files:
    if '_saved' not in file:
        print(file)
        try:
            df =pd.read_excel('crawl_data/'+file,index_col=0)
            print(len(df))
        except:
            os.remove('crawl_data/'+file)
            continue
        if len(df)<1:
            os.remove('crawl_data/' + file)
            continue
        else:
            for idx in df.index:
                row = df.loc[idx].to_dict()
                region = file.split('_')[1].replace('.xlsx', '')
                row['지역'] = region
                col.insert_one(row)
            df.to_excel('crawl_data/'+file.replace('.xlsx', '_saved.xlsx'))
            os.remove('crawl_data/' + file)

print(1)
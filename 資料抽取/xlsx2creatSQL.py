import pandas as pd
import pymysql
from sqlalchemy import create_engine
import os

import sys
# 讀取 Excel 檔案中的資料
if getattr(sys, 'frozen', False):
    target_excel = os.path.join(os.path.dirname(sys.executable), '2.解析結果輸出於此', '資訊列表.xlsx')
else:
    target_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), '2.解析結果輸出於此', '資訊列表.xlsx')

df = pd.read_excel(target_excel)
cases = []
try:
    # 資料庫連線設定
    host = 'localhost'
    port = 3306
    user = sys.argv[1] if len(sys.argv) >= 2 else input("請輸入使用者名稱：")
    password = sys.argv[2] if len(sys.argv) >= 2 else input("請輸入密碼：")
    database = 'mediation'

    # 連接到 MySQL 資料庫
    conn = pymysql.connect(host=host, port=port, user=user, password=password)
    # 創建資料庫
    cursor = conn.cursor()

    cursor.execute(f"CREATE DATABASE IF NOT EXISTS {database}")

    # 建立 SQLAlchemy 引擎
    engine = create_engine(f'mysql+pymysql://{user}:{password}@{host}:{port}/{database}')

    # 將 DataFrame 中的資料存入 MySQL 資料庫中
    df.to_sql('med_case', engine, if_exists='replace', index=False)

    # 關閉資料庫連線
    conn.close()

    # df.to_sql(case_table, engine, if_exists='replace', index=False,
    #           dtype={'收件編號': sqlalchemy.types.VARCHAR(length=255),
    #                  '收件日期': sqlalchemy.types.Date(),
    #                  '案號': sqlalchemy.types.VARCHAR(length=255),
    #                  '聲請人': sqlalchemy.types.VARCHAR(length=255),
    #                  '對造人': sqlalchemy.types.VARCHAR(length=255),
    #                  '事由': sqlalchemy.types.VARCHAR(length=255),
    #                  '過程摘要': sqlalchemy.types.TEXT(),
    #                  '調解結果': sqlalchemy.types.TEXT(),
    #                  '委員': sqlalchemy.types.VARCHAR(length=255),
    #                  '報院審查日期': sqlalchemy.types.Date(),
    #                  '報院審查文號': sqlalchemy.types.VARCHAR(length=255),
    #                  '法院發還日期': sqlalchemy.types.Date(),
    #                  '法院發還文號': sqlalchemy.types.VARCHAR(length=255),
    #                  '法院審查結果': sqlalchemy.types.VARCHAR(length=255),
    #                  '轉借機關': sqlalchemy.types.VARCHAR(length=255),
    #                  '轉介人': sqlalchemy.types.VARCHAR(length=255),
    #                  '歸檔日期': sqlalchemy.types.Date()})
except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()
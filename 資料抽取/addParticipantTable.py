import pandas as pd
import pymysql

import os
def split2List(str_names):
  result = []
  for name in str_names.split('、'):
     for spComa in name.split(','):
        result.append(spComa.split('(')[0].split('[')[0].strip())
  return result

class Case:
    def __init__(self):
        self.applicants = []           #聲請人
        self.opponents = []            #對造人
        self.case_number = ""

import sys
try:
    if getattr(sys, 'frozen', False):
        absPath = os.path.dirname(sys.executable)
    else:
        absPath = os.path.dirname(os.path.abspath(__file__))

    # 讀取 Excel 檔案中的資料
    target_excel = os.path.join(absPath, '2.解析結果輸出於此', '資訊列表.xlsx')
    df = pd.read_excel(target_excel)
    cases = []

    for index, row in df.iterrows():
        case = Case()
        case.case_number = row['案號']
        case.applicants = split2List(row['聲請人'])
        case.opponents = split2List(row['對造人'])
        cases.append(case)

    # 資料庫連線設定
    host = 'localhost'
    port = 3306
    user = sys.argv[1] if len(sys.argv) >= 2 else input("請輸入使用者名稱：")
    password = sys.argv[2] if len(sys.argv) >= 2 else input("請輸入密碼：")
    database = 'mediation'
    table = 'participant'
    # 連接到 MySQL 資料庫
    conn = pymysql.connect(host=host, port=port, user=user, password=password)
    cursor = conn.cursor()
    cursor.execute("CREATE DATABASE IF NOT EXISTS mediation")
    conn = pymysql.connect(host=host, port=port, user=user, password=password,database = database)
    cursor = conn.cursor()


    # 將 df 中出現的所有人名存入 participant table
    names = []

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {database}.{table}(
            `participant_name` varchar(100),
            `case_num` varchar(100),
            `address` VARCHAR(70));
    """)

    #放進table participant的name中
    for case in cases:
        for name in set(case.applicants + case.opponents):
            cursor.execute(f"SELECT * FROM {database}.{table} WHERE participant_name = '{name}' AND case_num = '{case.case_number}'")
            existing_data = cursor.fetchone()
        print(name)
        if not existing_data:
            print(f"新增{name}")
            cursor.execute(f"INSERT INTO {database}.{table} (`participant_name`,`case_num`) VALUES (%s,%s)", (name,case.case_number))

    # Commit the changes and close the database connection
    conn.commit() 
    # 關閉資料庫連線
    conn.close()
except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()
    

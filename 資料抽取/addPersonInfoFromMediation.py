import docx
import os
import re

class human:
    def __init__(self, name = "", address = ""):
        self.name = name
        self.address = address
class Case:
    def __init__(self, file_path):
        self.file_path = file_path
        self.number = None             #收件編號
        self.case_number = None        #案號
        self.applicants = []           #聲請人
        self.opponents = []            #對造人

    def parse_meditation(self): #調解書
        # 打開Word文件
        doc = docx.Document(self.file_path)
        # 提取表格中的內容
        tables = doc.tables
        for table in tables:
            for row in table.rows:
                try:
                    row_data = []
                    for cell in row.cells:
                        row_data.append(cell.text)
                    for i, data in enumerate(row_data): # 消除空白
                        row_data[i] = data.replace(" ", "").replace("\n", "").replace("\t", "")
                    for cell in row.cells:
                        row_data.append(cell.text)
                    # print(row_data)
                    # 解析表格資料
                    print(row_data)
                    if re.search(r'^聲請人\d*$', row_data[0]):
                        self.applicants.append(human(name=row_data[1], address=row_data[6]))
                        print(self.applicants[0].name)
                    if re.search(r'^對造人\d*$', row_data[0]):
                         self.opponents.append(human(name=row_data[1], address=row_data[6]))

                    for datas in row_data:
                        match = re.search(r'收件編號：(\d+)', datas)  # 收件編號
                        if match and self.number is None:
                            self.number = match.group(1).strip() 
                        match = re.search(r'案號(\d+年\S*調字第\d+號)\s*', datas.replace(" ","")) #案號
                        if match and self.case_number is None:
                            self.case_number = match.group(1).strip()
                    # print(row_data)
                except IndexError:
                  print(f"Error: row_data index out of range. row_data: {row_data}")

def update_case_data(existing_case, new_case):
    if existing_case == new_case:
        return existing_case
    updated_case = Case(existing_case.file_path)
    updated_case.number = existing_case.number if existing_case.number is not None else new_case.number
    updated_case.case_number = existing_case.case_number if existing_case.case_number is not None else new_case.case_number
    updated_case.applicants = existing_case.applicants if existing_case.applicants else new_case.applicants
    updated_case.opponents = existing_case.opponents if existing_case.opponents else new_case.opponents
    if existing_case.applicants and new_case.applicants:
        updated_case.applicants = list(set(existing_case.applicants + new_case.applicants))
    if existing_case.opponents and new_case.opponents:
        updated_case.opponents = list(set(existing_case.opponents + new_case.opponents))
    return updated_case

import sys
cases = {}
def search_and_parse_files(pattern, parse_func):
    global cases
    if getattr(sys, 'frozen', False):
      path = os.path.join(os.path.dirname(sys.executable))
    else:
      path = os.path.join(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(os.path.dirname(path),'格式轉換','請將欲處理文件備份後放入此處')
      
    files = [os.path.join(root, file) for root, dirs, files in os.walk(path)
             for file in files if re.search(pattern, file) and os.path.splitext(file)[1] == ".docx"]
    for file in files:
        case = Case(file)
        parse_func(case)
        print('dealing file:')
        print(case.file_path)
        
        if case.case_number is not None and case.case_number:
            if str(case.case_number) in cases:
                existing_case = cases[str(case.case_number)]
                updated_case = update_case_data(existing_case, case)
                cases[str(case.case_number)] = updated_case
            else:
                cases[str(case.case_number)] = case
# 搜尋並解析調解書
search_and_parse_files(r"^(?!.*送達)(?!.*聲請)(?:\d+\.[\s\S]*調解書[\s\S]*|調解書(?:\([\w\s]+\))?)\.docx$", Case.parse_meditation)

import pymysql
from sqlalchemy import create_engine
try:
    # 資料庫連線設定
    host = 'localhost'
    port = 3306
    user = 'root'
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
    # 建立 SQLAlchemy 引擎
    engine = create_engine(f'mysql+pymysql://{user}:{password}@{host}:{port}/{database}')

    cursor.execute(f"""
        CREATE TABLE IF NOT EXISTS {database}.{table}(
            `participant_name` varchar(100),
            `case_num` varchar(100),
            `address` VARCHAR(70));
    """)

    import json
    # 讀取舊的修正資料
    if getattr(sys, 'frozen', False):
        dirpath = os.path.join(os.path.dirname(sys.executable))
    else:
        dirpath = os.path.join(os.path.dirname(os.path.abspath(__file__)))

    try:
        with open(os.path.join(dirpath,"fix_data.txt"), "r",  encoding="utf-8") as file:
            fix = json.load(file)
    except FileNotFoundError:
        # 如果文件不存在，創建一個空的 fix 字典
        fix = {}

    for case_number, case in cases.items():
        unique_people = set(case.applicants + case.opponents)
        for people in unique_people:
            name = people.name.strip()
            # 檢查異常資料
            if len(name) > 4 or any(c in name for c in ['(', ')', '\\', '/', '[', ']']):
                if name in fix:
                    print(f"異常資料:\n{name}\n套用先前修正:\n{fix[name]}")
                    name = fix[name]
                else:
                    question = input(f"異常資料:\n{name}\n是否修正?(是:輸入新名稱後按enter/否:直接按enter)\n")
                    if question :
                        fix[name] = question
                        name = question
                    else:  # 套用不修正
                        fix[name] = name
            address = people.address.strip()
            if any(c in address for c in ['(', ')', '\\', '/', '[', ']']):
                if address in fix:
                    print(f"異常資料:\n{address}\n套用先前修正:\n{fix[address]}")
                    address = fix[address]
                else:
                    question = input(f"異常資料:\n{address}\n是否修正?(是:輸入新名稱後按enter/否:直接按enter)\n")
                    if question :                   
                        fix[address] = question  
                        address = question  
                    else:  # 套用不修正
                       fix[address] = address
            #檢查空白資料            
            if not name:
                question = input(f"資料不齊全:\n案號 {case_number} 地址: {address} 姓名為空白，是否補齊資料?(是:輸入新名稱後按enter/否:直接按enter)\n")
                if question :                   
                    fix[name] = question  
                    name = question  
            if not address:
                question = input(f"資料不齊全:\n案號 {case_number} 姓名: {name} 地址為空白，是否補齊資料?(是:輸入新名稱後按enter/否:直接按enter)\n")
                if question :                   
                    fix[address] = question  
                    address = question            

            # Check if the person already exists in participant table
            select_sql = f"SELECT * FROM {database}.{table} WHERE participant_name = %s AND case_num = %s"
            cursor.execute(select_sql, ( name, case_number))
            existing_data = cursor.fetchone()
            if existing_data:
                update_sql = f"UPDATE {database}.{table} SET address = %s WHERE participant_name = %s AND case_num = %s"
                cursor.execute(update_sql, ( address, name, case_number))
            else:
                # Insert a new row if person does not exist
                insert_sql = f"INSERT INTO {database}.{table} (participant_name, case_num, address) VALUES (%s, %s, %s)"
                cursor.execute(insert_sql, ( name, case_number, address))

    with open(os.path.join(dirpath,"fix_data.txt"), "w", encoding="utf-8") as file:
        fix = json.dump(fix,file,ensure_ascii=False)

    # Commit the changes and close the database connection
    conn.commit()
    conn.close()
except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()
    

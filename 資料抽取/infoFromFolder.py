import sys
arg_year = None
try:
  arg_year = sys.argv[1]
  print("年份: ",arg_year)
except:
  arg_year = input("請輸入民國年份:")
  print("年份: ",arg_year)

import os
import pandas as pd

class Case:
    def __init__(self):
        # self.file_path = file_path
        self.number = None             #收件編號
        self.time = None               #收件日期
        self.case_number = None        #案號
        self.applicants = ""           #聲請人
        self.opponents = ""            #對造人
        self.case_reason = None        #事由
        self.summary = ""           #過程摘要

def isTarget(dir_name) :
  if not "調解筆錄" in dir_name:
      return False
  if not "聲請人" in dir_name:
      return False
  if not "對造人" in dir_name:
      return False
  if not "開調解時間" in dir_name:
      return False
  return True

if getattr(sys, 'frozen', False):
    absPath = os.path.dirname(sys.executable)
else:
    absPath = os.path.dirname(os.path.abspath(__file__))

path = os.path.join(os.path.dirname(absPath),'格式轉換','請將欲處理文件備份後放入此處')
# 開始尋找目標資料夾
folder_list = []
for root, dirs, files in os.walk(path):  # 在存在的資料夾中進行搜尋
    for dir_name in dirs:
        if arg_year.strip() + '年' in dir_name.replace(" ","") and isTarget(dir_name):
            folder_list.append(os.path.join(root, dir_name))  # 加入找到的資料夾到清單中

new_data = []
# 將每個目標資料夾的名稱剖析並寫入 Excel
for i, folder_path in enumerate(folder_list):
    print(folder_path)
    case = Case()
    try:
        folder_name = os.path.basename(folder_path).replace(' ','')
        if '調解筆錄' in folder_name and '聲請人' in folder_name and '對造人' in  folder_name and '開調解時間' in folder_name:
          # 取得收件編號
          case.number = folder_name.split("調解筆錄")[1].split("(聲請人")[0]
          # 取得聲請人和對造人
          case.applicants = folder_name.split("聲請人")[1].split(")(對造人")[0]
          if ')(收' in folder_name:
            case.opponents = folder_name.split("對造人")[1].split(")(收")[0]

          # 取得案件編號
          if '收：' in folder_name:
            case.case_number = folder_name.split('收：')[1].split(')(開調解時間')[0]
          elif '收:' in folder_name:
            case.case_number = folder_name.split('收:')[1].split(')(開調解時間')[0]

          print(case.case_number)
          if case.case_number != None :
            if '刑' in case.case_number:
              case.case_number = case.case_number.split('刑')[0] + '年' + '刑調字第' + case.case_number.split('刑')[1].split(')')[0] +'號'
            elif '民' in case.case_number:
              case.case_number = case.case_number.split('民')[0] + '年' + '民調字第' + case.case_number.split('民')[1].split(')')[0] +'號'

          # 取得事由
          if '分)-(' in folder_name:
            case.case_reason = folder_name.split('分)-(')[1].split(')')[0]
          new_data.append({
              "收件編號": case.number,
              "收件日期": case.time,
              "案號": case.case_number,
              "聲請人": "".join(case.applicants),
              "對造人": "".join(case.opponents),
              "事由": case.case_reason,
              "過程摘要": case.summary,
          })
    except Exception as e:
        print(f"Error occurred: {str(e)}")

# 輸出檔案的路徑
output_path = os.path.join(absPath, '2.解析結果輸出於此', '資訊列表.xlsx')

if os.path.isfile(output_path):
    # 讀取現有的資料
    existing_df = pd.read_excel(output_path, dtype={"收件編號": str})
    df = pd.DataFrame(new_data, columns=["收件編號", "收件日期", "案號", "聲請人", "對造人", "事由", "過程摘要"])
    # 將 DataFrame 存成 xlsx 檔案
    with pd.ExcelWriter(output_path) as writer:
        pd.concat([existing_df, df], ignore_index=True).to_excel(writer, index=False)
else:
    df = pd.DataFrame(new_data, columns=["收件編號", "收件日期", "案號", "聲請人", "對造人", "事由", "過程摘要"])
    # 將 DataFrame 存成 xlsx 檔案
    with pd.ExcelWriter(output_path) as writer:
        df.to_excel(writer, index=False)


# 儲存 Excel 檔案
# excel_file_path = os.path.join(absPath, '2.解析結果輸出於此', 'infofromFolder_資訊列表.xlsx')

import os
import sys
import pandas as pd
import json

year = input("請輸入年份:")
current_dir = os.path.dirname(os.path.abspath(__file__))
try:
    # 讀取抽取前的Excel檔案
    df_old = pd.read_excel(os.path.join(current_dir,'2.解析結果輸出於此','資訊列表.xlsx'))
    # 获取当前脚本文件的绝对路径
    if getattr(sys, 'frozen', False):
        current_dir = os.path.dirname(sys.executable)
        os.remove(os.path.join(current_dir,'2.解析結果輸出於此','資訊列表.xlsx'))
        # 0. 执行 getInfo.py，不傳入參數
        os.system(f"{os.path.join(current_dir, 'getInfo.exe')} {year} ")
        # 1. 执行 getInfo.py，并传入参数"解析聲請調解書"
        os.system(f"{os.path.join(current_dir, 'getInfo.exe')} {year} '解析聲請調解書'")
        # 2. 执行 getInfo.py，并传入参数"解析調解筆錄"
        os.system(f"{os.path.join(current_dir, 'getInfo.exe')} {year} '解析調解筆錄'")
        # 3. 执行 getInfo.py，并传入参数"解析調解書"
        os.system(f"{os.path.join(current_dir, 'getInfo.exe')} {year} '解析調解書'")
        # 4. 执行 infoFromFolder.py
        os.system(f"{os.path.join(current_dir, 'infoFromFolder.exe')} {year}")
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        os.remove(os.path.join(current_dir,'2.解析結果輸出於此','資訊列表.xlsx'))
        # 0. 执行 getInfo.py，不傳入參數
        os.system(f"python {os.path.join(current_dir, 'getInfo.py')} {year} ")
        # 1. 执行 getInfo.py，并传入参数"解析聲請調解書"
        os.system(f"python {os.path.join(current_dir, 'getInfo.py')} {year} '解析聲請調解書'")
        # 2. 执行 getInfo.py，并传入参数"解析調解筆錄"
        os.system(f"python {os.path.join(current_dir, 'getInfo.py')} {year} '解析調解筆錄'")
        # 3. 执行 getInfo.py，并传入参数"解析調解書"
        os.system(f"python {os.path.join(current_dir, 'getInfo.py')} {year} '解析調解書' ")
        # 4. 执行 infoFromFolder.py
        os.system(f"python {os.path.join(current_dir, 'infoFromFolder.py')} {year}")


    # 讀取抽取後的Excel檔案
    df_new = pd.read_excel(os.path.join(current_dir,'2.解析結果輸出於此','資訊列表.xlsx'))
    # 檢查檔案是否存在
    json_file_path = os.path.join(current_dir,'old_cases.json')

    print('分析結束，正在過濾舊案')
    try:
        # 讀取 JSON 檔案
        with open(json_file_path, 'r', encoding='utf-8') as f:
            json_data = json.load(f)
        for old_case_num in json_data:
            while old_case_num in df_new['案號'].to_list():
                #創建一個布林遮罩
                mask = df_new['案號'].isin(json_data)
                #然後去除遮罩項目
                df_new = df_new[~mask]

    except FileNotFoundError:
        print('不存在已紀錄之舊案')
        
    except Exception as e:
        print('發生錯誤:', e)

    df_final = pd.concat([df_old, df_new], ignore_index=True)
    df_final.to_excel(os.path.join(current_dir,'2.解析結果輸出於此','資訊列表.xlsx'), index=False)

    # 提取指定欄位的數據
    old_case_data = list(set(df_final['案號'].tolist())) #確保list不重複
    # 將數據轉換為 JSON 格式
    json_data = json.dumps(old_case_data)


    if not os.path.exists(json_file_path):
        # 如果檔案不存在，則創建檔案
        with open(json_file_path, 'w', encoding='utf-8') as f:
            pass

    # 儲存 JSON 數據到檔案
    with open(json_file_path, 'w', encoding='utf-8') as f:
        json.dump(old_case_data, f,ensure_ascii=False)

except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()

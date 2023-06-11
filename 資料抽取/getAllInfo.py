import os
import sys

year = input("請輸入年份:")

try:
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
except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()

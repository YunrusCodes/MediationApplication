import os
import pandas as pd

# 輸出檔案的路徑
import sys
if getattr(sys, 'frozen', False):
    dirFolder = os.path.dirname(sys.executable)
else:
    dirFolder = os.path.dirname(os.path.abspath(__file__))
path = os.path.join( dirFolder, '紙本資料整理.xlsx')

# 將日期從西元年月日轉換成民國年月日
def convert_date(date_str):
    try:
      year, month, day = date_str.split('.')
      return year + '年' + month + '月' + day + '日'
    except:
       print(f"faild to convert:{date_str}")
try:
    if os.path.isfile(path):
        str_cols = ['案號' ,'收件編號' ,'轉介單位' ,'轉介人' ,'委任人' ,'受任人' ,'報院審查日期' ,'報院審查文號-1'
                    ,'報院審查文號-2' ,'法院發還文號-1' ,'法院發還文號-2' ,'調解委員' ,'主席' ,'法院審查結果','法院發還日期','最後送達日期']
        
        int_cols = ['聲請調解書' ,'委任書或請求權讓與同意書' ,'調解案件轉介單〈函〉' ,'身分證或戶籍謄本資料〈營利事業登記證影本〉', '調解事件處理單'
                    ,'證據文件資料' ,'調解期日通知書副本' ,'調解期日通知書送達證書', '調解筆錄'
                    ,'函報法院審核函副本' ,'法院核定函〈補正函〉'  ,'調解書' 
                    ,'檢送調解書予當事人函副本', '調解書送達證書', '發給調解不成立證明聲請書' 
                    ,'調解不成立證明書', '刑事事件調解不成立移送偵查聲請書' 
                    ,'刑事事件調解不成立移送偵查書副本', '調解撤回聲請書', '其它(底頁)' ]
        df = pd.read_excel(path, header=0
                        , converters={col: str for col in str_cols})
        # 補齊案號名稱
        df['案號'] = df['案號'].apply(lambda x: x[:3] + '年' + x[3] + '調字第' + x[4:] + '號')
        # 新增報院審查文號欄位
        df['報院審查文號'] = df['報院審查文號-1'] + '第' + df['報院審查文號-2'] + '號'
        # 刪除報院審查文號-1和報院審查文號-2欄位
        df = df.drop(columns=['報院審查文號-1', '報院審查文號-2']) 
        
        # 新增法院發還文號欄位
        df['法院發還文號'] = df['法院發還文號-1'] + '第' + df['法院發還文號-2'] + '號'
        # 刪除法院發還文號-1和法院發還文號-2欄位
        df = df.drop(columns=['法院發還文號-1', '法院發還文號-2']) 

        # 將報院審查日期和法院發還日期格式轉換成民國年月日
        df['報院審查日期'] = df['報院審查日期'].apply(convert_date)
        df['法院發還日期'] = df['法院發還日期'].apply(convert_date)
        df['最後送達日期'] = df['最後送達日期'].apply(convert_date)

    import pymysql
    from sqlalchemy import create_engine
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

    # # 將int_cols的欄位轉換為整數並將None值填充為0
    df[int_cols] = df[int_cols].fillna(0).astype(int)
    df.to_sql('contents', engine, if_exists='replace', index=False)
except Exception as e:
   input(f"發生錯誤:{e}\n請按任一鍵離開")
   sys.exit()
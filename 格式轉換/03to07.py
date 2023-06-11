import os
import sys

try:
  docCount = 0
  xlsCount = 0
  absPath = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
  print(os.path.join(absPath,'請將欲處理文件備份後放入此處'))
  for root, dir, files in os.walk(os.path.join(absPath,'請將欲處理文件備份後放入此處')):
      for file in files:
          if file.endswith('doc'):
              docCount += 1
          elif file.endswith('xls'):
              xlsCount += 1
              
  if getattr(sys, 'frozen', False):
      # 以多執行緒操作 doc to docx
      if docCount > 250 :
        os.system(os.path.join(absPath, 'doc2docx_muti.exe'))
      # 以多執行緒操作 xls to xlsx
      if xlsCount > 250:
        os.system(os.path.join(absPath, 'xls2xlsx_muti.exe'))
      # 以主執行緒操作 doc to docx
      os.system(os.path.join(absPath, 'doc2docx.exe'))
      # 以主執行緒操作 xls to xlsx
      os.system(os.path.join(absPath, 'xls2xlsx.exe'))
  else:
      if docCount > 250:
        os.system("python " + os.path.join(absPath, 'doc2docx_muti.py'))
      # 以多執行緒操作 xls to xlsx
      if xlsCount > 250 :
        os.system("python " + os.path.join(absPath, 'xls2xlsx_muti.py'))
      # 以主執行緒操作 doc to docx
      os.system("python " + os.path.join(absPath, 'doc2docx.py'))
      # 以主執行緒操作 xls to xlsx
      os.system("python " + os.path.join(absPath, 'xls2xlsx.py'))    

  if docCount > 250 or xlsCount > 250 :
      restart_explorer = input("程式執行完畢，建議重啟檔案總管，是否進行重啟？(y/n) ")
      if restart_explorer.lower() == 'y':
          os.system("taskkill /f /im explorer.exe")
          os.system("start explorer.exe")
  else:
    input("程式執行完畢，輸入任意鍵結束")

except Exception as e:
   input(f"發生錯誤: {e}\n輸入任意鍵繼續")
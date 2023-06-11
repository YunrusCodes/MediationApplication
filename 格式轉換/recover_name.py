import os
import shutil

def removeDir(path):
      if os.path.isfile(path):
        os.remove(path)
      else:
         shutil.rmtree(path)

import sys
# 讀取檔案
if getattr(sys, 'frozen', False):
    absPath = os.path.join(os.path.dirname(sys.executable))
else:
    absPath = os.path.dirname(os.path.abspath(__file__))

recover_path = os.path.join(os.path.join( absPath,'請將欲處理文件備份後放入此處'), "recover.txt")

with open(recover_path,  encoding='utf-8') as f:
    content = f.readlines()

# 取出目錄和更改後名字
directories = []
for line in content:
    if '->' in line:
        directories.append(line.strip().split('->'))

failed_to_recovered = []
# 印出結果
print('復原檔/目錄名:')
for directory in reversed(directories):
    print(directory[1] + ' -> ' + directory[0])
    try:
      if directory[0].endswith('.doc') or directory[0].endswith('.xls'):
        # 成功轉換的檔案
        try:
          os.rename(directory[1]+'x',directory[0]+'x')
        except FileNotFoundError:
           os.rename(directory[1],directory[0])
        except FileExistsError:
           print(directory[1]+'x' + ' -> ' + directory[0]+'x')
           print(f"已存在 {directory[0]+'x'}，刪除檔案不覆蓋")
           removeDir(directory[1]+'x')
      else:
        os.rename(directory[1],directory[0])

    except FileNotFoundError:
      print(directory[1] + ' -> ' + directory[0])
      print(f"不存在{directory[1]}，不予處理")
    except FileExistsError:
      print(directory[1] + ' -> ' + directory[0])
      print(f"已存在 {directory[0]}，刪除檔案不覆蓋")
      removeDir(directory[1])
    except Exception as e:
      failed_to_recovered.append((directory[0],directory[1]))
      print(directory[1] + ' -> ' + directory[0])
      print('復原失敗，請手動恢復')
      print(f"An error occurred: {e}")

# 將文件內容清空
with open(recover_path, 'w',  encoding='utf-8') as f:
    f.truncate(0)

failed_to_recover_path = os.path.join(os.path.join( absPath,'請將欲處理文件備份後放入此處'), "failed_to_recover.txt")

if not os.path.exists(failed_to_recover_path):
    open(failed_to_recover_path, 'w').close()
# 將失敗的目錄/檔案寫入failed_to_recover.txt

with open(failed_to_recover_path, 'a',  encoding='utf-8') as f:
    for item in failed_to_recovered:
        f.write(f"{item[0]}->{item[1]}\n")
if failed_to_recovered:
  print("名稱復原失敗的檔案清單已寫入 failed_to_recover.txt")
else:
  print("復原清單已無內容")

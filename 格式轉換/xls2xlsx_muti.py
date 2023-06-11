import os
import xlwings as xw
import threading
import queue
import pythoncom

folder_serial_num = 1
file_serial_num = 1
deal_Queue = queue.Queue()
gThread_num = 1
gFileNum = 1
def convert_Xls():
    global deal_Queue
    global gThread_num
    global gFileNum
    global gAllFileNum
    
    thred_num = gThread_num
    gThread_num = gThread_num + 1
    pythoncom.CoInitialize()
    app = xw.App(visible=False, add_book=False)
    while not deal_Queue.empty():
      file = deal_Queue.get()
      
      try:
        workbook = app.books.open(file)
        new_file = file + 'x'
        print('執行緒' + str(thred_num) + '- 正在操作: '+ os.path.basename(file),"->",os.path.basename(new_file))
        workbook.save(new_file)
        os.remove(file)
        workbook.close()
        print('執行緒' + str(thred_num) + '- 操作結束: ' + str(gFileNum) + '/' + str(gAllFileNum) )    
      except Exception as e:
        print(f"Failed to open file {file}: {e}")
        print('執行緒' + str(thred_num) + "- 操作失敗")
        failed_files.append(file)
      gFileNum = gFileNum - 1
    app.quit()
    pythoncom.CoUninitialize()

def has_special_char(string):
    special_chars = ['<', '>', '?', '[', ']', ':', '|', '*', '\n•']
    for char in special_chars:
        if char in string:
            return True
    return False

def Queue_files(path, q):
    for root, dirs, filenames in os.walk(path):
        for filename in filenames:
            if filename.endswith(".xls") and not filename.startswith(".~"):
               q.put(os.path.join(root, filename))

def substitute_folder_name(folder_name):
    global folder_serial_num
    if has_special_char(folder_name) :
        new_name = 'sub_folder' + str(folder_serial_num)
        folder_serial_num = folder_serial_num + 1
        return new_name
    else :
        return folder_name

def substitute_file_name(file_name):
    global file_serial_num
    if has_special_char(file_name) :
        new_name = 'sub_file' + str(file_serial_num)
        file_serial_num = file_serial_num + 1
        return new_name
    else :
        return file_name

import sys
if getattr(sys, 'frozen', False):
    path = os.path.join( os.path.dirname(sys.executable),'請將欲處理文件備份後放入此處')
else:
    path = os.path.join( os.path.dirname(os.path.abspath(__file__)),'請將欲處理文件備份後放入此處')

folder_to_be_renamed = [] # 儲存需要修改名稱的目錄名稱
file_to_be_renamed = []  # 儲存需要修改名稱的檔案名稱
max_file_number = 0
max_folder_number = 0
# 是否存在failed_to_recover.txt?若存在，找出於failed_to_recover.txt中，被以sub_命名，編號最大的數字，相加於serial_num避免衝突
faild_to_recover_path = os.path.join(path, "failed_to_recover.txt")
if os.path.exists(faild_to_recover_path):
    with open(faild_to_recover_path, encoding='utf-8') as f:
        for line in f:
            if 'sub_' in line:
                print(line)
                lastPart = line.split('sub_')[-1].strip()
                if 'folder' in lastPart and int(lastPart.split('folder')[1]) > max_folder_number:
                    max_folder_number = int(lastPart.split('folder')[1])
                elif 'file' in lastPart and int(lastPart.split('file')[1].split('.')[0]) > max_file_number:
                    max_file_number = int(lastPart.split('file')[1].split('.')[0])

file_serial_num += max_file_number
folder_serial_num += max_folder_number

# 建立txt檔，供後續復原檔/目錄名使用
recover_path = os.path.join(path, "recover.txt")
if not os.path.exists(recover_path):
    open(recover_path, 'w', encoding='utf-8').close()

with open(recover_path, 'a', encoding='utf-8') as f:
    # 遍歷所有目錄，若名稱過長，則將其標定為更改目標並預存新名稱
    for root, dirs, files in os.walk(path):                
        for dir in dirs:
            parent_folder_path = os.path.dirname(dir)
            folder_name = os.path.basename(dir)     
            new_folder_name = substitute_folder_name(folder_name)

            if new_folder_name != folder_name:
                folder_to_be_renamed.append((os.path.join(root, folder_name), os.path.join(root, new_folder_name)))
                print('此目錄名稱過長，暫時對其重新命名:')
                print(folder_name,"-->",new_folder_name)

    # 修改所有名稱過長之目錄
    for old_path, new_path in reversed(folder_to_be_renamed):
        f.write(f"{old_path}->{new_path}\n")
        os.rename(old_path, new_path)
    
    # 遍歷所有檔案，若名稱過長，則將其標定為更改目標並預存新名稱    
    for root, dirs, files in os.walk(path):
        for filename in files:
            if filename.endswith(".xls"):
                new_file_name = substitute_file_name(filename.replace(".xls","")) + '.xls'
                if filename != new_file_name :
                    file_to_be_renamed.append((os.path.join(root, filename), os.path.join(root, new_file_name)))
                    print('此檔案名稱過長，暫時對其重新命名:')
                    print(filename,"->",new_file_name)
    # 修改所有名稱過長之檔案
    for old_path, new_path in file_to_be_renamed:
        f.write(f"{old_path}->{new_path}\n")
        os.rename(old_path, new_path)

#以多執行緒操作xls轉xlsx
failed_files = []
Queue_files(path,deal_Queue)
gFileNum = deal_Queue.qsize()
gAllFileNum = gFileNum
threads = []
for i in range(os.cpu_count()):
    t = threading.Thread(target=convert_Xls)
    t.start()
    threads.append(t)
# 等待所有執行緒結束
for t in threads:
    t.join()


if getattr(sys, 'frozen', False):
  os.system(os.path.join(os.path.dirname(sys.executable), "recover_name.exe"))
else:
  os.system("python " + os.path.join( os.path.dirname(os.path.abspath(__file__)), 'recover_name.py'))
  
print("執行完畢，請確認檔案/資料夾名是否正確。")



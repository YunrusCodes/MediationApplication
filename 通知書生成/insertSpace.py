import os
import openpyxl
import sys

if getattr(sys, 'frozen', False):
  folder_path = os.path.join(os.path.dirname(sys.executable),"通知書輸出於此")
else:
  folder_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),"通知書輸出於此")

print(folder_path)
for folder in os.listdir(folder_path):
    print(folder)
    deSpace_folder = folder.replace(" ","") #先刪去空格，再依需求插入空格
    # 1.調解筆錄001(聲請人林阿語)(對造人陳小華)(收：112刑001號)(開調解時間112年01月17日（二）上午9時10分)-(車禍傷害糾紛案)
    name_cut1 = deSpace_folder.split("開調解時間") # 1.調解筆錄001(聲請人林阿語)(對造人陳小華)(收：112刑001號)( __|__ 112年01月17日（二）上午9時10分)-(車禍傷害糾紛案)
    name_cut2 = name_cut1[1].split("年")  # 112 __|__ 01月17日（二）上午9時10分)-(車禍傷害糾紛案)
    new_folder = name_cut1[0] + "開調解時間" + " " + name_cut2[0] + " 年" + name_cut2[1]

    os.rename(os.path.join(folder_path, folder), os.path.join(folder_path, new_folder))

    print("--->",os.path.join(folder_path, new_folder))

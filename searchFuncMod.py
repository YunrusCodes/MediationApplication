import requests
from bs4 import BeautifulSoup

def search_module(functions):
    for function in functions:
        print(f"搜尋函式 {function} 的相關模組:")
        url = f"https://pypi.org/search/?q={function}&c=Programming+Language+%3A%3A+Python"
        try:
            response = requests.get(url)
            response.raise_for_status()  # 检查请求是否成功
        except ConnectionError as e:
            print("无法建立连接:", str(e))
            print()
            continue
        except Exception as e:
            print("发生错误:", str(e))
            print()
            continue

        try:
            soup = BeautifulSoup(response.text, 'html.parser')
            results = soup.find_all('a', class_='package-snippet')
            
            if results:
                count = 0  # 添加计数器
                for result in results:
                    module_name = result.find('span', class_='package-snippet__name').text
                    module_description = result.find('p', class_='package-snippet__description').text
                    print(f"模組名稱: {module_name}")
                    print(f"模組描述: {module_description}")
                    print()
                    count += 1
                    if count == 3:  # 仅打印前三个结果
                        break
            else:
                print("找不到相關模組")
                print()
        except Exception as e:
            print("发生错误:", str(e))
            print()
            continue

def read_modules_file(file_path):
    try:
        with open(file_path, "r") as file:
            modules = file.readlines()
        return [module.strip() for module in modules]
    except FileNotFoundError:
        print(f"找不到檔案 {file_path}")

import sys
import os

if getattr(sys, 'frozen', False):
    absPath = os.path.dirname(sys.executable)
else:
    absPath = os.path.dirname(os.path.abspath(__file__))

modules_file = os.path.join(absPath,"modules.txt")
functions = read_modules_file(modules_file)

if functions:
    search_module(functions)

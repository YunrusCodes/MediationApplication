import ast
import os
import sys

def find_imported_modules(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        tree = ast.parse(file.read())

    imported_modules = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for alias in node.names:
                module_name = alias.name.split('.')[0]
                if not is_internal_module(module_name):
                    imported_modules.add(module_name)
        elif isinstance(node, ast.ImportFrom):
            module_name = node.module.split('.')[0]
            if not is_internal_module(module_name):
                imported_modules.add(module_name)

    return imported_modules

def is_internal_module(module_name):
    try:
        __import__(module_name)
    except ImportError:
        return True
    else:
        return False

if getattr(sys, 'frozen', False):
    # 程式被打包成執行檔 (exe)
    abspath = os.path.dirname(os.path.abspath(sys.argv[0]))
else:
    # 程式未被打包成執行檔
    abspath = os.path.dirname(os.path.abspath(__file__))

modules_set = set()
for path, directories, files in os.walk(abspath):
    for file in files:
        if file.endswith('.py'):
            file_path = os.path.join(path, file)
            # 查找匯入的模組
            imported_modules = find_imported_modules(file_path)
            modules_set.update(imported_modules)

modules_list = sorted(list(modules_set))

modules_txt_path = os.path.join(abspath, 'modules.txt')
if os.path.exists(modules_txt_path):
    os.remove(modules_txt_path)

with open(modules_txt_path, 'w', encoding='utf-8') as modules_file:
    for module in modules_list:
        modules_file.write(f"{module}\n")
    print(f"Created modules.txt and wrote imported modules.")

print("Process completed.")

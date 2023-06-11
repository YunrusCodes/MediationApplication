import os
import sys



def search_files_with_keyword(keyword, folder):
    matching_files = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if keyword in file:
                matching_files.append(os.path.join(root, file))
    return matching_files


def main():
    if getattr(sys, 'frozen', False):
      current_dir = os.path.dirname(sys.executable)
    else:
      current_dir = os.path.dirname(os.path.abspath(__file__))

    # 尋找指定的 docx 檔案
    docx_files = [
        "08.調解應攜帶證件（欠缺證件者恕無法辦理）.docx",
        "07.債權讓與同意書.docx",
        "06.委任書.docx"
    ]

    merge_exe_path = os.path.abspath(os.path.join(current_dir, "..", "mergePDF.exe"))
    # # 尋找名稱包含 "調解通知書" 的 xlsx 檔案並執行 mergePDF.exe
    for headfile in search_files_with_keyword("調解通知書", current_dir):
      args = f' \"{headfile}\" '
      for subfile in docx_files:
          args += f' \"{current_dir}\{subfile}\" '
      command = f'{merge_exe_path}{args}' 
      print(command)
      os.system(command)




if __name__ == "__main__":
    main()

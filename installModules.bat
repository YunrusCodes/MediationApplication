@echo off

set "modules_file=modules.txt"

rem 檢查 modules.txt 檔案是否存在
if not exist "%modules_file%" (
echo "modules.txt" 檔案不存在！請確認檔案是否存在並包含正確的模組名稱。
exit /b
)

rem 將當前目錄設定為 modules.txt 所在的目錄
cd /d "%~dp0"

rem 逐行讀取 modules.txt 檔案
for /f "usebackq delims=" %%m in ("%modules_file%") do (
echo 安裝模組: %%m
pip install %%m
)

echo 安裝完成！
pause
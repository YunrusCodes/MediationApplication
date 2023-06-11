@echo off
setlocal enabledelayedexpansion
set script_folder="%~dp0"
echo goto "%script_folder%"
cd "%script_folder%"
for /r %%i in (*.py) do (
    echo At %script_folder%
    echo %%~dpi
    if exist "%%~ni.exe" (
        echo Deleting %%~ni.exe
        del "%%~ni.exe"
    )
    set "s=%%~dpi"
    echo !s!
    set "s=!s:~0,-1!"
    echo !s!

    @REM echo %!s:~0,-1!%

    echo Do pyinstaller --onefile --distpath "!s!" "%%~fi"
    pyinstaller --onefile --distpath "!s!" "%%~fi"
)
pause
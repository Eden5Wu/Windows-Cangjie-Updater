@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
title 倉頡三代碼表自動更新工具 (路徑修正版)

:: 1. 設定變數
set "URL_7Z=https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z"
set "TARGET_DIR=%SystemRoot%\System32\zh-hk"
set "TEMP_WORK=%TEMP%\CJ3_Update"
set "FILE_7Z=%TEMP_WORK%\CJ3.7z"
set "EXE_7ZA=%TEMP_WORK%\7za.exe"

:: 這裡指定 7z 內部的精確路徑 (注意 7z 內部路徑斜線方向)
set "INTERNAL_FILE=Windows 10 2004及之后的Windows\ChtCangjieExt.lex"

echo [*] 正在建立工作目錄...
if exist "%TEMP_WORK%" rd /s /q "%TEMP_WORK%"
mkdir "%TEMP_WORK%"

echo [*] 正在下載 7-Zip 獨立版工具...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri 'https://github.com/mcmilk/7-Zip-zstd/releases/download/v22.01-v1.5.2-R1/7za.exe' -OutFile '%EXE_7ZA%'"

echo [*] 正在下載三代碼表壓縮檔...
powershell -Command "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%URL_7Z%' -OutFile '%FILE_7Z%'"

if %ERRORLEVEL% NEQ 0 (
    echo [!] 下載失敗，請檢查網路。
    pause
    exit /b
)

echo [*] 正在從 7z 提取指定版本：!INTERNAL_FILE!
:: 直接提取特定路徑檔案，避開遍歷錯誤
"%EXE_7ZA%" e "%FILE_7Z%" -o"%TEMP_WORK%" "%INTERNAL_FILE%" -r -y >nul

set "SOURCE_LEX=%TEMP_WORK%\ChtCangjieExt.lex"

if not exist "%SOURCE_LEX%" (
    echo [!] 提取失敗！找不到指定的路徑檔案。
    echo 預期路徑: %INTERNAL_FILE%
    pause
    exit /b
)

echo [*] 正在關閉輸入法背景進程...
taskkill /F /IM CTFMON.EXE /T >nul 2>&1
taskkill /F /IM MicrosoftIME.exe /T >nul 2>&1

echo [*] 執行權限檢查與備份...
set "DATE_STR=%date:~0,4%%date:~5,2%%date:~8,2%"
set "BACKUP_DIR=%TARGET_DIR%\Backup_CJ3_%DATE_STR%"
mkdir "%BACKUP_DIR%" 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [!] 權限不足！請右鍵點選「以系統管理員身分執行」。
    pause
    exit /b
)

copy /Y "%TARGET_DIR%\ChtCangjie.sdc" "%BACKUP_DIR%\" >nul
copy /Y "%TARGET_DIR%\ChtCangjie.spd" "%BACKUP_DIR%\" >nul
copy /Y "%TARGET_DIR%\ChtCangjieExt.lex" "%BACKUP_DIR%\" >nul

echo [*] 正在替換系統碼表 (三代 2004+ 版)...
del /F /Q "%TARGET_DIR%\ChtCangjie.sdc"
del /F /Q "%TARGET_DIR%\ChtCangjie.spd"
del /F /Q "%TARGET_DIR%\ChtCangjieExt.lex"
copy /Y "%SOURCE_LEX%" "%TARGET_DIR%\ChtCangjieExt.lex"

echo [*] 寫入 HKSCS 登錄檔設定...
reg add "HKEY_CURRENT_USER\Software\Microsoft\IME\15.0\CHT\Cangjie" /v "Enable HKSCS" /t REG_DWORD /d 1 /f >nul

echo.
echo ==================================================
echo 三代更新完畢 (已選用 Windows 10 2004 及之後版本)
echo 1. 已備份舊檔至 System32\zh-hk\Backup_CJ3_%DATE_STR%
echo 2. 請「立即重新啟動電腦」以讓碼表生效
echo ==================================================
rd /s /q "%TEMP_WORK%"
pause
start ctfmon.exe
exit
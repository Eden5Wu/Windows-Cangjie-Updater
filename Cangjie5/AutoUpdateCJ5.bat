@echo off
:: 切換至 UTF-8 碼頁，確保能處理中文字路徑
chcp 65001 >nul
setlocal enabledelayedexpansion
title 倉頡五代碼表自動更新工具 (壓縮檔提取版)

:: 1. 設定變數
set "ZIP_URL=https://github.com/Jackchows/Cangjie5/releases/download/v4.1-beta/MSCJData_Cangjie5_20260207.zip"
set "TARGET_DIR=%SystemRoot%\System32\zh-hk"
set "TEMP_WORK=%TEMP%\CJ5_Update"
set "ZIP_FILE=%TEMP_WORK%\CJ5.zip"

:: 這裡指定壓縮檔內的子路徑（對應您需要的「傳統漢字優先」版本）
set "INTERNAL_PATH=傳統漢字優先（偏好台灣用字習慣）/ChtCangjieExt.lex"

echo [*] 正在建立工作目錄...
if exist "%TEMP_WORK%" rd /s /q "%TEMP_WORK%"
mkdir "%TEMP_WORK%"

echo [*] 正在下載壓縮檔...
powershell -Command "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile '%ZIP_FILE%'"

if %ERRORLEVEL% NEQ 0 (
    echo [!] 下載失敗，請檢查網路連線。
    pause
    exit /b
)

echo [*] 正在從 ZIP 提取指定碼表 (ChtCangjieExt.lex)...
powershell -Command "Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%TEMP_WORK%\Extracted' -Force"

echo [*] 正在定位碼表檔案...
:: 直接組合路徑，並將 INTERNAL_PATH 的 / 換成 \
set "TEMP_INTERNAL=%INTERNAL_PATH:/=\%"
set "SOURCE_LEX=%TEMP_WORK%\Extracted\%TEMP_INTERNAL%"

:: 萬一直接指定失敗（通常是編碼導致路徑對不起來），就用 dir 指令掃描含有 "台灣" 的路徑
if not exist "%SOURCE_LEX%" (
    for /f "delims=" %%a in ('dir /s /b "%TEMP_WORK%\Extracted\ChtCangjieExt.lex" ^| findstr "台灣"') do (
        set "SOURCE_LEX=%%a"
    )
)

:: 最後檢查一次變數是否有值
if "%SOURCE_LEX%"=="" (
    echo [!] 找不到檔案，路徑可能有誤。
    pause
    exit /b
)

echo [*] 已定位檔案: "%SOURCE_LEX%"

echo [*] 正在停止 Microsoft IME 進程...
taskkill /F /IM CTFMON.EXE /T >nul 2>&1
taskkill /F /IM MicrosoftIME.exe /T >nul 2>&1

echo [*] 正在備份與清理舊檔案...
set "BACKUP_DIR=%TARGET_DIR%\Backup_%date:~0,4%%date:~5,2%%date:~8,2%"
if not exist "%BACKUP_DIR%" mkdir "%BACKUP_DIR%"
copy /Y "%TARGET_DIR%\ChtCangjie.sdc" "%BACKUP_DIR%\"
copy /Y "%TARGET_DIR%\ChtCangjie.spd" "%BACKUP_DIR%\"
copy /Y "%TARGET_DIR%\ChtCangjieExt.lex" "%BACKUP_DIR%\"

del /F /Q "%TARGET_DIR%\ChtCangjie.sdc"
del /F /Q "%TARGET_DIR%\ChtCangjie.spd"
del /F /Q "%TARGET_DIR%\ChtCangjieExt.lex"

echo [*] 正在部署新碼表...
copy /Y "%SOURCE_LEX%" "%TARGET_DIR%\ChtCangjieExt.lex"

echo [*] 正在自動開啟 HKSCS 選項 (修改登錄檔)...
reg add "HKEY_CURRENT_USER\Software\Microsoft\IME\15.0\CHT\Cangjie" /v "Enable HKSCS" /t REG_DWORD /d 1 /f >nul 2>&1

echo.
echo ==========================================
echo 更新成功！
echo 已提取：傳統漢字優先（偏好台灣用字習慣）
echo 已自動備份至：%BACKUP_DIR%
echo.
echo 請「重新啟動電腦」以生效。
echo ==========================================
rd /s /q "%TEMP_WORK%"
pause
start ctfmon.exe
exit
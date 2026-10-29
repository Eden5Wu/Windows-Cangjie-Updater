@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion
title 倉頡三代碼表自動更新工具

:: ==========================================
:: 1. 設定
:: ==========================================

set "VERSION=Windows 10 2004及之后的Windows"

set "DOWNLOAD_URL=https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z"

:: GitHub Release 4.2 官方附件資訊
set "EXPECTED_SIZE=613787"
set "EXPECTED_SHA256=342422de88f78a1a924042dad3b00689e23f1ce612e3a718243ebf7e070e7954"

set "TARGET_DIR=%SystemRoot%\System32\zh-hk"
set "BAT_FILE=%~f0"

set "USER_SID=%~1"
set "DEPLOY_STARTED=0"
set "REG_TOUCHED=0"
set "BACKUP_DIR="
set "WORK_DIR="
set "STAGE=初始化"

:: ==========================================
:: 2. 取得原始使用者 SID
:: ==========================================

if not defined USER_SID (
    for /f "delims=" %%S in ('
        powershell -NoProfile -Command "[Security.Principal.WindowsIdentity]::GetCurrent().User.Value"
    ') do set "USER_SID=%%S"
)

if not defined USER_SID (
    echo [ERROR] 無法取得使用者 SID。
    pause
    exit /b 1
)

:: ==========================================
:: 3. 自動取得管理員權限
:: ==========================================

fltmc >nul 2>&1

if errorlevel 1 (
    echo [INFO] 正在請求系統管理員權限...

    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process -FilePath $env:BAT_FILE -ArgumentList $env:USER_SID -Verb RunAs -ErrorAction Stop"

    if errorlevel 1 (
        echo [ERROR] 無法取得管理員權限。
        pause
        exit /b 1
    )

    exit /b 0
)

:: ==========================================
:: 4. 檢查執行環境
:: ==========================================

echo.
echo ==========================================
echo  倉頡三代碼表自動更新工具
echo ==========================================
echo.
echo 版本：%VERSION%
echo.

set "STAGE=檢查系統環境"

if not exist "%TARGET_DIR%\" (
    echo [ERROR] 找不到系統輸入法目錄：
    echo %TARGET_DIR%
    goto :fail
)

:: 確保修改的是原始使用者的 HKSCS 設定，
:: 而非 UAC 提升權限後的其他管理員帳戶。

reg query "HKU\%USER_SID%" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 找不到原始使用者的登錄檔。
    echo SID：%USER_SID%
    goto :fail
)

set "REG_KEY=HKU\%USER_SID%\Software\Microsoft\IME\15.0\CHT\Cangjie"

:: ==========================================
:: 5. 尋找已安裝的 7-Zip
:: ==========================================

echo [1/7] 檢查 7-Zip...

set "SEVENZIP="

if exist "%ProgramFiles%\7-Zip\7z.exe" (
    set "SEVENZIP=%ProgramFiles%\7-Zip\7z.exe"
)

if not defined SEVENZIP (
    if exist "%ProgramFiles(x86)%\7-Zip\7z.exe" (
        set "SEVENZIP=%ProgramFiles(x86)%\7-Zip\7z.exe"
    )
)

if not defined SEVENZIP (
    for /f "delims=" %%F in ('where 7z.exe 2^>nul') do (
        if not defined SEVENZIP set "SEVENZIP=%%F"
    )
)

if not defined SEVENZIP (
    echo.
    echo [ERROR] 找不到 7-Zip。
    echo.
    echo 請先安裝官方 7-Zip：
    echo https://www.7-zip.org/
    goto :fail
)

echo [OK] %SEVENZIP%

:: ==========================================
:: 6. 建立獨立暫存目錄
:: ==========================================

set "STAGE=建立工作目錄"

for /f "delims=" %%G in ('
    powershell -NoProfile -Command "[guid]::NewGuid().ToString('N')"
') do set "WORK_ID=%%G"

if not defined WORK_ID (
    echo [ERROR] 無法建立工作目錄識別碼。
    goto :fail
)

set "WORK_DIR=%TEMP%\CJ3_Update_%WORK_ID%"
set "EXTRACT_DIR=%WORK_DIR%\Extracted"
set "ARCHIVE_FILE=%WORK_DIR%\CJ3.7z"

mkdir "%WORK_DIR%" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 無法建立暫存目錄。
    goto :fail
)

mkdir "%EXTRACT_DIR%" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 無法建立解壓縮目錄。
    goto :fail
)

:: ==========================================
:: 7. 下載碼表
:: ==========================================

echo.
echo [2/7] 下載倉頡三代碼表...

set "STAGE=下載碼表"

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri $env:DOWNLOAD_URL -OutFile $env:ARCHIVE_FILE -UseBasicParsing -ErrorAction Stop"

if errorlevel 1 (
    echo [ERROR] 下載失敗。
    goto :fail
)

if not exist "%ARCHIVE_FILE%" (
    echo [ERROR] 找不到下載的壓縮檔。
    goto :fail
)

:: ==========================================
:: 8. 驗證下載檔案
:: ==========================================

echo.
echo [3/7] 驗證下載檔案...

set "STAGE=驗證下載檔案"

for %%F in ("%ARCHIVE_FILE%") do set "ACTUAL_SIZE=%%~zF"

if not "%ACTUAL_SIZE%"=="%EXPECTED_SIZE%" (
    echo [ERROR] 檔案大小不符。
    echo 預期：%EXPECTED_SIZE% Bytes
    echo 實際：%ACTUAL_SIZE% Bytes
    goto :fail
)

set "ACTUAL_SHA256="

for /f "delims=" %%H in ('
    powershell -NoProfile -Command "(Get-FileHash -LiteralPath $env:ARCHIVE_FILE -Algorithm SHA256).Hash"
') do set "ACTUAL_SHA256=%%H"

if /I not "%ACTUAL_SHA256%"=="%EXPECTED_SHA256%" (
    echo [ERROR] SHA-256 驗證失敗。
    echo.
    echo 預期：%EXPECTED_SHA256%
    echo 實際：%ACTUAL_SHA256%
    goto :fail
)

echo [OK] 檔案大小與 SHA-256 驗證成功。

:: ==========================================
:: 9. 解壓縮與尋找碼表
:: ==========================================

echo.
echo [4/7] 解壓縮倉頡三代碼表...

set "STAGE=解壓縮碼表"

"%SEVENZIP%" x "%ARCHIVE_FILE%" "-o%EXTRACT_DIR%" -y >nul

if errorlevel 1 (
    echo [ERROR] 7-Zip 解壓縮失敗。
    goto :fail
)

:: 使用 PowerShell 原生檔案物件查找，
:: 避免 CMD 處理中文目錄名稱時產生比對問題。

set "SOURCE_LEX="

for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command ^
    "[Console]::OutputEncoding=[Text.Encoding]::UTF8; Get-ChildItem -LiteralPath $env:EXTRACT_DIR -Recurse -File -Filter 'ChtCangjieExt.lex' ^| Where-Object { $_.DirectoryName -like '*Windows 10 2004及之后的Windows*' } ^| Select-Object -First 1 -ExpandProperty FullName"`) do (
    set "SOURCE_LEX=%%F"
)

if not defined SOURCE_LEX (
    echo [ERROR] 找不到指定版本的碼表。
    echo 版本：%VERSION%
    goto :fail
)

if not exist "%SOURCE_LEX%" (
    echo [ERROR] 碼表檔案不存在。
    goto :fail
)

for %%F in ("%SOURCE_LEX%") do set "LEX_SIZE=%%~zF"

if "%LEX_SIZE%"=="0" (
    echo [ERROR] 碼表檔案大小為 0。
    goto :fail
)

echo [OK] 找到碼表。
echo %SOURCE_LEX%

:: ==========================================
:: 10. 建立唯一備份目錄
:: ==========================================

echo.
echo [5/7] 備份原始系統碼表...

set "STAGE=備份系統碼表"

for /f "delims=" %%T in ('
    powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"
') do set "DATE_STR=%%T"

if not defined DATE_STR (
    echo [ERROR] 無法取得備份時間。
    goto :fail
)

set "BACKUP_DIR=%TARGET_DIR%\Backup_CJ3_%DATE_STR%_%RANDOM%"

mkdir "%BACKUP_DIR%" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 無法建立備份目錄。
    goto :fail
)

:: 記錄更新前各檔案是否存在，
:: 以便部署失敗時還原原始狀態。

set "HAD_SDC=0"
set "HAD_SPD=0"
set "HAD_LEX=0"

if exist "%TARGET_DIR%\ChtCangjie.sdc" set "HAD_SDC=1"
if exist "%TARGET_DIR%\ChtCangjie.spd" set "HAD_SPD=1"
if exist "%TARGET_DIR%\ChtCangjieExt.lex" set "HAD_LEX=1"

:: 備份全部完成並驗證後，才允許修改系統檔案。

for %%F in (
    ChtCangjie.sdc
    ChtCangjie.spd
    ChtCangjieExt.lex
) do (
    if exist "%TARGET_DIR%\%%F" (
        copy /B /Y "%TARGET_DIR%\%%F" "%BACKUP_DIR%\%%F" >nul

        if errorlevel 1 (
            echo [ERROR] 備份失敗：%%F
            goto :fail
        )

        fc /B "%TARGET_DIR%\%%F" "%BACKUP_DIR%\%%F" >nul 2>&1

        if errorlevel 1 (
            echo [ERROR] 備份內容驗證失敗：%%F
            goto :fail
        )

        echo [OK] 已備份：%%F
    )
)

echo.
echo [OK] 備份完成。
echo %BACKUP_DIR%

:: ==========================================
:: 11. 記錄 HKSCS 原始設定
:: ==========================================

set "OLD_HKSCS="

for /f "tokens=4" %%V in ('
    reg query "%REG_KEY%" /v "Enable HKSCS" 2^>nul ^| findstr /I /C:"Enable HKSCS"
') do set "OLD_HKSCS=%%V"

:: ==========================================
:: 12. 停止 IME 程序
:: ==========================================

echo.
echo [6/7] 部署新碼表...

set "STAGE=部署系統碼表"

taskkill /F /IM ctfmon.exe /T >nul 2>&1
taskkill /F /IM MicrosoftIME.exe /T >nul 2>&1

timeout /t 1 /nobreak >nul

:: ==========================================
:: 13. 替換系統碼表
:: ==========================================

set "DEPLOY_STARTED=1"

for %%F in (
    ChtCangjie.sdc
    ChtCangjie.spd
    ChtCangjieExt.lex
) do (
    if exist "%TARGET_DIR%\%%F" (
        del /F /Q "%TARGET_DIR%\%%F" >nul 2>&1

        if exist "%TARGET_DIR%\%%F" (
            echo [ERROR] 無法刪除：%%F
            goto :fail
        )
    )
)

copy /B /Y "%SOURCE_LEX%" "%TARGET_DIR%\ChtCangjieExt.lex" >nul

if errorlevel 1 (
    echo [ERROR] 新碼表複製失敗。
    goto :fail
)

if not exist "%TARGET_DIR%\ChtCangjieExt.lex" (
    echo [ERROR] 目標碼表不存在。
    goto :fail
)

:: 驗證來源與部署檔案完全一致。

fc /B "%SOURCE_LEX%" "%TARGET_DIR%\ChtCangjieExt.lex" >nul 2>&1

if errorlevel 1 (
    echo [ERROR] 部署後檔案內容驗證失敗。
    goto :fail
)

echo [OK] 碼表部署成功，內容比對一致。

:: ==========================================
:: 14. 啟用 HKSCS
:: ==========================================

echo.
echo [7/7] 啟用 HKSCS 支援...

set "STAGE=設定 HKSCS"
set "REG_TOUCHED=1"

reg add "%REG_KEY%" /v "Enable HKSCS" /t REG_DWORD /d 1 /f >nul 2>&1

if errorlevel 1 (
    echo [ERROR] HKSCS 設定失敗。
    goto :fail
)

reg query "%REG_KEY%" /v "Enable HKSCS" 2>nul | findstr /I /C:"0x1" >nul

if errorlevel 1 (
    echo [ERROR] HKSCS 設定驗證失敗。
    goto :fail
)

echo [OK] HKSCS 已啟用。

set "DEPLOY_STARTED=0"
set "REG_TOUCHED=0"

goto :success

:: ==========================================
:: 15. 更新失敗與還原
:: ==========================================

:fail

echo.
echo ==========================================
echo  [ERROR] 更新失敗！
echo ==========================================
echo 失敗階段：%STAGE%
echo.

if "%DEPLOY_STARTED%"=="1" (
    echo [INFO] 正在還原原始系統碼表...

    call :restore_file "ChtCangjie.sdc" "%HAD_SDC%"
    call :restore_file "ChtCangjie.spd" "%HAD_SPD%"
    call :restore_file "ChtCangjieExt.lex" "%HAD_LEX%"
)

if "%REG_TOUCHED%"=="1" (
    echo [INFO] 正在還原 HKSCS 設定...

    if defined OLD_HKSCS (
        reg add "%REG_KEY%" /v "Enable HKSCS" /t REG_DWORD /d "%OLD_HKSCS%" /f >nul 2>&1
    ) else (
        reg delete "%REG_KEY%" /v "Enable HKSCS" /f >nul 2>&1
    )
)

if defined BACKUP_DIR (
    echo.
    echo 原始備份保留於：
    echo %BACKUP_DIR%
)

if defined WORK_DIR (
    if exist "%WORK_DIR%\" (
        rd /S /Q "%WORK_DIR%" >nul 2>&1
    )
)

echo.
echo 若系統檔案未能完全還原，請先檢查備份。
echo.

pause
exit /b 1

:: ==========================================
:: 16. 還原單一檔案
:: ==========================================

:restore_file

set "RESTORE_NAME=%~1"
set "RESTORE_EXISTED=%~2"

if "%RESTORE_EXISTED%"=="1" (
    if not exist "%BACKUP_DIR%\%RESTORE_NAME%" (
        echo [ERROR] 找不到備份：%RESTORE_NAME%
        exit /b 1
    )

    copy /B /Y "%BACKUP_DIR%\%RESTORE_NAME%" "%TARGET_DIR%\%RESTORE_NAME%" >nul 2>&1

    if errorlevel 1 (
        echo [ERROR] 還原失敗：%RESTORE_NAME%
        exit /b 1
    )

    fc /B "%BACKUP_DIR%\%RESTORE_NAME%" "%TARGET_DIR%\%RESTORE_NAME%" >nul 2>&1

    if errorlevel 1 (
        echo [ERROR] 還原檔案驗證失敗：%RESTORE_NAME%
        exit /b 1
    )
) else (
    if exist "%TARGET_DIR%\%RESTORE_NAME%" (
        del /F /Q "%TARGET_DIR%\%RESTORE_NAME%" >nul 2>&1

        if exist "%TARGET_DIR%\%RESTORE_NAME%" (
            echo [ERROR] 無法清除：%RESTORE_NAME%
            exit /b 1
        )
    )
)

echo [OK] 已還原：%RESTORE_NAME%
exit /b 0

:: ==========================================
:: 17. 更新成功
:: ==========================================

:success

echo.
echo ==========================================
echo  倉頡三代碼表更新成功！
echo ==========================================
echo.
echo 已安裝版本：
echo %VERSION%
echo.
echo 碼表位置：
echo %TARGET_DIR%\ChtCangjieExt.lex
echo.
echo 原始備份：
echo %BACKUP_DIR%
echo.
echo 已啟用 HKSCS 支援。
echo.
echo 請重新啟動 Windows，使新的碼表完全生效。
echo ==========================================
echo.

:: ==========================================
:: 18. 清理暫存檔案
:: ==========================================

if defined WORK_DIR (
    if exist "%WORK_DIR%\" (
        rd /S /Q "%WORK_DIR%" >nul 2>&1

        if exist "%WORK_DIR%\" (
            echo [WARN] 暫存目錄未能完全清理。
            echo %WORK_DIR%
        )
    )
)

pause
exit /b 0

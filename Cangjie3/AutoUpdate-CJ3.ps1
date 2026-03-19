# 請直接對著 .ps1 檔案按右鍵 ->「用 PowerShell 執行」來安裝
# ==========================================
# 1. 自動升權區塊
# ==========================================
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "權限不足，正在請求以管理員身分啟動..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$Host.UI.RawUI.WindowTitle = "倉頡三代碼表自動更新工具 (PowerShell 版)"

# 確保支援 TLS 1.2，避免舊版系統下載 GitHub 檔案失敗
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ==========================================
# 2. 設定通用變數
# ==========================================
$url7z = "https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z"
$url7za = "https://github.com/mcmilk/7-Zip-zstd/releases/download/v22.01-v1.5.2-R1/7za.exe"

$targetDir = "$env:SystemRoot\System32\zh-hk"
$tempWork = "$env:TEMP\CJ3_Update"
$file7z = "$tempWork\CJ3.7z"
$exe7za = "$tempWork\7za.exe"

# ==========================================
# 3. 互動式選單 (含 10 秒倒數與預設值)
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " 請選擇你要安裝的倉頡三代碼表版本：" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  1. Windows 10 2004 及之后的Windows [預設]"
Write-Host "  2. Windows 10 2004 之前的版本"
Write-Host "  0. 退出程式"
Write-Host "==========================================`n" -ForegroundColor Cyan

$timeout = 10          # 倒數秒數
$defaultChoice = '1'   # 預設選項 (1號：2004及之後版本)
$choice = $null
$validChoice = $false

$timer = [System.Diagnostics.Stopwatch]::StartNew()

while (-not $validChoice) {
    if ($timer.Elapsed.TotalSeconds -ge $timeout) {
        $choice = $defaultChoice
        $validChoice = $true
        Write-Host "`n`n[!] 倒數結束，自動選擇預設值 (1)。" -ForegroundColor Yellow
        break
    }
    
    $timeLeft = $timeout - [math]::Floor($timer.Elapsed.TotalSeconds)
    Write-Host "`r請輸入選項代碼 (0-2) [倒數 $timeLeft 秒後自動選擇 2004+ 版]: " -NoNewline -ForegroundColor Yellow
    
    if ([console]::KeyAvailable) {
        $key = [console]::ReadKey($true)
        $choice = $key.KeyChar
        
        if ($choice -match '^[0-2]$') {
            $validChoice = $true
            Write-Host "`n"
        } else {
            Write-Host "`n`n[!] 無效的輸入 '$choice'，請輸入 0 到 2 之間的數字。" -ForegroundColor Red
            $timer.Restart()
        }
    }
    Start-Sleep -Milliseconds 100
}
$timer.Stop()

switch ($choice) {
    '1' { $prefVersion = "Windows 10 2004及之后的Windows" }
    '2' { $prefVersion = "Windows 10 2004之前的版本" }
    '0' { 
        Write-Host "已取消執行，即將關閉視窗。" -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        exit 
    }
}

Write-Host "[*] 你選擇了：「$prefVersion」" -ForegroundColor Green
Write-Host "[*] 準備開始執行更新程序...`n" -ForegroundColor Cyan
Start-Sleep -Seconds 1

# ==========================================
# 4. 建立工作目錄與下載
# ==========================================
Write-Host "[*] 正在建立工作目錄..." -ForegroundColor Cyan
if (Test-Path $tempWork) { Remove-Item -Path $tempWork -Recurse -Force }
New-Item -ItemType Directory -Path $tempWork | Out-Null

try {
    Write-Host "[*] 正在下載 7-Zip 解壓縮工具 (7za.exe)..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $url7za -OutFile $exe7za -ErrorAction Stop
    
    Write-Host "[*] 正在下載倉頡三代碼表壓縮檔 (.7z)..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $url7z -OutFile $file7z -ErrorAction Stop
} catch {
    Write-Host "[!] 下載失敗，請檢查網路連線。錯誤細節: $($_.Exception.Message)" -ForegroundColor Red
    Pause
    exit
}

# ==========================================
# 5. 提取指定的碼表檔案
# ==========================================
Write-Host "[*] 正在從 7z 提取指定版本的碼表..." -ForegroundColor Cyan

# 組合 7z 內部路徑 (例如: "Windows 10 2004及之后的Windows\ChtCangjieExt.lex")
$internalFile = "$prefVersion\ChtCangjieExt.lex"
$sourceLex = "$tempWork\ChtCangjieExt.lex"

# 呼叫 7za.exe 進行解壓縮 (e: 提取至平坦目錄, -r: 遞迴, -y: 全部皆是)
$7zArgs = "e `"$file7z`" -o`"$tempWork`" `"$internalFile`" -r -y"
Start-Process -FilePath $exe7za -ArgumentList $7zArgs -Wait -NoNewWindow

if (!(Test-Path $sourceLex)) {
    Write-Host "[!] 提取失敗！找不到指定的路徑檔案。" -ForegroundColor Red
    Write-Host "預期路徑: $internalFile" -ForegroundColor Red
    Pause
    exit
}

# ==========================================
# 6. 停止進程、備份與清理舊檔案
# ==========================================
Write-Host "[*] 正在停止 Microsoft IME 進程..." -ForegroundColor Cyan
Stop-Process -Name "ctfmon", "MicrosoftIME" -Force -ErrorAction SilentlyContinue

Write-Host "[*] 正在備份與清理舊檔案..." -ForegroundColor Cyan
$dateStr = Get-Date -Format "yyyyMMdd"
$backupDir = "$targetDir\Backup_CJ3_$dateStr"

if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

$filesToProcess = @("ChtCangjie.sdc", "ChtCangjie.spd", "ChtCangjieExt.lex")
foreach ($file in $filesToProcess) {
    $fullPath = "$targetDir\$file"
    if (Test-Path $fullPath) {
        Copy-Item -Path $fullPath -Destination $backupDir -Force
        Remove-Item -Path $fullPath -Force
    }
}

# ==========================================
# 7. 部署新碼表與修改 Registry
# ==========================================
Write-Host "[*] 正在部署新碼表..." -ForegroundColor Cyan
Copy-Item -Path $sourceLex -Destination "$targetDir\ChtCangjieExt.lex" -Force

Write-Host "[*] 正在自動開啟 HKSCS 選項 (修改登錄檔)..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\IME\15.0\CHT\Cangjie"
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "Enable HKSCS" -Value 1 -Type DWord -Force

# ==========================================
# 8. 清理與完成
# ==========================================
Write-Host "`n==================================================" -ForegroundColor Green
Write-Host " 三代更新完畢！" -ForegroundColor Green
Write-Host " 已選用版本：$prefVersion" -ForegroundColor Green
Write-Host " 舊檔已備份至：$backupDir" -ForegroundColor Green
Write-Host ""
Write-Host " 請「立即重新啟動電腦」以讓碼表完全生效。" -ForegroundColor Yellow
Write-Host "==================================================`n"

Write-Host "[*] 清理暫存檔案..." -ForegroundColor Cyan
Remove-Item -Path $tempWork -Recurse -Force

Pause
Start-Process ctfmon.exe
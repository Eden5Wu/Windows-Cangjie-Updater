# 請直接對著 .ps1 檔案按右鍵 ->「用 PowerShell 執行」來安裝
# ==========================================
# 1. 自動升權區塊
# ==========================================
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "權限不足，正在請求以管理員身分啟動..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$Host.UI.RawUI.WindowTitle = "倉頡碼表自動更新工具 (PowerShell 版)"

# ==========================================
# 2. 設定通用變數
# ==========================================
$zipUrl = "https://github.com/Jackchows/Cangjie5/releases/download/v4.1-beta/MSCJData_Cangjie5_20260207.zip"
$targetDir = "$env:SystemRoot\System32\zh-hk"
$tempWork = "$env:TEMP\CJ5_Update"
$zipFile = "$tempWork\CJ5.zip"

# ==========================================
# 3. 互動式選單 (含 10 秒倒數與預設值)
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Cyan
Write-Host " 請選擇你要安裝的倉頡碼表版本：" -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "  1. 一般排序"
Write-Host "  2. 傳統漢字優先（偏好台灣用字習慣） [預設]"
Write-Host "  3. 傳統漢字優先（偏好香港用字習慣）"
Write-Host "  4. 簡化字優先"
Write-Host "  0. 退出程式"
Write-Host "==========================================`n" -ForegroundColor Cyan

$timeout = 10          # 倒數秒數
$defaultChoice = '2'   # 預設選項 (台灣版)
$choice = $null
$validChoice = $false

$timer = [System.Diagnostics.Stopwatch]::StartNew()

# 倒數迴圈
while (-not $validChoice) {
    # 檢查是否超時
    if ($timer.Elapsed.TotalSeconds -ge $timeout) {
        $choice = $defaultChoice
        $validChoice = $true
        Write-Host "`n`n[!] 倒數結束，自動選擇預設值 (2)。" -ForegroundColor Yellow
        break
    }
    
    # 計算剩餘時間並使用 `r (回車) 在同一行動態更新畫面
    $timeLeft = $timeout - [math]::Floor($timer.Elapsed.TotalSeconds)
    Write-Host "`r請輸入選項代碼 (0-4) [倒數 $timeLeft 秒後自動選擇台灣版]: " -NoNewline -ForegroundColor Yellow
    
    # 檢查使用者是否按下了按鍵
    if ([console]::KeyAvailable) {
        $key = [console]::ReadKey($true) # $true 代表不把按下的鍵顯示在畫面上
        $choice = $key.KeyChar
        
        # 判斷輸入是否在 0 到 4 之間
        if ($choice -match '^[0-4]$') {
            $validChoice = $true
            Write-Host "`n" # 換行
        } else {
            Write-Host "`n`n[!] 無效的輸入 '$choice'，請輸入 0 到 4 之間的數字。" -ForegroundColor Red
            $timer.Restart() # 輸入錯誤，重新開始 10 秒倒數
        }
    }
    Start-Sleep -Milliseconds 100 # 短暫暫停，降低 CPU 使用率
}

$timer.Stop()

# 根據選項設定版本變數
switch ($choice) {
    '1' { $prefVersion = "一般排序" }
    '2' { $prefVersion = "傳統漢字優先（偏好台灣用字習慣）" }
    '3' { $prefVersion = "傳統漢字優先（偏好香港用字習慣）" }
    '4' { $prefVersion = "簡化字優先" }
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

Write-Host "[*] 正在下載壓縮檔..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -ErrorAction Stop
} catch {
    Write-Host "[!] 下載失敗，請檢查網路連線。錯誤細節: $($_.Exception.Message)" -ForegroundColor Red
    Pause
    exit
}

Write-Host "[*] 正在從 ZIP 提取檔案..." -ForegroundColor Cyan
Expand-Archive -Path $zipFile -DestinationPath "$tempWork\Extracted" -Force

# ==========================================
# 5. 尋找解壓後的指定碼表
# ==========================================
Write-Host "[*] 正在尋找指定的碼表..." -ForegroundColor Cyan

$sourceLex = Get-ChildItem -Path "$tempWork\Extracted" -Recurse -Filter "ChtCangjieExt.lex" |
             Where-Object { $_.DirectoryName -like "*$prefVersion*" } |
             Select-Object -First 1

if (!$sourceLex) {
    Write-Host "[!] 找不到指定的碼表檔案，請確認壓縮檔內容。" -ForegroundColor Red
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
$backupDir = "$targetDir\Backup_$dateStr"

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
Copy-Item -Path $sourceLex.FullName -Destination "$targetDir\ChtCangjieExt.lex" -Force

Write-Host "[*] 正在自動開啟 HKSCS 選項 (修改登錄檔)..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\IME\15.0\CHT\Cangjie"
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "Enable HKSCS" -Value 1 -Type DWord -Force

# ==========================================
# 8. 清理與完成
# ==========================================
Write-Host "`n==========================================" -ForegroundColor Green
Write-Host "更新成功！" -ForegroundColor Green
Write-Host "已提取版本：$prefVersion" -ForegroundColor Green
Write-Host "已自動備份至：$backupDir" -ForegroundColor Green
Write-Host ""
Write-Host "請「重新啟動電腦」以完全生效。" -ForegroundColor Yellow
Write-Host "==========================================`n"

Write-Host "[*] 清理暫存檔案..." -ForegroundColor Cyan
Remove-Item -Path $tempWork -Recurse -Force

Pause
Start-Process ctfmon.exe
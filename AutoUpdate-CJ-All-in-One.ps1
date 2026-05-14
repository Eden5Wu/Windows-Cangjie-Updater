# ==========================================
# 1. 自動升權區塊
# ==========================================
if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "權限不足，正在請求以管理員身分啟動..." -ForegroundColor Yellow
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

$Host.UI.RawUI.WindowTitle = "倉頡碼表自動更新工具 (三代/五代整合版)"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ==========================================
# 2. 設定通用變數
# ==========================================
$targetDir = "$env:SystemRoot\System32\zh-hk"
$tempWork  = "$env:TEMP\CJ_Update_All"

# ==========================================
# 3. 共用選單倒數函式
# ==========================================
function Show-CountdownMenu {
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$PromptText,
        [string]$DefaultChoice,
        [string]$ValidRegex,
        [int]$Timeout = 10
    )
    Write-Host "`n==========================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan
    foreach ($opt in $Options) { Write-Host "  $opt" }
    Write-Host "==========================================`n" -ForegroundColor Cyan

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $choice = $null
    $validChoice = $false

    while (-not $validChoice) {
        if ($timer.Elapsed.TotalSeconds -ge $Timeout) {
            $choice = $DefaultChoice
            $validChoice = $true
            Write-Host "`n`n[!] 倒數結束，自動選擇預設值 ($DefaultChoice)。" -ForegroundColor Yellow
            break
        }
        
        $timeLeft = $Timeout - [math]::Floor($timer.Elapsed.TotalSeconds)
        Write-Host "`r$PromptText [倒數 $timeLeft 秒後自動選擇]: " -NoNewline -ForegroundColor Yellow
        
        if ([console]::KeyAvailable) {
            $key = [console]::ReadKey($true)
            $choice = $key.KeyChar
            
            if ($choice -match $ValidRegex) {
                $validChoice = $true
                Write-Host "`n"
            } else {
                Write-Host "`n`n[!] 無效的輸入 '$choice'，請重新輸入。" -ForegroundColor Red
                $timer.Restart()
            }
        }
        Start-Sleep -Milliseconds 100
    }
    $timer.Stop()
    return $choice
}

# ==========================================
# 4. 第一階選單 (主選單：選擇版本)
# ==========================================
$mainOptions = @(
    "1. 倉頡五代 (Cangjie 5) [預設]",
    "2. 倉頡三代 (Cangjie 3)",
    "0. 退出程式"
)
$mainChoice = Show-CountdownMenu -Title "請選擇要更新的倉頡版本：" -Options $mainOptions -PromptText "輸入代碼 (0-2)" -DefaultChoice '1' -ValidRegex '^[0-2]$'

if ($mainChoice -eq '0') {
    Write-Host "已取消執行，即將關閉視窗。" -ForegroundColor Yellow
    Start-Sleep -Seconds 2
    exit 
}

# ==========================================
# 5. 第二階選單 (子選單：選擇偏好) 與下載邏輯
# ==========================================
Write-Host "[*] 正在建立工作目錄..." -ForegroundColor Cyan
if (Test-Path $tempWork) { Remove-Item -Path $tempWork -Recurse -Force }
New-Item -ItemType Directory -Path $tempWork | Out-Null

$sourceLex = $null
$prefVersion = ""
$versionTag = ""

if ($mainChoice -eq '1') {
    # ---------- 倉頡五代邏輯 ----------
    $versionTag = "CJ5"
    $subOptions = @(
        "1. 一般排序",
        "2. 傳統漢字優先（偏好台灣用字習慣） [預設]",
        "3. 傳統漢字優先（偏好香港用字習慣）",
        "4. 簡化字優先",
        "0. 退出程式"
    )
    $subChoice = Show-CountdownMenu -Title "【倉頡五代】請選擇碼表版本：" -Options $subOptions -PromptText "輸入代碼 (0-4)" -DefaultChoice '2' -ValidRegex '^[0-4]$'
    
    if ($subChoice -eq '0') { exit }
    switch ($subChoice) {
        '1' { $prefVersion = "一般排序" }
        '2' { $prefVersion = "傳統漢字優先（偏好台灣用字習慣）" }
        '3' { $prefVersion = "傳統漢字優先（偏好香港用字習慣）" }
        '4' { $prefVersion = "簡化字優先" }
    }
    Write-Host "[*] 已選定五代版本：「$prefVersion」" -ForegroundColor Green

    $zipUrl = "https://github.com/Jackchows/Cangjie5/releases/download/v4.1-beta/MSCJData_Cangjie5_20260207.zip"
    $zipFile = "$tempWork\CJ5.zip"

    Write-Host "[*] 正在下載與解壓縮五代碼表..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -ErrorAction Stop
    Expand-Archive -Path $zipFile -DestinationPath "$tempWork\Extracted" -Force

    $sourceLexObj = Get-ChildItem -Path "$tempWork\Extracted" -Recurse -Filter "ChtCangjieExt.lex" |
                    Where-Object { $_.DirectoryName -like "*$prefVersion*" } | Select-Object -First 1
    if ($sourceLexObj) { $sourceLex = $sourceLexObj.FullName }

} elseif ($mainChoice -eq '2') {
    # ---------- 倉頡三代邏輯 ----------
    $versionTag = "CJ3"
    $subOptions = @(
        "1. Windows 10 2004 及之后的Windows [預設]",
        "2. Windows 10 2004 之前的版本",
        "0. 退出程式"
    )
    $subChoice = Show-CountdownMenu -Title "【倉頡三代】請選擇碼表版本：" -Options $subOptions -PromptText "輸入代碼 (0-2)" -DefaultChoice '1' -ValidRegex '^[0-2]$'
    
    if ($subChoice -eq '0') { exit }
    switch ($subChoice) {
        '1' { $prefVersion = "Windows 10 2004及之后的Windows" }
        '2' { $prefVersion = "Windows 10 2004之前的版本" }
    }
    Write-Host "[*] 已選定三代版本：「$prefVersion」" -ForegroundColor Green

    $url7z = "https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z"
    $url7za = "https://github.com/mcmilk/7-Zip-zstd/releases/download/v22.01-v1.5.2-R1/7za.exe"
    $file7z = "$tempWork\CJ3.7z"
    $exe7za = "$tempWork\7za.exe"

    Write-Host "[*] 正在下載 7-Zip 工具與三代碼表..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $url7za -OutFile $exe7za -ErrorAction Stop
    Invoke-WebRequest -Uri $url7z -OutFile $file7z -ErrorAction Stop

    Write-Host "[*] 正在從 7z 提取指定碼表..." -ForegroundColor Cyan
    $internalFile = "$prefVersion\ChtCangjieExt.lex"
    $sourceLex = "$tempWork\ChtCangjieExt.lex"
    
    $7zArgs = "e `"$file7z`" -o`"$tempWork`" `"$internalFile`" -r -y"
    Start-Process -FilePath $exe7za -ArgumentList $7zArgs -Wait -NoNewWindow
}

# 檢查檔案是否成功提取
if (!(Test-Path $sourceLex)) {
    Write-Host "[!] 提取失敗！找不到指定的碼表檔案。" -ForegroundColor Red
    Pause
    exit
}

# ==========================================
# 6. 共用執行區塊 (備份、替換、登錄檔)
# ==========================================
Write-Host "[*] 正在停止 Microsoft IME 進程..." -ForegroundColor Cyan
Stop-Process -Name "ctfmon", "MicrosoftIME" -Force -ErrorAction SilentlyContinue

Write-Host "[*] 正在備份與清理舊檔案..." -ForegroundColor Cyan
$dateStr = Get-Date -Format "yyyyMMdd"
$backupDir = "$targetDir\Backup_${versionTag}_$dateStr"

if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir | Out-Null }

$filesToProcess = @("ChtCangjie.sdc", "ChtCangjie.spd", "ChtCangjieExt.lex")
foreach ($file in $filesToProcess) {
    $fullPath = "$targetDir\$file"
    if (Test-Path $fullPath) {
        Copy-Item -Path $fullPath -Destination $backupDir -Force
        Remove-Item -Path $fullPath -Force
    }
}

Write-Host "[*] 正在部署新碼表..." -ForegroundColor Cyan
Copy-Item -Path $sourceLex -Destination "$targetDir\ChtCangjieExt.lex" -Force

Write-Host "[*] 正在自動開啟 HKSCS 選項 (修改登錄檔)..." -ForegroundColor Cyan
$regPath = "HKCU:\Software\Microsoft\IME\15.0\CHT\Cangjie"
if (!(Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
Set-ItemProperty -Path $regPath -Name "Enable HKSCS" -Value 1 -Type DWord -Force

# ==========================================
# 7. 清理與完成
# ==========================================
Write-Host "`n==================================================" -ForegroundColor Green
Write-Host " 更新成功！ ($versionTag)" -ForegroundColor Green
Write-Host " 已選用版本：$prefVersion" -ForegroundColor Green
Write-Host " 舊檔已備份至：$backupDir" -ForegroundColor Green
Write-Host ""
Write-Host " 請「立即重新啟動電腦」以讓碼表完全生效。" -ForegroundColor Yellow
Write-Host "==================================================`n"

Write-Host "[*] 清理暫存檔案..." -ForegroundColor Cyan
Remove-Item -Path $tempWork -Recurse -Force

Pause
Start-Process ctfmon.exe
# File: Cangjie3/AutoUpdate-CJ3.ps1
# Windows Cangjie 3 Dictionary Updater
# Requires: Windows PowerShell 5.1 and 7-Zip
# Source: https://github.com/Arthurmcarthur/Cangjie3-Plus

param(
    [string]$OriginalUserSid = ''
)

$ErrorActionPreference = 'Stop'

# ==========================================
# 1. Administrator Privileges
# ==========================================

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($identity)

if ([string]::IsNullOrWhiteSpace($OriginalUserSid)) {
    $OriginalUserSid = $identity.User.Value
}

$isAdministrator = $principal.IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator
)

if (-not $isAdministrator) {
    Write-Host "正在請求系統管理員權限..." -ForegroundColor Yellow

    try {
        $arguments = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -OriginalUserSid "{1}"' -f `
            $PSCommandPath, $OriginalUserSid

        Start-Process `
            -FilePath 'powershell.exe' `
            -ArgumentList $arguments `
            -Verb RunAs `
            -ErrorAction Stop

        exit
    }
    catch {
        Write-Host "無法取得系統管理員權限。" -ForegroundColor Red
        Write-Host $_.Exception.Message
        Read-Host "按 Enter 結束"
        exit 1
    }
}

$Host.UI.RawUI.WindowTitle = "倉頡三代碼表自動更新工具"

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ==========================================
# 2. Configuration
# ==========================================

$downloadUrl = 'https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z'

# File size of the pinned GitHub Release asset.
$expectedArchiveSize = 613787

$targetDir = Join-Path $env:SystemRoot 'System32\zh-hk'

$tempWork = Join-Path $env:TEMP (
    'CJ3_Update_' + [guid]::NewGuid().ToString('N')
)

$archiveFile = Join-Path $tempWork 'CJ3.7z'
$extractDir = Join-Path $tempWork 'Extracted'

$targetFiles = @(
    'ChtCangjie.sdc',
    'ChtCangjie.spd',
    'ChtCangjieExt.lex'
)

$backupDir = $null
$originalFiles = @{}
$deploymentStarted = $false
$imeStopped = $false
$regModified = $false
$oldRegExists = $false
$oldRegValue = $null
$regPath = $null

# ==========================================
# 3. Countdown Selection
# ==========================================

function Show-CountdownMenu {
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$DefaultChoice,
        [string]$ValidRegex,
        [int]$Timeout = 10
    )

    Write-Host ""
    Write-Host "==========================================" -ForegroundColor Cyan
    Write-Host " $Title" -ForegroundColor Cyan
    Write-Host "==========================================" -ForegroundColor Cyan

    foreach ($option in $Options) {
        Write-Host " $option"
    }

    Write-Host "==========================================" -ForegroundColor Cyan

    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    while ($true) {
        if ($timer.Elapsed.TotalSeconds -ge $Timeout) {
            Write-Host ""
            Write-Host "倒數結束，自動選擇預設值：$DefaultChoice" -ForegroundColor Yellow
            return $DefaultChoice
        }

        $remaining = $Timeout - [math]::Floor(
            $timer.Elapsed.TotalSeconds
        )

        Write-Host "`r請輸入選項（剩餘 $remaining 秒）：" `
            -NoNewline -ForegroundColor Yellow

        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $choice = [string]$key.KeyChar

            if ($choice -match $ValidRegex) {
                Write-Host ""
                return $choice
            }

            Write-Host ""
            Write-Host "無效選項：$choice" -ForegroundColor Red
            $timer.Restart()
        }

        Start-Sleep -Milliseconds 100
    }
}

# ==========================================
# 4. Select Cangjie Version
# ==========================================

$options = @(
    '1. Windows 10 2004 及之後版本 [預設]',
    '2. Windows 10 2004 之前版本',
    '0. 退出程式'
)

$choice = Show-CountdownMenu `
    -Title '請選擇倉頡三代碼表版本' `
    -Options $options `
    -DefaultChoice '1' `
    -ValidRegex '^[012]$' `
    -Timeout 10

if ($choice -eq '0') {
    Write-Host "已取消更新。" -ForegroundColor Yellow
    exit 0
}

switch ($choice) {
    '1' {
        $prefVersion = 'Windows 10 2004及之后的Windows'
    }
    '2' {
        $prefVersion = 'Windows 10 2004之前的版本'
    }
}

Write-Host ""
Write-Host "已選擇：$prefVersion" -ForegroundColor Green

# ==========================================
# 5. Find Installed 7-Zip
# ==========================================

function Find-7Zip {
    $candidates = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    $command = Get-Command '7z.exe' -ErrorAction SilentlyContinue

    if ($command) {
        return $command.Source
    }

    return $null
}

# ==========================================
# 6. Restore Original Files on Failure
# ==========================================

function Restore-OriginalFiles {
    param(
        [string]$Destination,
        [string]$Backup,
        [string[]]$Files,
        [hashtable]$OriginalState
    )

    Write-Host ""
    Write-Host "正在嘗試還原原始碼表..." -ForegroundColor Yellow

    foreach ($file in $Files) {
        $destinationFile = Join-Path $Destination $file
        $backupFile = Join-Path $Backup $file

        try {
            if ($OriginalState[$file]) {
                if (-not (Test-Path -LiteralPath $backupFile)) {
                    throw "找不到備份檔案：$backupFile"
                }

                Copy-Item `
                    -LiteralPath $backupFile `
                    -Destination $destinationFile `
                    -Force `
                    -ErrorAction Stop
            }
            else {
                if (Test-Path -LiteralPath $destinationFile) {
                    Remove-Item `
                        -LiteralPath $destinationFile `
                        -Force `
                        -ErrorAction Stop
                }
            }

            Write-Host "已還原：$file" -ForegroundColor Green
        }
        catch {
            Write-Host "還原失敗：$file" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
}

# ==========================================
# 7. Main Update Process
# ==========================================

try {
    # --------------------------------------
    # 7.1 Check 7-Zip
    # --------------------------------------

    Write-Host ""
    Write-Host "[1/8] 檢查 7-Zip..." -ForegroundColor Cyan

    $exe7z = Find-7Zip

    if (-not $exe7z) {
        throw @"
找不到 7-Zip。

請先安裝官方 7-Zip：
https://www.7-zip.org/

安裝完成後重新執行本工具。
"@
    }

    Write-Host "找到 7-Zip：$exe7z" -ForegroundColor Green

    # --------------------------------------
    # 7.2 Validate Target Directory
    # --------------------------------------

    Write-Host ""
    Write-Host "[2/8] 檢查 Windows 輸入法目錄..." -ForegroundColor Cyan

    if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
        throw "找不到系統輸入法目錄：$targetDir"
    }

    Write-Host "目標目錄：$targetDir" -ForegroundColor Green

    # Validate the original user's registry hive
    # before modifying any system files.

    $userHive = "Registry::HKEY_USERS\$OriginalUserSid"

    if (-not (Test-Path -LiteralPath $userHive)) {
        throw "找不到原始使用者的登錄檔 Hive：$OriginalUserSid"
    }

    $regPath = "$userHive\Software\Microsoft\IME\15.0\CHT\Cangjie"

    # --------------------------------------
    # 7.3 Download Archive
    # --------------------------------------

    Write-Host ""
    Write-Host "[3/8] 下載倉頡三代碼表..." -ForegroundColor Cyan

    New-Item `
        -ItemType Directory `
        -Path $tempWork `
        -Force `
        -ErrorAction Stop | Out-Null

    Invoke-WebRequest `
        -Uri $downloadUrl `
        -OutFile $archiveFile `
        -UseBasicParsing `
        -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $archiveFile -PathType Leaf)) {
        throw "下載失敗：找不到壓縮檔。"
    }

    $archiveInfo = Get-Item -LiteralPath $archiveFile

    if ($archiveInfo.Length -ne $expectedArchiveSize) {
        throw (
            "下載檔案大小不符合預期。" +
            "預期：$expectedArchiveSize Bytes；" +
            "實際：$($archiveInfo.Length) Bytes。"
        )
    }

    Write-Host "下載完成。" -ForegroundColor Green

    # --------------------------------------
    # 7.4 Extract Archive
    # --------------------------------------

    Write-Host ""
    Write-Host "[4/8] 解壓縮指定碼表..." -ForegroundColor Cyan

    New-Item `
        -ItemType Directory `
        -Path $extractDir `
        -Force `
        -ErrorAction Stop | Out-Null

    & $exe7z x $archiveFile "-o$extractDir" -y

    if ($LASTEXITCODE -ne 0) {
        throw "7-Zip 解壓縮失敗，ExitCode：$LASTEXITCODE"
    }

    $sourceLexObj = Get-ChildItem `
        -LiteralPath $extractDir `
        -Recurse `
        -File `
        -Filter 'ChtCangjieExt.lex' `
        -ErrorAction Stop |
        Where-Object {
            $_.DirectoryName -like "*$prefVersion*"
        } |
        Select-Object -First 1

    if (-not $sourceLexObj) {
        throw "找不到指定版本的碼表：$prefVersion"
    }

    $sourceLex = $sourceLexObj.FullName

    if ($sourceLexObj.Length -eq 0) {
        throw "碼表檔案大小為 0，已取消更新。"
    }

    Write-Host "找到碼表：$sourceLex" -ForegroundColor Green
    Write-Host "檔案大小：$($sourceLexObj.Length) Bytes"

    # --------------------------------------
    # 7.5 Back Up Original Files
    # --------------------------------------

    Write-Host ""
    Write-Host "[5/8] 備份原始系統碼表..." -ForegroundColor Cyan

    $dateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
    $uniqueId = [guid]::NewGuid().ToString('N').Substring(0, 8)

    $backupDir = Join-Path $targetDir (
        "Backup_CJ3_${dateStr}_${uniqueId}"
    )

    New-Item `
        -ItemType Directory `
        -Path $backupDir `
        -ErrorAction Stop | Out-Null

    foreach ($file in $targetFiles) {
        $originalPath = Join-Path $targetDir $file
        $backupPath = Join-Path $backupDir $file

        $exists = Test-Path -LiteralPath $originalPath -PathType Leaf
        $originalFiles[$file] = $exists

        if ($exists) {
            Copy-Item `
                -LiteralPath $originalPath `
                -Destination $backupPath `
                -Force `
                -ErrorAction Stop

            if (-not (Test-Path -LiteralPath $backupPath)) {
                throw "備份失敗：$file"
            }

            $originalSize = (Get-Item -LiteralPath $originalPath).Length
            $backupSize = (Get-Item -LiteralPath $backupPath).Length

            if ($originalSize -ne $backupSize) {
                throw "備份檔案大小不一致：$file"
            }

            Write-Host "已備份：$file" -ForegroundColor Green
        }
        else {
            Write-Host "原始檔案不存在，略過：$file" -ForegroundColor Yellow
        }
    }

    Write-Host "備份目錄：$backupDir" -ForegroundColor Green

    # --------------------------------------
    # 7.6 Stop IME Processes
    # --------------------------------------

    Write-Host ""
    Write-Host "[6/8] 停止輸入法相關程序..." -ForegroundColor Cyan

    Stop-Process `
        -Name 'ctfmon', 'MicrosoftIME' `
        -Force `
        -ErrorAction SilentlyContinue

    $imeStopped = $true

    Start-Sleep -Milliseconds 800

    # --------------------------------------
    # 7.7 Deploy Dictionary
    # --------------------------------------

    Write-Host ""
    Write-Host "[7/8] 部署倉頡三代碼表..." -ForegroundColor Cyan

    $deploymentStarted = $true

    foreach ($file in $targetFiles) {
        $targetFile = Join-Path $targetDir $file

        if (Test-Path -LiteralPath $targetFile) {
            Remove-Item `
                -LiteralPath $targetFile `
                -Force `
                -ErrorAction Stop
        }
    }

    $destinationLex = Join-Path $targetDir 'ChtCangjieExt.lex'

    Copy-Item `
        -LiteralPath $sourceLex `
        -Destination $destinationLex `
        -Force `
        -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $destinationLex -PathType Leaf)) {
        throw "碼表部署失敗：目標檔案不存在。"
    }

    $sourceHash = (Get-FileHash `
        -LiteralPath $sourceLex `
        -Algorithm SHA256).Hash

    $destinationHash = (Get-FileHash `
        -LiteralPath $destinationLex `
        -Algorithm SHA256).Hash

    if ($sourceHash -ne $destinationHash) {
        throw "碼表部署失敗：來源與目標檔案雜湊值不一致。"
    }

    Write-Host "碼表部署及 SHA-256 比對成功。" -ForegroundColor Green

    # --------------------------------------
    # 7.8 Enable HKSCS
    # --------------------------------------

    Write-Host ""
    Write-Host "[8/8] 啟用 HKSCS 支援..." -ForegroundColor Cyan

    if (Test-Path -LiteralPath $regPath) {
        $existingReg = Get-ItemProperty `
            -LiteralPath $regPath `
            -Name 'Enable HKSCS' `
            -ErrorAction SilentlyContinue

        if ($null -ne $existingReg) {
            $oldRegExists = $true
            $oldRegValue = $existingReg.'Enable HKSCS'
        }
    }

    if (-not (Test-Path -LiteralPath $regPath)) {
        New-Item `
            -Path $regPath `
            -Force `
            -ErrorAction Stop | Out-Null
    }

    $regModified = $true

    New-ItemProperty `
        -LiteralPath $regPath `
        -Name 'Enable HKSCS' `
        -Value 1 `
        -PropertyType DWord `
        -Force `
        -ErrorAction Stop | Out-Null

    $actualValue = (
        Get-ItemProperty `
            -LiteralPath $regPath `
            -Name 'Enable HKSCS' `
            -ErrorAction Stop
    ).'Enable HKSCS'

    if ($actualValue -ne 1) {
        throw "HKSCS 登錄檔設定驗證失敗。"
    }

    Write-Host "HKSCS 已啟用。" -ForegroundColor Green

    # ======================================
    # 8. Completed
    # ======================================

    $deploymentStarted = $false
    $regModified = $false

    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Green
    Write-Host " 倉頡三代碼表更新成功！" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "已安裝版本：$prefVersion"
    Write-Host "碼表位置：$destinationLex"
    Write-Host "原始備份：$backupDir"
    Write-Host ""
    Write-Host "請重新啟動 Windows，使輸入法設定完全生效。" -ForegroundColor Yellow
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host " 更新失敗！" -ForegroundColor Red
    Write-Host "==================================================" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red

    # Restore original dictionary files if deployment
    # had already started.

    if ($deploymentStarted -and $backupDir) {
        Restore-OriginalFiles `
            -Destination $targetDir `
            -Backup $backupDir `
            -Files $targetFiles `
            -OriginalState $originalFiles
    }

    # Attempt to restore the previous registry value.

    if ($regModified -and $regPath) {
        try {
            if ($oldRegExists) {
                New-ItemProperty `
                    -LiteralPath $regPath `
                    -Name 'Enable HKSCS' `
                    -Value $oldRegValue `
                    -PropertyType DWord `
                    -Force `
                    -ErrorAction Stop | Out-Null
            }
            else {
                Remove-ItemProperty `
                    -LiteralPath $regPath `
                    -Name 'Enable HKSCS' `
                    -ErrorAction Stop
            }

            Write-Host "HKSCS 登錄檔設定已還原。" -ForegroundColor Green
        }
        catch {
            Write-Host "HKSCS 登錄檔還原失敗。" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }

    if ($backupDir) {
        Write-Host ""
        Write-Host "原始備份保留於：$backupDir" -ForegroundColor Yellow
    }

    Write-Host "若還原不完整，請勿重複執行，先檢查備份檔案。" -ForegroundColor Yellow
    Write-Host ""

    Read-Host "按 Enter 結束"
    exit 1
}
finally {
    # ======================================
    # 9. Cleanup
    # ======================================

    if ($tempWork -and (Test-Path -LiteralPath $tempWork)) {
        try {
            Remove-Item `
                -LiteralPath $tempWork `
                -Recurse `
                -Force `
                -ErrorAction Stop
        }
        catch {
            Write-Host "暫存目錄清理失敗：$tempWork" -ForegroundColor Yellow
        }
    }

    if ($imeStopped) {
        try {
            Start-Process `
                -FilePath "$env:SystemRoot\System32\ctfmon.exe" `
                -ErrorAction Stop
        }
        catch {
            Write-Host "無法重新啟動 ctfmon.exe，請重新啟動 Windows。" -ForegroundColor Yellow
        }
    }
}

Read-Host "按 Enter 結束"
exit 0

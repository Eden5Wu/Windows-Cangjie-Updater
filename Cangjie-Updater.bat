@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "SELF=%~f0"

net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ==================================================
    echo  [!] 需要系統管理員權限
    echo ==================================================
    echo.
    echo  請關閉此視窗，對「%~nx0」按右鍵，
    echo  選擇【以系統管理員身分執行】。
    echo.
    echo ==================================================
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "$c = [System.IO.File]::ReadAllText($env:SELF, [System.Text.Encoding]::UTF8); $m = '<#PS_' + 'START#>'; $i = $c.IndexOf($m); if ($i -lt 0) { throw '找不到 PowerShell 程式區塊。' }; [scriptblock]::Create($c.Substring($i + $m.Length)).Invoke()"
exit /b %errorlevel%

<#PS_START#>
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$Host.UI.RawUI.WindowTitle = 'Windows 倉頡碼表自動更新工具'
$tempWork = Join-Path $env:TEMP 'CJ_Update_All'

# ==========================================
# 共用輔助函式
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
    foreach ($opt in $Options) {
        Write-Host "  $opt"
    }
    Write-Host "==========================================`n" -ForegroundColor Cyan

    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $choice = $null
    $validChoice = $false
    $lastTimeLeft = -1

    while (-not $validChoice) {
        if ($timer.Elapsed.TotalSeconds -ge $Timeout) {
            $choice = $DefaultChoice
            $validChoice = $true
            Write-Host "`n`n[!] 倒數結束，自動選擇預設值 ($DefaultChoice)。" -ForegroundColor Yellow
            break
        }

        $timeLeft = $Timeout - [math]::Floor($timer.Elapsed.TotalSeconds)
        if ($timeLeft -ne $lastTimeLeft) {
            $lastTimeLeft = $timeLeft
            Write-Host "`r$PromptText [倒數 $timeLeft 秒後自動選擇]:   " -NoNewline -ForegroundColor Yellow
        }

        if ([console]::KeyAvailable) {
            $key = [console]::ReadKey($true)
            if ($key.Key -eq [System.ConsoleKey]::Enter) {
                $choice = $DefaultChoice
                $validChoice = $true
                Write-Host "`n"
            }
            else {
                $choice = $key.KeyChar
                if ($choice -match $ValidRegex) {
                    $validChoice = $true
                    Write-Host "`n"
                }
                else {
                    $displayChar = if ([char]::IsControl($choice)) { $key.Key.ToString() } else { [string]$choice }
                    Write-Host "`n`n[!] 無效的輸入 '$displayChar'，請重新輸入。" -ForegroundColor Red
                    $timer.Restart()
                    $lastTimeLeft = -1
                }
            }
        }
        Start-Sleep -Milliseconds 100
    }

    $timer.Stop()
    return $choice
}

function Find-SevenZip {
    $candidates = @(
        "$env:ProgramFiles\7-Zip\7z.exe",
        "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
    )
    $found = $candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $found) {
        $cmd = Get-Command '7z.exe' -ErrorAction SilentlyContinue
        if ($cmd) { $found = $cmd.Source }
    }
    return $found
}

function Get-DesktopUserCangjieRegPath {
    $sessionId = (Get-Process -Id $PID).SessionId
    $userSids = @(
        Get-CimInstance Win32_Process -Filter "Name = 'explorer.exe'" |
            Where-Object { $_.SessionId -eq $sessionId } |
            ForEach-Object {
                $owner = Invoke-CimMethod -InputObject $_ -MethodName 'GetOwnerSid'
                if (($owner.ReturnValue -eq 0) -and $owner.Sid) { $owner.Sid }
            } |
            Select-Object -Unique
    )

    if ($userSids.Count -ne 1) {
        throw '無法唯一確認目前桌面登入使用者，已取消更新。'
    }

    $userHive = "Registry::HKEY_USERS\$($userSids[0])"
    if (-not (Test-Path -LiteralPath $userHive)) {
        throw "找不到使用者登錄檔：$($userSids[0])"
    }

    return "$userHive\Software\Microsoft\IME\15.0\CHT\Cangjie"
}

function Backup-TargetFiles {
    param(
        [string]$TargetDir,
        [string]$BackupDir,
        [string[]]$FileNames
    )

    New-Item -ItemType Directory -Path $BackupDir -Force -ErrorAction Stop | Out-Null
    $originalMap = @{}

    foreach ($name in $FileNames) {
        $origPath = Join-Path $TargetDir $name
        $bakPath = Join-Path $BackupDir $name
        $exists = Test-Path -LiteralPath $origPath -PathType Leaf
        $originalMap[$name] = $exists

        if ($exists) {
            Copy-Item -LiteralPath $origPath -Destination $bakPath -Force -ErrorAction Stop
            $origHash = (Get-FileHash -LiteralPath $origPath -Algorithm SHA256).Hash
            $bakHash = (Get-FileHash -LiteralPath $bakPath -Algorithm SHA256).Hash
            if ($origHash -ne $bakHash) {
                throw "備份驗證失敗：$name"
            }
            Write-Host "[OK] 已備份：$name" -ForegroundColor Green
        }
        else {
            Write-Host "[INFO] 原始檔案不存在：$name" -ForegroundColor Yellow
        }
    }

    return $originalMap
}

function Restore-TargetFiles {
    param(
        [string]$TargetDir,
        [string]$BackupDir,
        [string[]]$FileNames,
        [hashtable]$OriginalMap
    )

    Write-Host "`n[*] 正在嘗試還原原始碼表..." -ForegroundColor Yellow
    foreach ($name in $FileNames) {
        $targetPath = Join-Path $TargetDir $name
        $bakPath = Join-Path $BackupDir $name

        try {
            if ($OriginalMap[$name]) {
                if (-not (Test-Path -LiteralPath $bakPath -PathType Leaf)) {
                    throw "找不到備份檔案：$name"
                }
                Copy-Item -LiteralPath $bakPath -Destination $targetPath -Force -ErrorAction Stop
                $restoredHash = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash
                $savedHash = (Get-FileHash -LiteralPath $bakPath -Algorithm SHA256).Hash
                if ($restoredHash -ne $savedHash) {
                    throw "還原驗證失敗：$name"
                }
            }
            elseif (Test-Path -LiteralPath $targetPath) {
                Remove-Item -LiteralPath $targetPath -Force -ErrorAction Stop
            }
            Write-Host "[OK] 已還原：$name" -ForegroundColor Green
        }
        catch {
            Write-Host "[ERROR] 還原失敗：$name ($($_.Exception.Message))" -ForegroundColor Red
        }
    }
}

function Deploy-FileMap {
    param(
        [string]$TargetDir,
        [string[]]$RemoveNames,
        [hashtable]$SourceMap
    )

    foreach ($name in $RemoveNames) {
        $path = Join-Path $TargetDir $name
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force -ErrorAction Stop
        }
    }

    foreach ($name in $SourceMap.Keys) {
        $srcPath = $SourceMap[$name]
        $dstPath = Join-Path $TargetDir $name

        if (-not (Test-Path -LiteralPath $srcPath -PathType Leaf)) {
            throw "找不到來源檔案：$srcPath"
        }

        Copy-Item -LiteralPath $srcPath -Destination $dstPath -Force -ErrorAction Stop

        $srcHash = (Get-FileHash -LiteralPath $srcPath -Algorithm SHA256).Hash
        $dstHash = (Get-FileHash -LiteralPath $dstPath -Algorithm SHA256).Hash
        if ($srcHash -ne $dstHash) {
            throw "碼表部署驗證失敗：$name"
        }
        Write-Host "[OK] 已部署並驗證：$name" -ForegroundColor Green
    }
}

# ==========================================
# 第一階選單：選擇倉頡版本
# ==========================================

$mainOptions = @(
    '1. 倉頡五代 (Cangjie 5) [預設]',
    '2. 倉頡三代 (Cangjie 3 - v4.3)',
    '0. 退出程式'
)
$mainChoice = Show-CountdownMenu -Title '請選擇要更新的倉頡版本：' -Options $mainOptions -PromptText '輸入代碼 (0-2)' -DefaultChoice '1' -ValidRegex '^[0-2]$'

if ($mainChoice -eq '0') {
    Write-Host '已取消執行。' -ForegroundColor Yellow
    return
}

if (Test-Path -LiteralPath $tempWork) {
    Remove-Item -LiteralPath $tempWork -Recurse -Force -ErrorAction SilentlyContinue
}
New-Item -ItemType Directory -Path $tempWork -Force | Out-Null

$imeStopped = $false

try {
    if ($mainChoice -eq '1') {
        # ==========================================
        # 倉頡五代 (CJ5) 流程
        # ==========================================
        $subOptions = @(
            '1. 一般排序',
            '2. 傳統漢字優先（偏好台灣用字習慣） [預設]',
            '3. 傳統漢字優先（偏好香港用字習慣）',
            '4. 簡化字優先',
            '0. 退出程式'
        )
        $subChoice = Show-CountdownMenu -Title '【倉頡五代】請選擇碼表版本：' -Options $subOptions -PromptText '輸入代碼 (0-4)' -DefaultChoice '2' -ValidRegex '^[0-4]$'
        if ($subChoice -eq '0') { return }

        $prefVersion = switch ($subChoice) {
            '1' { '一般排序' }
            '2' { '傳統漢字優先（偏好台灣用字習慣）' }
            '3' { '傳統漢字優先（偏好香港用字習慣）' }
            '4' { '簡化字優先' }
        }
        Write-Host "[*] 已選定五代版本：$prefVersion" -ForegroundColor Green

        $targetDir = Join-Path $env:SystemRoot 'System32\zh-hk'
        if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
            throw "找不到系統輸入法目錄：$targetDir。請確認 Windows 已新增「中文（繁體）」語言與內建「微軟倉頡」輸入法。"
        }

        $regPath = Get-DesktopUserCangjieRegPath
        $zipUrl = 'https://github.com/Jackchows/Cangjie5/releases/download/v4.1-beta/MSCJData_Cangjie5_20260207.zip'
        $zipFile = Join-Path $tempWork 'CJ5.zip'
        $extractDir = Join-Path $tempWork 'CJ5_Extracted'

        Write-Host "`n[*] 正在下載與解壓縮倉頡五代碼表..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing -ErrorAction Stop
        Expand-Archive -LiteralPath $zipFile -DestinationPath $extractDir -Force

        $sourceLexObj = Get-ChildItem -LiteralPath $extractDir -Recurse -File -Filter 'ChtCangjieExt.lex' |
            Where-Object { $_.DirectoryName -like "*$prefVersion*" } |
            Select-Object -First 1

        if (-not $sourceLexObj -or $sourceLexObj.Length -eq 0) {
            throw "提取失敗！找不到指定版本的碼表：$prefVersion"
        }

        $filesToProcess = @('ChtCangjie.sdc', 'ChtCangjie.spd', 'ChtCangjieExt.lex')
        $dateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
        $uid = [guid]::NewGuid().ToString('N').Substring(0, 8)
        $backupDir = Join-Path $targetDir "Backup_CJ5_${dateStr}_${uid}"

        $deploymentStarted = $false
        $registryTouched = $false
        $oldRegExists = $false
        $oldRegValue = $null
        $originalMap = $null

        try {
            Write-Host "`n[*] 正在備份原始系統碼表..." -ForegroundColor Cyan
            $originalMap = Backup-TargetFiles -TargetDir $targetDir -BackupDir $backupDir -FileNames $filesToProcess

            $oldReg = Get-ItemProperty -LiteralPath $regPath -Name 'Enable HKSCS' -ErrorAction SilentlyContinue
            if ($null -ne $oldReg) {
                $oldRegExists = $true
                $oldRegValue = $oldReg.'Enable HKSCS'
            }

            Write-Host "`n[*] 正在停止 Microsoft IME 程序..." -ForegroundColor Cyan
            Stop-Process -Name 'ctfmon', 'MicrosoftIME' -Force -ErrorAction SilentlyContinue
            $imeStopped = $true
            Start-Sleep -Milliseconds 800

            Write-Host "`n[*] 正在部署倉頡五代碼表..." -ForegroundColor Cyan
            $deploymentStarted = $true
            Deploy-FileMap -TargetDir $targetDir -RemoveNames $filesToProcess -SourceMap @{
                'ChtCangjieExt.lex' = $sourceLexObj.FullName
            }

            Write-Host "`n[*] 正在啟用 HKSCS 支援..." -ForegroundColor Cyan
            $registryTouched = $true
            if (-not (Test-Path -LiteralPath $regPath)) {
                New-Item -Path $regPath -Force -ErrorAction Stop | Out-Null
            }
            New-ItemProperty -LiteralPath $regPath -Name 'Enable HKSCS' -Value 1 -PropertyType DWord -Force -ErrorAction Stop | Out-Null

            $deploymentStarted = $false
            $registryTouched = $false

            Write-Host "`n==================================================" -ForegroundColor Green
            Write-Host ' 倉頡五代碼表更新成功！' -ForegroundColor Green
            Write-Host " 已選用版本：$prefVersion" -ForegroundColor Green
            Write-Host " 原始備份：$backupDir" -ForegroundColor Green
            Write-Host ''
            Write-Host ' 請「重新啟動電腦」以讓碼表完全生效。' -ForegroundColor Yellow
            Write-Host '==================================================' -ForegroundColor Green
        }
        catch {
            Write-Host "`n[ERROR] 倉頡五代更新失敗：$($_.Exception.Message)" -ForegroundColor Red
            if ($deploymentStarted -and $originalMap) {
                Restore-TargetFiles -TargetDir $targetDir -BackupDir $backupDir -FileNames $filesToProcess -OriginalMap $originalMap
            }
            if ($registryTouched) {
                if ($oldRegExists) {
                    New-ItemProperty -LiteralPath $regPath -Name 'Enable HKSCS' -Value $oldRegValue -PropertyType DWord -Force -ErrorAction SilentlyContinue | Out-Null
                }
                else {
                    Remove-ItemProperty -LiteralPath $regPath -Name 'Enable HKSCS' -ErrorAction SilentlyContinue
                }
            }
            throw
        }
    }
    elseif ($mainChoice -eq '2') {
        # ==========================================
        # 倉頡三代 (CJ3 v4.3) 流程
        # ==========================================
        $winOptions = @(
            '1. Windows 10 2004 及之後版本 / Windows 11 [預設]',
            '2. Windows 10 2004 之前版本',
            '0. 退出程式'
        )
        $winChoice = Show-CountdownMenu -Title '【倉頡三代】請選擇 Windows 版本：' -Options $winOptions -PromptText '輸入代碼 (0-2)' -DefaultChoice '1' -ValidRegex '^[0-2]$'
        if ($winChoice -eq '0') { return }

        if ($winChoice -eq '1') {
            $winFolder = 'Windows_10_2004之後的版本及Windows_11'
            $targetDir = Join-Path $env:SystemRoot 'System32\zh-hk'
            $cj3Files = @('ChtCangjie.sdc', 'ChtCangjie.spd', 'ChtCangjieExt.lex')
        }
        else {
            $winFolder = 'Windows_10_2004之前的版本'
            $targetDir = Join-Path $env:SystemRoot 'InputMethod\CHT'
            $cj3Files = @('ChtChangjie.lex', 'ChtChangjie.spd', 'ChtChangjieExt.lex')
        }

        $charOptions = @(
            '1. 僅漢字 [預設]',
            '2. 包含部首筆畫兼容區漢字',
            '0. 退出程式'
        )
        $charChoice = Show-CountdownMenu -Title '【倉頡三代】請選擇字集版本：' -Options $charOptions -PromptText '輸入代碼 (0-2)' -DefaultChoice '1' -ValidRegex '^[0-2]$'
        if ($charChoice -eq '0') { return }

        $charFolder = if ($charChoice -eq '1') { '僅漢字' } else { '包含部首筆畫兼容區漢字' }
        Write-Host "[*] 已選定三代版本：$winFolder / $charFolder" -ForegroundColor Green

        Write-Host "`n[1/6] 檢查 7-Zip..." -ForegroundColor Cyan
        $sevenZip = Find-SevenZip
        if (-not $sevenZip) {
            throw '找不到 7-Zip。請先安裝：https://www.7-zip.org/'
        }
        Write-Host "[OK] $sevenZip" -ForegroundColor Green

        Write-Host "`n[2/6] 檢查系統目錄..." -ForegroundColor Cyan
        if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
            throw "找不到系統輸入法目錄：$targetDir。請確認 Windows 已新增「中文（繁體）」語言與內建「微軟倉頡」輸入法。"
        }
        Write-Host "[OK] 目標目錄：$targetDir" -ForegroundColor Green

        $cj3Url = 'https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.3/MSCJData_20260908_Cangjie3_WithExtJ.7z'
        $cj3ExpectedSize = 1663583
        $cj3ExpectedHash = '51704297d2a1618c4bce5239650f7bd7e7a528b12de02a9f1e33d83cae5beede'
        $cj3Archive = Join-Path $tempWork 'CJ3.7z'
        $cj3ExtractDir = Join-Path $tempWork 'CJ3_Extracted'

        Write-Host "`n[3/6] 下載並驗證倉頡三代 v4.3 碼表..." -ForegroundColor Cyan
        Invoke-WebRequest -Uri $cj3Url -OutFile $cj3Archive -UseBasicParsing -ErrorAction Stop

        $archiveInfo = Get-Item -LiteralPath $cj3Archive
        if ($archiveInfo.Length -ne $cj3ExpectedSize) {
            throw "壓縮檔大小不符：$($archiveInfo.Length) Bytes"
        }
        $actualHash = (Get-FileHash -LiteralPath $cj3Archive -Algorithm SHA256).Hash
        if ($actualHash -ine $cj3ExpectedHash) {
            throw '壓縮檔 SHA-256 驗證失敗，已取消更新。'
        }
        Write-Host '[OK] 下載完成，SHA-256 驗證成功。' -ForegroundColor Green

        Write-Host "`n[4/6] 解壓縮並定位碼表..." -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $cj3ExtractDir -Force | Out-Null
        & $sevenZip x $cj3Archive "-o$cj3ExtractDir" -y | Out-Null
        if ($LASTEXITCODE -ne 0) {
            throw "7-Zip 解壓縮失敗，ExitCode：$LASTEXITCODE"
        }

        $sourceDir = Join-Path $cj3ExtractDir "MSCJData_20260908_Cangjie3_WithExtJ\$winFolder\$charFolder"
        if (-not (Test-Path -LiteralPath $sourceDir -PathType Container)) {
            throw "找不到指定版本資料夾：$sourceDir"
        }

        $sourceMap = @{}
        foreach ($name in $cj3Files) {
            $p = Join-Path $sourceDir $name
            if (-not (Test-Path -LiteralPath $p -PathType Leaf) -or (Get-Item -LiteralPath $p).Length -eq 0) {
                throw "碼表檔案缺失或為空：$name"
            }
            $sourceMap[$name] = $p
        }
        Write-Host "[OK] 已定位全部 $($cj3Files.Count) 個碼表檔案。" -ForegroundColor Green

        $dateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
        $uid = [guid]::NewGuid().ToString('N').Substring(0, 8)
        $backupDir = Join-Path $targetDir "Backup_CJ3_${dateStr}_${uid}"
        $deploymentStarted = $false
        $originalMap = $null

        try {
            Write-Host "`n[5/6] 備份原始系統碼表..." -ForegroundColor Cyan
            $originalMap = Backup-TargetFiles -TargetDir $targetDir -BackupDir $backupDir -FileNames $cj3Files

            Write-Host "`n[*] 正在停止 Microsoft IME 程序..." -ForegroundColor Cyan
            Stop-Process -Name 'ctfmon', 'MicrosoftIME' -Force -ErrorAction SilentlyContinue
            $imeStopped = $true
            Start-Sleep -Milliseconds 800

            Write-Host "`n[6/6] 部署倉頡三代 v4.3 碼表..." -ForegroundColor Cyan
            $deploymentStarted = $true
            Deploy-FileMap -TargetDir $targetDir -RemoveNames $cj3Files -SourceMap $sourceMap

            $deploymentStarted = $false

            Write-Host "`n==================================================" -ForegroundColor Green
            Write-Host ' 倉頡三代 v4.3 碼表更新成功！' -ForegroundColor Green
            Write-Host " 已安裝版本：$winFolder / $charFolder" -ForegroundColor Green
            Write-Host " 目標目錄：$targetDir" -ForegroundColor Green
            Write-Host " 原始備份：$backupDir" -ForegroundColor Green
            Write-Host ''
            Write-Host ' 請「重新啟動電腦」以讓碼表完全生效。' -ForegroundColor Yellow
            Write-Host '==================================================' -ForegroundColor Green
        }
        catch {
            Write-Host "`n[ERROR] 倉頡三代更新失敗：$($_.Exception.Message)" -ForegroundColor Red
            if ($deploymentStarted -and $originalMap) {
                Restore-TargetFiles -TargetDir $targetDir -BackupDir $backupDir -FileNames $cj3Files -OriginalMap $originalMap
            }
            throw
        }
    }
}
catch {
    Write-Host "`n[!] 更新程序已中止：$($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if (Test-Path -LiteralPath $tempWork) {
        Remove-Item -LiteralPath $tempWork -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($imeStopped) {
        Start-Process 'ctfmon.exe' -ErrorAction SilentlyContinue
    }
    Write-Host ''
    Read-Host '按 Enter 鍵結束視窗'
}

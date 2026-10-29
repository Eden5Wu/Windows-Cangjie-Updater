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
    # ==========================================
    # 倉頡三代（CJ3）完整更新流程
    # 本分支自行完成部署並退出，
    # 不再執行後面的共用部署區塊。
    # CJ5 流程不受影響。
    # ==========================================

    $versionTag = 'CJ3'

    # ------------------------------------------
    # 1. 選擇版本
    # ------------------------------------------

    $subOptions = @(
        '1. Windows 10 2004 及之後版本 [預設]',
        '2. Windows 10 2004 之前版本',
        '0. 退出程式'
    )

    $subChoice = Show-CountdownMenu `
        -Title '【倉頡三代】請選擇碼表版本：' `
        -Options $subOptions `
        -PromptText '輸入代碼 (0-2)' `
        -DefaultChoice '1' `
        -ValidRegex '^[0-2]$'

    if ($subChoice -eq '0') {
        Write-Host '已取消更新。' -ForegroundColor Yellow

        if (Test-Path -LiteralPath $tempWork) {
            Remove-Item -LiteralPath $tempWork -Recurse -Force `
                -ErrorAction SilentlyContinue
        }

        exit 0
    }

    switch ($subChoice) {
        '1' {
            $prefVersion = 'Windows 10 2004及之后的Windows'
        }
        '2' {
            $prefVersion = 'Windows 10 2004之前的版本'
        }
    }

    Write-Host "[*] 已選定三代版本：$prefVersion" -ForegroundColor Green

    # ------------------------------------------
    # 2. 基本設定
    # ------------------------------------------

    $cj3PreviousErrorAction = $ErrorActionPreference
    $ErrorActionPreference = 'Stop'

    $cj3Url = 'https://github.com/Arthurmcarthur/Cangjie3-Plus/releases/download/4.2/MSCJData_20251014_Cangjie3_WithExtJ.7z'

    # GitHub Release 4.2 官方附件資訊
    $cj3ExpectedSize = 613787
    $cj3ExpectedHash = '342422de88f78a1a924042dad3b00689e23f1ce612e3a718243ebf7e070e7954'

    $cj3Archive = Join-Path $tempWork 'CJ3.7z'
    $cj3ExtractDir = Join-Path $tempWork 'CJ3_Extracted'

    $cj3Files = @(
        'ChtCangjie.sdc',
        'ChtCangjie.spd',
        'ChtCangjieExt.lex'
    )

    $cj3OriginalFiles = @{}
    $cj3BackupDir = $null
    $cj3SourceLex = $null

    $cj3DeploymentStarted = $false
    $cj3RegistryTouched = $false
    $cj3ImeStopped = $false
    $cj3Success = $false

    $cj3RegPath = $null
    $cj3OldRegExists = $false
    $cj3OldRegValue = $null

    try {
        # ======================================
        # 3. 檢查 7-Zip
        # ======================================

        Write-Host ''
        Write-Host '[1/8] 檢查 7-Zip...' -ForegroundColor Cyan

        $cj3SevenZip = @(
            "$env:ProgramFiles\7-Zip\7z.exe",
            "${env:ProgramFiles(x86)}\7-Zip\7z.exe"
        ) |
        Where-Object {
            $_ -and (Test-Path -LiteralPath $_ -PathType Leaf)
        } |
        Select-Object -First 1

        if (-not $cj3SevenZip) {
            $cj3Command = Get-Command '7z.exe' `
                -ErrorAction SilentlyContinue

            if ($cj3Command) {
                $cj3SevenZip = $cj3Command.Source
            }
        }

        if (-not $cj3SevenZip) {
            throw "找不到 7-Zip。請先安裝：https://www.7-zip.org/"
        }

        Write-Host "[OK] $cj3SevenZip" -ForegroundColor Green

        # ======================================
        # 4. 檢查系統環境與使用者
        # ======================================

        Write-Host ''
        Write-Host '[2/8] 檢查系統環境...' -ForegroundColor Cyan

        if (-not (Test-Path -LiteralPath $targetDir -PathType Container)) {
            throw "找不到系統輸入法目錄：$targetDir"
        }

        # 整合版原有 UAC 程式沒有傳遞使用者 SID。
        # 透過目前桌面工作階段的 Explorer，
        # 找出原始登入使用者，避免將 HKCU
        # 寫入另一個管理員帳戶。
        #
        # 無法唯一確認使用者時停止更新，
        # 避免修改錯誤帳戶的登錄檔。

        $cj3SessionId = (Get-Process -Id $PID).SessionId

        $cj3UserSids = @(
            Get-CimInstance Win32_Process `
                -Filter "Name = 'explorer.exe'" |
            Where-Object {
                $_.SessionId -eq $cj3SessionId
            } |
            ForEach-Object {
                $owner = Invoke-CimMethod `
                    -InputObject $_ `
                    -MethodName 'GetOwnerSid'

                if (($owner.ReturnValue -eq 0) -and $owner.Sid) {
                    $owner.Sid
                }
            } |
            Select-Object -Unique
        )

        if ($cj3UserSids.Count -ne 1) {
            throw '無法唯一確認目前桌面登入使用者，已取消更新。'
        }

        $cj3UserSid = $cj3UserSids[0]
        $cj3UserHive = "Registry::HKEY_USERS\$cj3UserSid"

        if (-not (Test-Path -LiteralPath $cj3UserHive)) {
            throw "找不到使用者登錄檔：$cj3UserSid"
        }

        $cj3RegPath = "$cj3UserHive\Software\Microsoft\IME\15.0\CHT\Cangjie"

        Write-Host "[OK] 目標目錄：$targetDir" -ForegroundColor Green
        Write-Host "[OK] 使用者 SID：$cj3UserSid" -ForegroundColor Green

        # ======================================
        # 5. 下載並驗證碼表
        # ======================================

        Write-Host ''
        Write-Host '[3/8] 下載並驗證倉頡三代碼表...' -ForegroundColor Cyan

        Invoke-WebRequest `
            -Uri $cj3Url `
            -OutFile $cj3Archive `
            -UseBasicParsing `
            -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $cj3Archive -PathType Leaf)) {
            throw '下載失敗：找不到壓縮檔。'
        }

        $cj3ArchiveInfo = Get-Item -LiteralPath $cj3Archive

        if ($cj3ArchiveInfo.Length -ne $cj3ExpectedSize) {
            throw "壓縮檔大小不符：$($cj3ArchiveInfo.Length) Bytes"
        }

        $cj3ActualHash = (
            Get-FileHash `
                -LiteralPath $cj3Archive `
                -Algorithm SHA256
        ).Hash

        if ($cj3ActualHash -ine $cj3ExpectedHash) {
            throw '壓縮檔 SHA-256 驗證失敗，已取消更新。'
        }

        Write-Host '[OK] 下載完成，SHA-256 驗證成功。' `
            -ForegroundColor Green

        # ======================================
        # 6. 解壓縮並定位碼表
        # ======================================

        Write-Host ''
        Write-Host '[4/8] 解壓縮指定碼表...' -ForegroundColor Cyan

        New-Item `
            -ItemType Directory `
            -Path $cj3ExtractDir `
            -Force `
            -ErrorAction Stop | Out-Null

        & $cj3SevenZip x $cj3Archive "-o$cj3ExtractDir" -y

        if ($LASTEXITCODE -ne 0) {
            throw "7-Zip 解壓縮失敗，ExitCode：$LASTEXITCODE"
        }

        $cj3SourceObject = Get-ChildItem `
            -LiteralPath $cj3ExtractDir `
            -Recurse `
            -File `
            -Filter 'ChtCangjieExt.lex' |
            Where-Object {
                $_.DirectoryName -like "*$prefVersion*"
            } |
            Select-Object -First 1

        if (-not $cj3SourceObject) {
            throw "找不到指定版本的碼表：$prefVersion"
        }

        if ($cj3SourceObject.Length -eq 0) {
            throw '碼表檔案大小為 0，已取消更新。'
        }

        $cj3SourceLex = $cj3SourceObject.FullName

        Write-Host "[OK] 找到碼表：$cj3SourceLex" `
            -ForegroundColor Green

        # ======================================
        # 7. 備份原始系統檔案
        # ======================================

        Write-Host ''
        Write-Host '[5/8] 備份原始系統碼表...' -ForegroundColor Cyan

        $cj3DateStr = Get-Date -Format 'yyyyMMdd_HHmmss'
        $cj3UniqueId = [guid]::NewGuid().ToString('N').Substring(0, 8)

        $cj3BackupDir = Join-Path $targetDir (
            "Backup_CJ3_${cj3DateStr}_${cj3UniqueId}"
        )

        New-Item `
            -ItemType Directory `
            -Path $cj3BackupDir `
            -ErrorAction Stop | Out-Null

        foreach ($cj3File in $cj3Files) {
            $cj3OriginalPath = Join-Path $targetDir $cj3File
            $cj3BackupPath = Join-Path $cj3BackupDir $cj3File

            $cj3Exists = Test-Path `
                -LiteralPath $cj3OriginalPath `
                -PathType Leaf

            $cj3OriginalFiles[$cj3File] = $cj3Exists

            if ($cj3Exists) {
                Copy-Item `
                    -LiteralPath $cj3OriginalPath `
                    -Destination $cj3BackupPath `
                    -Force `
                    -ErrorAction Stop

                $cj3OriginalHash = (
                    Get-FileHash `
                        -LiteralPath $cj3OriginalPath `
                        -Algorithm SHA256
                ).Hash

                $cj3BackupHash = (
                    Get-FileHash `
                        -LiteralPath $cj3BackupPath `
                        -Algorithm SHA256
                ).Hash

                if ($cj3OriginalHash -ne $cj3BackupHash) {
                    throw "備份驗證失敗：$cj3File"
                }

                Write-Host "[OK] 已備份：$cj3File" `
                    -ForegroundColor Green
            }
            else {
                Write-Host "[INFO] 原始檔案不存在：$cj3File" `
                    -ForegroundColor Yellow
            }
        }

        Write-Host "[OK] 備份目錄：$cj3BackupDir" `
            -ForegroundColor Green

        # ======================================
        # 8. 保存原始 HKSCS 設定
        # ======================================

        if (Test-Path -LiteralPath $cj3RegPath) {
            $cj3OldReg = Get-ItemProperty `
                -LiteralPath $cj3RegPath `
                -Name 'Enable HKSCS' `
                -ErrorAction SilentlyContinue

            if ($null -ne $cj3OldReg) {
                $cj3OldRegExists = $true
                $cj3OldRegValue = $cj3OldReg.'Enable HKSCS'
            }
        }

        # ======================================
        # 9. 停止輸入法程序
        # ======================================

        Write-Host ''
        Write-Host '[6/8] 停止 Microsoft IME 程序...' `
            -ForegroundColor Cyan

        Stop-Process `
            -Name 'ctfmon', 'MicrosoftIME' `
            -Force `
            -ErrorAction SilentlyContinue

        $cj3ImeStopped = $true

        Start-Sleep -Milliseconds 800

        # ======================================
        # 10. 部署新碼表
        # ======================================

        Write-Host ''
        Write-Host '[7/8] 部署倉頡三代碼表...' -ForegroundColor Cyan

        # 從此處開始若發生錯誤，
        # 必須依照備份還原原始檔案。

        $cj3DeploymentStarted = $true

        foreach ($cj3File in $cj3Files) {
            $cj3TargetFile = Join-Path $targetDir $cj3File

            if (Test-Path -LiteralPath $cj3TargetFile) {
                Remove-Item `
                    -LiteralPath $cj3TargetFile `
                    -Force `
                    -ErrorAction Stop
            }
        }

        $cj3DestinationLex = Join-Path $targetDir 'ChtCangjieExt.lex'

        Copy-Item `
            -LiteralPath $cj3SourceLex `
            -Destination $cj3DestinationLex `
            -Force `
            -ErrorAction Stop

        if (-not (Test-Path -LiteralPath $cj3DestinationLex -PathType Leaf)) {
            throw '碼表部署失敗：找不到目標檔案。'
        }

        $cj3SourceHash = (
            Get-FileHash `
                -LiteralPath $cj3SourceLex `
                -Algorithm SHA256
        ).Hash

        $cj3DestinationHash = (
            Get-FileHash `
                -LiteralPath $cj3DestinationLex `
                -Algorithm SHA256
        ).Hash

        if ($cj3SourceHash -ne $cj3DestinationHash) {
            throw '碼表部署失敗：來源與目標檔案不一致。'
        }

        Write-Host '[OK] 碼表部署與 SHA-256 比對成功。' `
            -ForegroundColor Green

        # ======================================
        # 11. 啟用 HKSCS
        # ======================================

        Write-Host ''
        Write-Host '[8/8] 啟用 HKSCS 支援...' -ForegroundColor Cyan

        $cj3RegistryTouched = $true

        if (-not (Test-Path -LiteralPath $cj3RegPath)) {
            New-Item `
                -Path $cj3RegPath `
                -Force `
                -ErrorAction Stop | Out-Null
        }

        New-ItemProperty `
            -LiteralPath $cj3RegPath `
            -Name 'Enable HKSCS' `
            -Value 1 `
            -PropertyType DWord `
            -Force `
            -ErrorAction Stop | Out-Null

        $cj3ActualRegValue = (
            Get-ItemProperty `
                -LiteralPath $cj3RegPath `
                -Name 'Enable HKSCS' `
                -ErrorAction Stop
        ).'Enable HKSCS'

        if ($cj3ActualRegValue -ne 1) {
            throw 'HKSCS 登錄檔驗證失敗。'
        }

        Write-Host '[OK] HKSCS 已啟用。' `
            -ForegroundColor Green

        # ======================================
        # 12. 更新完成
        # ======================================

        $cj3Success = $true
        $cj3DeploymentStarted = $false
        $cj3RegistryTouched = $false

        Write-Host ''
        Write-Host '==================================================' `
            -ForegroundColor Green
        Write-Host ' 倉頡三代碼表更新成功！' `
            -ForegroundColor Green
        Write-Host '==================================================' `
            -ForegroundColor Green

        Write-Host " 已安裝版本：$prefVersion"
        Write-Host " 碼表位置：$cj3DestinationLex"
        Write-Host " 原始備份：$cj3BackupDir"
        Write-Host ''
        Write-Host ' 請重新啟動 Windows，使設定完全生效。' `
            -ForegroundColor Yellow
        Write-Host '==================================================' `
            -ForegroundColor Green
    }
    catch {
        # ======================================
        # 13. 錯誤處理與還原
        # ======================================

        Write-Host ''
        Write-Host '==================================================' `
            -ForegroundColor Red
        Write-Host ' 倉頡三代更新失敗！' `
            -ForegroundColor Red
        Write-Host '==================================================' `
            -ForegroundColor Red

        Write-Host $_.Exception.Message -ForegroundColor Red

        # --------------------------------------
        # 還原原始系統碼表
        # --------------------------------------

        if ($cj3DeploymentStarted -and $cj3BackupDir) {
            Write-Host ''
            Write-Host '[*] 正在嘗試還原原始碼表...' `
                -ForegroundColor Yellow

            foreach ($cj3File in $cj3Files) {
                $cj3TargetFile = Join-Path $targetDir $cj3File
                $cj3BackupFile = Join-Path $cj3BackupDir $cj3File

                try {
                    if ($cj3OriginalFiles[$cj3File]) {
                        if (-not (Test-Path -LiteralPath $cj3BackupFile)) {
                            throw "找不到備份檔案：$cj3File"
                        }

                        Copy-Item `
                            -LiteralPath $cj3BackupFile `
                            -Destination $cj3TargetFile `
                            -Force `
                            -ErrorAction Stop

                        $cj3RestoredHash = (
                            Get-FileHash `
                                -LiteralPath $cj3TargetFile `
                                -Algorithm SHA256
                        ).Hash

                        $cj3SavedHash = (
                            Get-FileHash `
                                -LiteralPath $cj3BackupFile `
                                -Algorithm SHA256
                        ).Hash

                        if ($cj3RestoredHash -ne $cj3SavedHash) {
                            throw "還原驗證失敗：$cj3File"
                        }
                    }
                    elseif (Test-Path -LiteralPath $cj3TargetFile) {
                        Remove-Item `
                            -LiteralPath $cj3TargetFile `
                            -Force `
                            -ErrorAction Stop
                    }

                    Write-Host "[OK] 已還原：$cj3File" `
                        -ForegroundColor Green
                }
                catch {
                    Write-Host "[ERROR] 還原失敗：$cj3File" `
                        -ForegroundColor Red
                    Write-Host $_.Exception.Message -ForegroundColor Red
                }
            }
        }

        # --------------------------------------
        # 還原原始 HKSCS 設定
        # --------------------------------------

        if ($cj3RegistryTouched -and $cj3RegPath) {
            try {
                if ($cj3OldRegExists) {
                    New-ItemProperty `
                        -LiteralPath $cj3RegPath `
                        -Name 'Enable HKSCS' `
                        -Value $cj3OldRegValue `
                        -PropertyType DWord `
                        -Force `
                        -ErrorAction Stop | Out-Null
                }
                else {
                    Remove-ItemProperty `
                        -LiteralPath $cj3RegPath `
                        -Name 'Enable HKSCS' `
                        -ErrorAction SilentlyContinue
                }

                Write-Host '[OK] HKSCS 設定已還原。' `
                    -ForegroundColor Green
            }
            catch {
                Write-Host '[ERROR] HKSCS 設定還原失敗。' `
                    -ForegroundColor Red
                Write-Host $_.Exception.Message -ForegroundColor Red
            }
        }

        if ($cj3BackupDir) {
            Write-Host ''
            Write-Host "備份保留位置：$cj3BackupDir" `
                -ForegroundColor Yellow
        }

        Write-Host ''
        Write-Host '若系統檔案未完全還原，請先檢查備份。' `
            -ForegroundColor Yellow
    }
    finally {
        # ======================================
        # 14. 清理暫存目錄
        # ======================================

        if (Test-Path -LiteralPath $tempWork) {
            try {
                Remove-Item `
                    -LiteralPath $tempWork `
                    -Recurse `
                    -Force `
                    -ErrorAction Stop

                Write-Host '[INFO] 暫存檔案已清理。'
            }
            catch {
                Write-Host "[WARN] 暫存目錄清理失敗：$tempWork" `
                    -ForegroundColor Yellow
            }
        }

        # ======================================
        # 15. 重新啟動輸入法程序
        # ======================================

        if ($cj3ImeStopped) {
            try {
                Start-Process `
                    -FilePath "$env:SystemRoot\System32\ctfmon.exe" `
                    -ErrorAction Stop
            }
            catch {
                Write-Host '無法重新啟動 ctfmon.exe，請重新啟動 Windows。' `
                    -ForegroundColor Yellow
            }
        }

        $ErrorActionPreference = $cj3PreviousErrorAction
    }

    # ==========================================
    # 16. 結束 CJ3 分支
    # 避免繼續執行 CJ5 共用部署流程
    # ==========================================

    Read-Host '按 Enter 結束'

    if ($cj3Success) {
        exit 0
    }
    else {
        exit 1
    }
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

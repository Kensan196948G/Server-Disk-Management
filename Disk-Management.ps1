<#
.SYNOPSIS
    ディスク容量管理運用表 作成スクリプト

.DESCRIPTION
    対象機器ごとに以下を実施し、結果を表形式の HTML として出力する。
      1. 名前解決  (ホスト名 -> IPv4 アドレス)
      2. 導通確認  (ping)
      3. 容量取得  (WMI Win32_LogicalDisk / DCOM 優先・WinRM フォールバック)
         ※ 認証エラー (アクセス拒否・パスワード誤り・ロックアウト) の場合は、
            失敗ログオン回数の増加によるアカウントロックアウトを防ぐため
            フォールバック再試行を行わない (設定 StopOnAuthError)
      4. 運用日誌転記 (共有フォルダ上の Excel ブックへ MENU 日付と容量を転記・保存)
         -SkipExcelUpdate を指定すると省略できる (容量取得の確認だけ行いたい場合など)

    出力先     : スクリプトと同じフォルダ直下
    ファイル名 : Disk-Management-YYYYMMDDHHMM.HTML
    容量表記   : ドライブレター  空き容量 (小数点以下2桁) / 全容量  ※GB = 1024^3 バイト

.PARAMETER OutputDir
    HTML 出力先フォルダ (既定: スクリプトと同じフォルダ)

.PARAMETER Exclude
    今回の実行から除外する機器名 (Targets の Name)。保守中の機器を一時的に外す場合などに使う。
    例: .\Disk-Management.ps1 -Exclude 'MAGNET.mirai.local','GMSV0002'

.PARAMETER LogShareRoot
    運用日誌 Excel ブックが置かれている共有フォルダのルート (既定: 本番の \\gmsv0002\006情シスs\...)。
    テスト時は test フォルダ配下の疑似共有構造を指すパスに差し替える。

.PARAMETER SkipExcelUpdate
    Phase 2 (運用日誌への転記) を省略する場合に指定する。

.NOTES
    Disk-Management.bat から「管理者権限 + ExecutionPolicy Bypass」で起動される想定。
    終了コード : 0 = 全機器正常 / 2 = 導通NG・認証エラー・取得失敗・ドライブ未検出・Phase2異常あり / 1 = スクリプト異常
#>
[CmdletBinding()]
param(
    [string]$OutputDir = $PSScriptRoot,
    [string[]]$Exclude = @(),
    [string]$LogShareRoot = '\\gmsv0002\006情シスs\04.共通管理\00.システム運用日誌',
    [switch]$SkipExcelUpdate
)

$ErrorActionPreference = 'Stop'

# Get-Credential は Microsoft.PowerShell.Security モジュールのコマンドレット。
# 端末によっては、このモジュールの自動読み込みが TypeData 重複エラーで失敗することがある
# (PSModulePath に複数バージョンの PowerShell のモジュールパスが混在している環境などで発生)。
# その場合でも $PSHOME 配下のモジュールをフルパス指定すれば読み込めるため、
# 自動読み込みが失敗している場合のみフォールバックとして試みる。
if (-not (Get-Command Get-Credential -ErrorAction SilentlyContinue)) {
    $secModulePath = Join-Path $PSHOME 'Modules\Microsoft.PowerShell.Security\Microsoft.PowerShell.Security.psd1'
    if (Test-Path -LiteralPath $secModulePath) {
        try { Import-Module $secModulePath -Force -ErrorAction Stop } catch { }
    }
    if (-not (Get-Command Get-Credential -ErrorAction SilentlyContinue)) {
        Write-Host '[ERROR] Get-Credential コマンドレットを読み込めません。この端末の Microsoft.PowerShell.Security モジュールを確認してください。' -ForegroundColor Red
        exit 1
    }
}

# ============================================================
#  設定 (対象機器・認証情報はここで管理する)
# ============================================================

# 認証グループ。機器ごとに異なる資格情報を使うため、Targets の CredGroup で参照する。
# パスワードはスクリプトに保存せず、実行の都度 Get-Credential で入力する (実行端末が
# AD ドメインに参加していることが前提)。UserName は入力の手間を減らすための初期値。
#   GroupA   : 9台共通のディスク容量取得 (WMI/DCOM) に使う (VMSV3001 / GMSV0001 / 0002 / 0006 / 0008 / 0009 / 0011 / 0012)
#              GMSV0002 専用アカウント (MK0623 相当) は WMI/DCOM 権限を持たないため、容量取得には使わない。
#   GMSV0002 : GMSV0002 上の共有フォルダ (Excel 運用日誌) への接続専用。MIRAI ドメインに登録された
#              個別アカウントでのみログインする、というユーザー指定に基づき Phase 2 の SMB マッピングで使う。
#   MAGNET   : MAGNET.mirai.local 専用。過去にロックアウト事故があるため、
#              このグループの認証エラー時は絶対に別グループへの読み替えや再試行を行わない。
$CredentialGroups = [ordered]@{
    GroupA   = @{ UserName = 'MIRAI\administrator'; Prompt = 'グループA共通アカウント (VMSV3001 / GMSV0001 / 0002 / 0006 / 0008 / 0009 / 0011 / 0012 の容量取得用)' }
    GMSV0002 = @{ UserName = 'MIRAI\MK0623';         Prompt = 'GMSV0002 共有フォルダ (運用日誌 Excel) 接続専用アカウント' }
    MAGNET   = @{ UserName = '1\administrator';      Prompt = 'MAGNET.mirai.local 専用アカウント (認証エラー時は再試行しません)' }
}

$Config = @{
    PingCount        = 2      # ping 送信回数
    CimTimeoutSec    = 30     # WinRM(CIM) 操作タイムアウト (秒)
    WarnPercent      = 80     # 使用率 警告しきい値 (%)
    CritPercent      = 90     # 使用率 危険しきい値 (%)
    StopOnAuthError  = $true  # 認証エラー時に WinRM フォールバックを行わない (ロックアウト防止)
    CapacityDiffWarnGB = 5    # 運用日誌の「全容量」列と実測値の差がこれを超えたら備考に記録する。全容量は常に実測値で上書きする。
}

# Name     : 運用表に表示する機器名
# Address  : 接続先 (ホスト名または IP アドレス)
# Drives   : 取得対象ドライブレター
# CredGroup: $CredentialGroups のキー
$Targets = @(
    [pscustomobject]@{ Name = 'VMSV3001';           Address = 'VMSV3001';      Drives = @('C');           CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0001';           Address = 'GMSV0001';      Drives = @('C');           CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0002';           Address = 'GMSV0002';      Drives = @('C', 'D', 'E'); CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0006';           Address = 'GMSV0006';      Drives = @('C', 'D');      CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0008';           Address = 'GMSV0008';      Drives = @('C', 'D');      CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0009';           Address = 'GMSV0009';      Drives = @('C', 'D');      CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'GMSV0011';           Address = 'GMSV0011';      Drives = @('C');           CredGroup = 'GroupA' }
    [pscustomobject]@{ Name = 'MAGNET.mirai.local'; Address = '192.168.245.5'; Drives = @('C');           CredGroup = 'MAGNET' }
    [pscustomobject]@{ Name = 'GMSV0012';           Address = 'GMSV0012';      Drives = @('C', 'D');      CredGroup = 'GroupA' }
)

# 運用日誌 Excel「サーバー管理日誌(統合)」シート上の行対応 (固定レイアウト、現物ファイルで確認済み 2026-08-27)
# 機器の追加・削除やシート改版があった場合はここを更新すること。
$SheetRowMap = @(
    [pscustomobject]@{ Row = 18; Name = 'VMSV3001';           Drive = 'C' }
    [pscustomobject]@{ Row = 19; Name = 'GMSV0001';           Drive = 'C' }
    [pscustomobject]@{ Row = 20; Name = 'GMSV0002';           Drive = 'C' }
    [pscustomobject]@{ Row = 21; Name = 'GMSV0002';           Drive = 'D' }
    [pscustomobject]@{ Row = 22; Name = 'GMSV0002';           Drive = 'E' }
    [pscustomobject]@{ Row = 23; Name = 'GMSV0006';           Drive = 'C' }
    [pscustomobject]@{ Row = 24; Name = 'GMSV0006';           Drive = 'D' }
    [pscustomobject]@{ Row = 25; Name = 'GMSV0008';           Drive = 'C' }
    [pscustomobject]@{ Row = 26; Name = 'GMSV0008';           Drive = 'D' }
    [pscustomobject]@{ Row = 27; Name = 'GMSV0009';           Drive = 'C' }
    [pscustomobject]@{ Row = 28; Name = 'GMSV0009';           Drive = 'D' }
    [pscustomobject]@{ Row = 29; Name = 'GMSV0011';           Drive = 'C' }
    [pscustomobject]@{ Row = 30; Name = 'MAGNET.mirai.local'; Drive = 'C' }
    [pscustomobject]@{ Row = 31; Name = 'GMSV0012';           Drive = 'C' }
    [pscustomobject]@{ Row = 32; Name = 'GMSV0012';           Drive = 'D' }
)
$SheetColDrive = 5   # E列: ドライブレター
$SheetColTotal = 6   # F列: 容量(全容量) … 確認のみ、上書きしない
$SheetColFree  = 7   # G列: 空き(空き容量) … 最新値で上書きする

# ============================================================
#  関数
# ============================================================

# ホスト名 / IP 文字列から IPv4 アドレスを 1 件返す (解決不可なら $null)
function Resolve-TargetIPv4 {
    param([string]$Address)
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($Address) |
              Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
              Select-Object -First 1
        if ($ip) { return $ip.IPAddressToString }
    }
    catch { }
    return $null
}

# ping 導通確認。1 回でも応答があれば OK とする
function Test-TargetPing {
    param([string]$Address, [int]$Count)
    $replies = @(Test-Connection -ComputerName $Address -Count $Count -ErrorAction SilentlyContinue)
    if ($replies.Count -gt 0) {
        return [pscustomobject]@{ Success = $true; ResponseTime = ($replies | Select-Object -Last 1).ResponseTime }
    }
    return [pscustomobject]@{ Success = $false; ResponseTime = $null }
}

# 認証系エラー (アクセス拒否 / パスワード誤り / ロックアウト) かを判定する。
# 認証系エラーの場合に別経路で再試行すると失敗ログオン回数が増え、
# アカウントロックアウトを誘発するため、フォールバック打ち切りの判断に使う。
function Test-AuthError {
    param([System.Exception]$Exception)
    # 0x80070005 E_ACCESSDENIED / 0x8007052E ERROR_LOGON_FAILURE / 0x80070775 ERROR_ACCOUNT_LOCKED_OUT
    $authHResults = @(-2147024891, -2147023570, -2147023499)
    $e = $Exception
    while ($null -ne $e) {
        if ($authHResults -contains $e.HResult) { return $true }
        if ($e.Message -match 'アクセスが拒否|Access is denied|ロックアウト|locked out|パスワードが(正しく|間違って)|password is incorrect|ログオンに失敗|Logon failure') { return $true }
        $e = $e.InnerException
    }
    return $false
}

# Win32_LogicalDisk (ローカルディスクのみ) を取得。WMI(DCOM) -> WinRM の順に試行
function Get-TargetLogicalDisk {
    param([string]$Address, [pscredential]$Credential, [int]$TimeoutSec, [bool]$StopOnAuthError)
    $errors    = @()
    $authError = $false

    # --- 1. WMI (DCOM) : Get-WmiObject ------------------------------------------
    #   New-CimSession -Protocol Dcom は DCOM 強化 (KB5004442) 適用済みサーバーで
    #   アクセス拒否となる場合があるため、認証レベルを明示できる Get-WmiObject を使う
    try {
        $disks = @(Get-WmiObject -ComputerName $Address -Credential $Credential -Class Win32_LogicalDisk -Filter 'DriveType = 3' `
                                 -Authentication PacketPrivacy -Impersonation Impersonate -ErrorAction Stop |
                   Select-Object DeviceID, Size, FreeSpace, VolumeName)
        return [pscustomobject]@{ Success = $true; Method = 'WMI(DCOM)'; Disks = $disks; Error = $null; AuthError = $false }
    }
    catch {
        $errors   += ('WMI(DCOM): {0}' -f $_.Exception.Message.Trim())
        $authError = Test-AuthError -Exception $_.Exception
    }

    # --- 2. WinRM (CIM / WSMan) : フォールバック --------------------------------
    if ($authError -and $StopOnAuthError) {
        $errors += 'WinRM: 認証エラーのため再試行を中止 (ロックアウト防止)'
    }
    else {
        $session = $null
        try {
            $option  = New-CimSessionOption -Protocol Wsman
            $session = New-CimSession -ComputerName $Address -Credential $Credential -SessionOption $option `
                                      -OperationTimeoutSec $TimeoutSec -ErrorAction Stop
            $disks   = @(Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' `
                                         -OperationTimeoutSec $TimeoutSec -ErrorAction Stop |
                         Select-Object DeviceID, Size, FreeSpace, VolumeName)
            return [pscustomobject]@{ Success = $true; Method = 'WinRM'; Disks = $disks; Error = $null; AuthError = $false }
        }
        catch {
            $errors += ('WinRM: {0}' -f $_.Exception.Message.Trim())
            if (-not $authError) { $authError = Test-AuthError -Exception $_.Exception }
        }
        finally {
            if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
        }
    }
    return [pscustomobject]@{ Success = $false; Method = $null; Disks = @(); Error = ($errors -join ' / '); AuthError = $authError }
}

# COM の Range オブジェクトへ値を設定する。
# PowerShell 関数スコープ内で $range.Value2 = $value という直接代入を行うと、
# COM Interop の VARIANT 変換で InvalidCastException になることがある既知の癖があるため、
# GetType().InvokeMember() 経由でセッターを明示的に呼び出す、より確実な経路を使う。
function Set-ComCellValue {
    param($Range, $Value)
    [void]$Range.GetType().InvokeMember('Value2', [System.Reflection.BindingFlags]::SetProperty, $null, $Range, @($Value))
}

# 日付から日本の会計年度 (4月始まり) を返す。例: 2026/8 -> 2026, 2026/2 -> 2025
function Get-JapaneseFiscalYear {
    param([datetime]$Date)
    if ($Date.Month -ge 4) { return $Date.Year }
    return $Date.Year - 1
}

# 運用日誌 Excel ブックへ MENU 日付と各機器の容量を転記して保存する (Phase 2)。
# 戻り値: 各手順の PASS/FAIL/BLOCKED/NOT RUN と詳細メモを持つ順序付きハッシュテーブル。
function Invoke-OperationLogUpdate {
    param(
        [string]$ShareRoot,
        [pscredential]$MappingCredential,
        [datetime]$TodayDate,
        [System.Collections.Generic.List[object]]$Results,
        [System.Collections.Generic.List[object]]$SheetRowMap,
        [int]$ColDrive,
        [int]$ColTotal,
        [int]$ColFree,
        [double]$CapacityDiffWarnGB
    )

    $phase2 = [ordered]@{
        FolderResolved   = 'NOT RUN'
        PrevFileFound    = 'NOT RUN'
        TodayFileReady   = 'NOT RUN'
        MenuUpdated      = 'NOT RUN'
        CapacityUpdated  = 'NOT RUN'
        Saved            = 'NOT RUN'
        Notes            = @()
    }

    $prevDate    = $TodayDate.AddDays(-1)
    $fiscalYear  = Get-JapaneseFiscalYear -Date $TodayDate
    $monthFolder = '{0:yyyyMM}' -f $TodayDate
    $folder      = Join-Path $ShareRoot ('{0}年度\{1}' -f $fiscalYear, $monthFolder)
    $prevFileName  = 'システム運用日誌_{0:yyyyMMdd}.xlsm' -f $prevDate
    $todayFileName = 'システム運用日誌_{0:yyyyMMdd}.xlsm' -f $TodayDate
    $prevFile  = Join-Path $folder $prevFileName
    $todayFile = Join-Path $folder $todayFileName

    $mapped = $false
    $excel  = $null
    $wb     = $null
    try {
        # UNC共有への認証。New-PSDrive はこのプロセス内でしか有効でなく Excel COM (別プロセス) から
        # 見えないため、システムレベルでマッピングされる New-SmbMapping を使う。
        if ($ShareRoot -like '\\*' -and $MappingCredential) {
            Get-SmbMapping -RemotePath $ShareRoot -ErrorAction SilentlyContinue | Remove-SmbMapping -Force -ErrorAction SilentlyContinue
            New-SmbMapping -RemotePath $ShareRoot -Credential $MappingCredential -Persistent $false -ErrorAction Stop | Out-Null
            $mapped = $true
        }

        if (-not (Test-Path -LiteralPath $folder)) {
            $phase2.FolderResolved = 'FAIL'
            $phase2.Notes += "対象フォルダが見つかりません: $folder"
            return $phase2
        }
        $phase2.FolderResolved = 'PASS'

        if (-not (Test-Path -LiteralPath $prevFile)) {
            $phase2.PrevFileFound = 'BLOCKED'
            $phase2.Notes += "前日ファイルが見つかりません: $prevFile"
            return $phase2
        }
        $phase2.PrevFileFound = 'PASS'

        if (Test-Path -LiteralPath $todayFile) {
            $phase2.TodayFileReady = 'PASS'
            $phase2.Notes += '当日ファイルは既に存在するため、それを更新する (前日ファイルからのコピーはスキップ)'
        }
        else {
            Copy-Item -LiteralPath $prevFile -Destination $todayFile -ErrorAction Stop
            $phase2.TodayFileReady = 'PASS'
            $phase2.Notes += '前日ファイルをコピーして当日ファイルを作成した'
        }

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
        try { $excel.AutomationSecurity = 3 } catch { }
        $missing = [Type]::Missing
        # UpdateLinks=0, ReadOnly=$false, IgnoreReadOnly=$true, Notify=$false
        $wb = $excel.Workbooks.Open($todayFile, 0, $false, $missing, $missing, $missing, $true, $missing, $missing, $missing, $false, $missing, $false)

        # --- MENU シート: D10(前日)/E10(当日) ---
        $menu = $wb.Worksheets.Item('MENU')
        $d10 = $menu.Cells.Item(10, 4)
        $e10 = $menu.Cells.Item(10, 5)
        $expectedPrevText  = '{0}月{1}日' -f $prevDate.Month, $prevDate.Day
        $expectedTodayText = '{0}月{1}日' -f $TodayDate.Month, $TodayDate.Day

        if ($d10.Text -ne $expectedPrevText) {
            $beforeText = $d10.Text
            Set-ComCellValue -Range $d10 -Value (Get-Date -Year $TodayDate.Year -Month $prevDate.Month -Day $prevDate.Day)
            $phase2.Notes += ("MENU!D10 を '{0}' -> '{1}' に修正した" -f $beforeText, $d10.Text)
        }
        if ($e10.Text -ne $expectedTodayText) {
            $phase2.Notes += ("MENU!E10 が想定値と異なる (現在値='{0}' 想定値='{1}')。手動確認が必要。" -f $e10.Text, $expectedTodayText)
        }
        $phase2.MenuUpdated = if ($d10.Text -eq $expectedPrevText) { 'PASS' } else { 'FAIL' }

        # --- サーバー管理日誌(統合) シート: 容量転記 ---
        $sheet = $wb.Worksheets.Item('サーバー管理日誌(統合)')
        $capacityOk = $true
        foreach ($rowMap in $SheetRowMap) {
            $result = $Results | Where-Object { $_.Name -eq $rowMap.Name } | Select-Object -First 1
            if (-not $result) { continue }
            $driveRow = $result.DriveRows | Where-Object { $_.Drive -eq "$($rowMap.Drive):" } | Select-Object -First 1
            if (-not $driveRow -or $null -eq $driveRow.TotalGB) {
                $phase2.Notes += ("{0} {1}: 容量未取得のためシート更新をスキップ" -f $rowMap.Name, $rowMap.Drive)
                continue
            }

            $cellDrive = $sheet.Cells.Item($rowMap.Row, $ColDrive)
            if ($cellDrive.Text -ne $rowMap.Drive) {
                $capacityOk = $false
                $phase2.Notes += ("行{0}: ドライブ列が想定外 (シート='{1}' 想定='{2}')。この行の更新をスキップ" -f $rowMap.Row, $cellDrive.Text, $rowMap.Drive)
                continue
            }

            $cellTotal = $sheet.Cells.Item($rowMap.Row, $ColTotal)
            $existingTotal = 0.0
            [double]::TryParse([string]$cellTotal.Value2, [ref]$existingTotal) | Out-Null
            $diff = [math]::Abs($existingTotal - $driveRow.TotalGB)
            if ($diff -gt $CapacityDiffWarnGB) {
                $phase2.Notes += ("{0} {1}: 全容量差異あり (シート記載={2:N2}GB 実測={3:N2}GB, 差={4:N2}GB)。実測値で更新しました。" -f `
                    $rowMap.Name, $rowMap.Drive, $existingTotal, $driveRow.TotalGB, $diff)
            }
            $totalVal = [double]([math]::Round([double]$driveRow.TotalGB, 2))
            Set-ComCellValue -Range $cellTotal -Value $totalVal

            $cellFree = $sheet.Cells.Item($rowMap.Row, $ColFree)
            $freeVal = [double]([math]::Round([double]$driveRow.FreeGB, 2))
            Set-ComCellValue -Range $cellFree -Value $freeVal
        }
        $phase2.CapacityUpdated = if ($capacityOk) { 'PASS' } else { 'FAIL' }

        $wb.Save()
        $phase2.Saved = 'PASS'
    }
    catch {
        $exType = $_.Exception.GetType().FullName
        $inner = if ($_.Exception.InnerException) { ' | Inner: ' + $_.Exception.InnerException.Message } else { '' }
        $phase2.Notes += ('ERROR (line {0}, {1}): {2}{3}' -f $_.InvocationInfo.ScriptLineNumber, $exType, $_.Exception.Message, $inner)
        foreach ($key in @('FolderResolved','PrevFileFound','TodayFileReady','MenuUpdated','CapacityUpdated','Saved')) {
            if ($phase2[$key] -eq 'NOT RUN') { $phase2[$key] = 'FAIL' }
        }
    }
    finally {
        if ($wb) { try { $wb.Close($false) } catch { } }
        if ($excel) { try { $excel.Quit() } catch { }; [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null }
        if ($mapped) { Get-SmbMapping -RemotePath $ShareRoot -ErrorAction SilentlyContinue | Remove-SmbMapping -Force -ErrorAction SilentlyContinue }
    }

    return $phase2
}

# バイト数 -> "1,234.56" (GB, 小数点以下 2 桁)
function Format-GB {
    param($Bytes)
    if ($null -eq $Bytes) { return '-' }
    return ('{0:#,##0.00}' -f ([double]$Bytes / 1GB))
}

function ConvertTo-HtmlText {
    param([AllowNull()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($Text)
}

# ============================================================
#  メイン処理
# ============================================================
try {
    $startTime = Get-Date
    if ([string]::IsNullOrWhiteSpace($OutputDir)) { $OutputDir = (Get-Location).Path }
    $OutputDir = [System.IO.Path]::GetFullPath($OutputDir)   # "..\." などを正規化
    if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }

    $activeTargets = @($Targets | Where-Object { $Exclude -notcontains $_.Name })

    # 今回の実行で実際に必要な認証グループだけ、その場でログイン認証 (Get-Credential) を行う。
    # パスワードはスクリプトに保存しない。対象を絞ることで、除外した機器やスキップした
    # Phase 2 の分の入力を運用担当者に求めずに済む。
    $neededGroups = New-Object System.Collections.Generic.List[string]
    foreach ($t in $activeTargets) {
        if (-not $neededGroups.Contains($t.CredGroup)) { $neededGroups.Add($t.CredGroup) }
    }
    if (-not $SkipExcelUpdate -and -not $neededGroups.Contains('GMSV0002')) {
        $neededGroups.Add('GMSV0002')
    }

    Write-Host '------------------------------------------------------------'
    Write-Host '  ログイン認証 (この端末が AD ドメインに参加している必要があります)'
    $credentialsByGroup = @{}
    foreach ($groupName in $CredentialGroups.Keys) {
        if (-not $neededGroups.Contains($groupName)) { continue }
        $g = $CredentialGroups[$groupName]
        Write-Host ('    [{0}] {1}' -f $groupName, $g.Prompt) -ForegroundColor Cyan
        $cred = Get-Credential -UserName $g.UserName -Message $g.Prompt
        if (-not $cred) {
            throw ('認証情報が入力されませんでした ({0})。処理を中止します。' -f $g.Prompt)
        }
        $credentialsByGroup[$groupName] = $cred
    }
    Write-Host '------------------------------------------------------------'

    Write-Host '============================================================'
    Write-Host ' ディスク容量管理運用表 作成'
    Write-Host ('  開始: {0:yyyy/MM/dd HH:mm:ss}   対象: {1} 台' -f $startTime, $activeTargets.Count)
    if ($Exclude.Count -gt 0) { Write-Host ('  除外: {0}' -f ($Exclude -join ', ')) -ForegroundColor DarkYellow }
    Write-Host '============================================================'

    $results = New-Object System.Collections.Generic.List[object]
    $no = 0

    foreach ($target in $activeTargets) {
        $no++
        Write-Host ('[{0}/{1}] {2,-20} ' -f $no, $activeTargets.Count, $target.Name) -NoNewline

        # --- 1. 名前解決 ---
        $ip = Resolve-TargetIPv4 -Address $target.Address

        # --- 2. 導通確認 ---
        $ping = Test-TargetPing -Address $target.Address -Count $Config.PingCount
        if ($ping.Success) {
            Write-Host ('ping OK ({0}ms) ' -f $ping.ResponseTime) -ForegroundColor Green -NoNewline
        }
        else {
            Write-Host 'ping NG ' -ForegroundColor Red -NoNewline
        }

        # --- 3. 容量取得 (ping NG でも WMI 到達可能な場合があるため必ず試行) ---
        $targetCredential = $credentialsByGroup[$target.CredGroup]
        $query = Get-TargetLogicalDisk -Address $target.Address -Credential $targetCredential `
                                       -TimeoutSec $Config.CimTimeoutSec -StopOnAuthError $Config.StopOnAuthError
        if ($query.Success) {
            Write-Host ('容量取得 OK [{0}]' -f $query.Method) -ForegroundColor Green
        }
        else {
            $ngLabel = if ($query.AuthError) { '容量取得 NG (認証エラー)' } else { '容量取得 NG' }
            Write-Host $ngLabel -ForegroundColor Red
            Write-Host ('      {0}' -f $query.Error) -ForegroundColor DarkYellow
        }

        # --- ドライブ別行の生成 ---
        $driveRows = @()
        foreach ($drive in $target.Drives) {
            $letter = ($drive.ToString().TrimEnd(':')).ToUpper()
            $row = [pscustomobject]@{
                Drive   = "${letter}:"
                FreeGB  = $null
                TotalGB = $null
                UsedPct = $null
                Status  = ''
                Note    = ''
            }

            if (-not $query.Success) {
                $row.Status = if ($query.AuthError) { '認証エラー' } elseif ($ping.Success) { '取得失敗' } else { '導通NG' }
                $row.Note   = $query.Error
            }
            else {
                $disk = $query.Disks | Where-Object { $_.DeviceID -eq "${letter}:" } | Select-Object -First 1
                if ($null -eq $disk -or $null -eq $disk.Size -or [double]$disk.Size -le 0) {
                    $row.Status = 'ドライブ未検出'
                    $row.Note   = ('対象機器に {0}: (ローカルディスク) が存在しません' -f $letter)
                }
                else {
                    $row.FreeGB  = [double]$disk.FreeSpace / 1GB
                    $row.TotalGB = [double]$disk.Size / 1GB
                    $row.UsedPct = [math]::Round((1 - ([double]$disk.FreeSpace / [double]$disk.Size)) * 100, 1)
                    $row.Status  = if ($row.UsedPct -ge $Config.CritPercent) { '危険' }
                                   elseif ($row.UsedPct -ge $Config.WarnPercent) { '警告' }
                                   else { '正常' }
                    $row.Note    = if ($disk.VolumeName) { '{0} / {1}' -f $query.Method, $disk.VolumeName } else { $query.Method }
                }
                Write-Host ('      {0}  空き {1,12} GB / 全体 {2,12} GB  ({3})' -f $row.Drive, (Format-GB $disk.FreeSpace), (Format-GB $disk.Size), $row.Status)
            }
            $driveRows += $row
        }

        $results.Add([pscustomobject]@{
            No        = $no
            Name      = $target.Name
            Address   = $target.Address
            IP        = $ip
            Ping      = $ping
            Query     = $query
            DriveRows = $driveRows
        })
    }

    # ============================================================
    #  Phase 2: 運用日誌 (共有 Excel ブック) への転記
    # ============================================================
    Write-Host '------------------------------------------------------------'
    if ($SkipExcelUpdate) {
        Write-Host '  Phase 2 (運用日誌転記): -SkipExcelUpdate 指定によりスキップ' -ForegroundColor DarkYellow
        $phase2Result = [ordered]@{
            FolderResolved = 'NOT RUN'; PrevFileFound = 'NOT RUN'; TodayFileReady = 'NOT RUN'
            MenuUpdated = 'NOT RUN'; CapacityUpdated = 'NOT RUN'; Saved = 'NOT RUN'; Notes = @('-SkipExcelUpdate 指定によりスキップ')
        }
    }
    else {
        Write-Host '  Phase 2 (運用日誌転記) 実行中...'
        $phase2Result = Invoke-OperationLogUpdate -ShareRoot $LogShareRoot `
            -MappingCredential $credentialsByGroup['GMSV0002'] `
            -TodayDate $startTime.Date -Results $results -SheetRowMap $SheetRowMap `
            -ColDrive $SheetColDrive -ColTotal $SheetColTotal -ColFree $SheetColFree `
            -CapacityDiffWarnGB $Config.CapacityDiffWarnGB
        foreach ($key in @('FolderResolved','PrevFileFound','TodayFileReady','MenuUpdated','CapacityUpdated','Saved')) {
            $color = switch ($phase2Result[$key]) { 'PASS' { 'Green' } 'BLOCKED' { 'DarkYellow' } 'FAIL' { 'Red' } default { 'Gray' } }
            Write-Host ('    {0,-16}: {1}' -f $key, $phase2Result[$key]) -ForegroundColor $color
        }
        foreach ($note in $phase2Result.Notes) { Write-Host ('    - {0}' -f $note) -ForegroundColor DarkYellow }
    }
    Write-Host '------------------------------------------------------------'

    # ============================================================
    #  HTML 生成
    # ============================================================
    $endTime      = Get-Date
    $pingOkCount  = @($results | Where-Object { $_.Ping.Success }).Count
    $queryOkCount = @($results | Where-Object { $_.Query.Success }).Count
    $allRows      = @($results | ForEach-Object { $_.DriveRows })
    $ngRowCount   = @($allRows | Where-Object { $_.Status -in @('導通NG', '認証エラー', '取得失敗', 'ドライブ未検出') }).Count
    $warnRowCount = @($allRows | Where-Object { $_.Status -in @('警告', '危険') }).Count

    $css = @'
    body { font-family: "Meiryo UI", Meiryo, "Yu Gothic UI", "Segoe UI", sans-serif; font-size: 13px; color: #222; margin: 24px; background: #fff; }
    h1   { font-size: 20px; margin: 0 0 6px 0; }
    .meta { color: #555; margin-bottom: 14px; line-height: 1.7; }
    .meta span { display: inline-block; margin-right: 18px; }
    .summary { margin-bottom: 14px; }
    .summary span { display: inline-block; padding: 3px 10px; margin-right: 6px; border-radius: 3px; background: #eef2f7; border: 1px solid #cfd8e3; }
    .summary .ng   { background: #fde8e8; border-color: #f2b8b8; color: #a00; }
    .summary .warn { background: #fff5db; border-color: #f0d48a; color: #7a5a00; }
    table { border-collapse: collapse; width: 100%; max-width: 1400px; table-layout: fixed; }
    th, td { border: 1px solid #b8c2cc; padding: 6px 10px; vertical-align: middle; }
    th { background: #2f5b8f; color: #fff; font-weight: bold; text-align: center; }
    td.num  { text-align: right; font-family: Consolas, "Courier New", monospace; white-space: nowrap; }
    td.center { text-align: center; }
    td.name { font-weight: bold; overflow-wrap: anywhere; }
    td.note { font-size: 11px; color: #555; overflow-wrap: anywhere; line-height: 1.5; }
    tr.server-first td { border-top: 2px solid #6d7f93; }
    .ok   { color: #0a7a2f; font-weight: bold; }
    .ng   { color: #c00000; font-weight: bold; }
    .st-ok   { background: #e8f6ec; }
    .st-warn { background: #fff5db; }
    .st-crit { background: #fde8e8; }
    .st-ng   { background: #eeeeee; color: #888; }
    .bar { display: inline-block; width: 110px; height: 11px; background: #e4e8ec; border: 1px solid #b8c2cc; vertical-align: middle; margin-right: 6px; }
    .bar .fill { height: 100%; background: #4a90d9; }
    .bar .fill.warn { background: #f0b429; }
    .bar .fill.crit { background: #d9534f; }
    .legend { margin-top: 12px; color: #555; font-size: 11px; line-height: 1.7; }
    .phase2 { margin: 4px 0 18px; padding: 10px 14px; border: 1px solid #cfd8e3; border-radius: 4px; background: #f7f9fb; }
    .phase2 h2 { font-size: 14px; margin: 0 0 8px; }
    .phase2 .steps span { display: inline-block; padding: 2px 9px; margin: 0 6px 6px 0; border-radius: 3px; font-size: 12px; font-weight: bold; }
    .phase2 .step-pass    { background: #e8f6ec; color: #0a7a2f; border: 1px solid #b9e0c4; }
    .phase2 .step-fail    { background: #fde8e8; color: #c00000; border: 1px solid #f2b8b8; }
    .phase2 .step-blocked { background: #fff5db; color: #7a5a00; border: 1px solid #f0d48a; }
    .phase2 .step-notrun  { background: #eeeeee; color: #888; border: 1px solid #ddd; }
    .phase2 ul { margin: 6px 0 0; padding-left: 20px; font-size: 12px; color: #555; }
    @media print { body { margin: 8mm; } .bar { border-color: #999; } }
'@

    $title = 'ディスク容量管理運用表 {0:yyyy/MM/dd HH:mm}' -f $startTime
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<!DOCTYPE html>')
    [void]$sb.AppendLine('<html lang="ja">')
    [void]$sb.AppendLine('<head>')
    [void]$sb.AppendLine('<meta charset="UTF-8">')
    [void]$sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
    [void]$sb.AppendLine(('<title>{0}</title>' -f (ConvertTo-HtmlText $title)))
    [void]$sb.AppendLine('<style>')
    [void]$sb.AppendLine($css)
    [void]$sb.AppendLine('</style>')
    [void]$sb.AppendLine('</head>')
    [void]$sb.AppendLine('<body>')
    [void]$sb.AppendLine('<h1>ディスク容量管理運用表</h1>')
    [void]$sb.AppendLine('<div class="meta">')
    [void]$sb.AppendLine(('<span>作成日時: {0:yyyy/MM/dd HH:mm:ss}</span>' -f $startTime))
    [void]$sb.AppendLine(('<span>所要時間: {0:0} 秒</span>' -f ($endTime - $startTime).TotalSeconds))
    [void]$sb.AppendLine(('<span>実行ホスト: {0}</span>' -f (ConvertTo-HtmlText $env:COMPUTERNAME)))
    [void]$sb.AppendLine(('<span>実行ユーザー: {0}</span>' -f (ConvertTo-HtmlText ('{0}\{1}' -f $env:USERDOMAIN, $env:USERNAME))))
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('<div class="summary">')
    [void]$sb.AppendLine(('<span>対象機器: {0} 台</span>' -f $results.Count))
    [void]$sb.AppendLine(('<span>導通OK: {0} / {1} 台</span>' -f $pingOkCount, $results.Count))
    [void]$sb.AppendLine(('<span>容量取得OK: {0} / {1} 台</span>' -f $queryOkCount, $results.Count))
    [void]$sb.AppendLine(('<span class="{0}">異常 (導通NG/認証エラー/取得失敗/未検出): {1} 件</span>' -f $(if ($ngRowCount -gt 0) { 'ng' } else { '' }), $ngRowCount))
    if ($Exclude.Count -gt 0) {
        [void]$sb.AppendLine(('<span class="warn">除外: {0}</span>' -f (ConvertTo-HtmlText ($Exclude -join ', '))))
    }
    [void]$sb.AppendLine(('<span class="{0}">使用率 警告/危険: {1} 件</span>' -f $(if ($warnRowCount -gt 0) { 'warn' } else { '' }), $warnRowCount))
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<div class="phase2">')
    [void]$sb.AppendLine('<h2>運用日誌 (Excel) 転記結果</h2>')
    [void]$sb.AppendLine('<div class="steps">')
    $phase2Labels = [ordered]@{
        FolderResolved  = 'フォルダ特定'
        PrevFileFound   = '前日ファイル確認'
        TodayFileReady  = '当日ファイル準備'
        MenuUpdated     = 'MENU日付更新'
        CapacityUpdated = '容量転記'
        Saved           = '保存'
    }
    foreach ($key in $phase2Labels.Keys) {
        $val = $phase2Result[$key]
        $stepClass = switch ($val) { 'PASS' { 'step-pass' } 'FAIL' { 'step-fail' } 'BLOCKED' { 'step-blocked' } default { 'step-notrun' } }
        [void]$sb.AppendLine(('<span class="{0}">{1}: {2}</span>' -f $stepClass, (ConvertTo-HtmlText $phase2Labels[$key]), (ConvertTo-HtmlText $val)))
    }
    [void]$sb.AppendLine('</div>')
    if ($phase2Result.Notes.Count -gt 0) {
        [void]$sb.AppendLine('<ul>')
        foreach ($note in $phase2Result.Notes) {
            [void]$sb.AppendLine(('<li>{0}</li>' -f (ConvertTo-HtmlText $note)))
        }
        [void]$sb.AppendLine('</ul>')
    }
    [void]$sb.AppendLine('</div>')

    [void]$sb.AppendLine('<table>')
    [void]$sb.AppendLine('<colgroup><col style="width:4%"><col style="width:14%"><col style="width:12%"><col style="width:9%"><col style="width:6%"><col style="width:17%"><col style="width:14%"><col style="width:7%"><col></colgroup>')
    [void]$sb.AppendLine('<thead><tr>')
    [void]$sb.AppendLine('<th>No.</th><th>機器名</th><th>IPアドレス</th><th>導通 (ping)</th><th>ドライブ</th><th>ディスク容量<br>空き容量 / 全容量 (GB)</th><th>使用率</th><th>状態</th><th>備考</th>')
    [void]$sb.AppendLine('</tr></thead>')
    [void]$sb.AppendLine('<tbody>')

    foreach ($r in $results) {
        $span = $r.DriveRows.Count
        if (-not $r.IP) {
            $ipText = '<span class="ng">名前解決失敗</span>'
        }
        elseif ($r.IP -ne $r.Address) {
            $ipText = '{0}<br><small>({1})</small>' -f (ConvertTo-HtmlText $r.IP), (ConvertTo-HtmlText $r.Address)
        }
        else {
            $ipText = ConvertTo-HtmlText $r.IP
        }
        $pingHtml = if ($r.Ping.Success) { '<span class="ok">OK</span> ({0} ms)' -f $r.Ping.ResponseTime } else { '<span class="ng">NG</span>' }

        $first = $true
        foreach ($d in $r.DriveRows) {
            $rowClass = if ($first) { ' class="server-first"' } else { '' }
            [void]$sb.Append(('<tr{0}>' -f $rowClass))
            if ($first) {
                [void]$sb.Append(('<td class="center" rowspan="{0}">{1}</td>' -f $span, $r.No))
                [void]$sb.Append(('<td class="name" rowspan="{0}">{1}</td>' -f $span, (ConvertTo-HtmlText $r.Name)))
                [void]$sb.Append(('<td class="center" rowspan="{0}">{1}</td>' -f $span, $ipText))
                [void]$sb.Append(('<td class="center" rowspan="{0}">{1}</td>' -f $span, $pingHtml))
                $first = $false
            }

            switch ($d.Status) {
                '正常'  { $stClass = 'st-ok';   $fillClass = '' }
                '警告'  { $stClass = 'st-warn'; $fillClass = ' warn' }
                '危険'  { $stClass = 'st-crit'; $fillClass = ' crit' }
                default { $stClass = 'st-ng';   $fillClass = '' }
            }

            [void]$sb.Append(('<td class="center"><b>{0}</b></td>' -f (ConvertTo-HtmlText $d.Drive)))
            if ($null -ne $d.TotalGB) {
                $freeText  = '{0:#,##0.00}' -f $d.FreeGB
                $totalText = '{0:#,##0.00}' -f $d.TotalGB
                $barWidth  = [math]::Min(100, [math]::Max(0, $d.UsedPct))
                [void]$sb.Append(('<td class="num">{0} GB / {1} GB</td>' -f $freeText, $totalText))
                [void]$sb.Append(('<td class="num"><span class="bar"><span class="fill{0}" style="display:block;width:{1}%"></span></span>{2:0.0}%</td>' -f $fillClass, $barWidth, $d.UsedPct))
            }
            else {
                [void]$sb.Append('<td class="center">-</td><td class="center">-</td>')
            }
            [void]$sb.Append(('<td class="center {0}">{1}</td>' -f $stClass, (ConvertTo-HtmlText $d.Status)))
            # 備考が長い場合 (エラー文など) は 200 文字で省略し、全文は title 属性 (マウスオーバー) で参照可能にする
            $noteFull = [string]$d.Note
            $noteText = if ($noteFull.Length -gt 200) { $noteFull.Substring(0, 200) + '…' } else { $noteFull }
            $noteAttr = if ($noteFull.Length -gt 200) { ' title="{0}"' -f (ConvertTo-HtmlText $noteFull) } else { '' }
            [void]$sb.Append(('<td class="note"{0}>{1}</td>' -f $noteAttr, (ConvertTo-HtmlText $noteText)))
            [void]$sb.AppendLine('</tr>')
        }
    }

    [void]$sb.AppendLine('</tbody>')
    [void]$sb.AppendLine('</table>')
    [void]$sb.AppendLine('<div class="legend">')
    [void]$sb.AppendLine(('※ 容量は GB (1 GB = 1,024^3 バイト、エクスプローラー表示と同じ基準)、小数点以下 2 桁。使用率 {0}% 以上を「警告」、{1}% 以上を「危険」と表示。<br>' -f $Config.WarnPercent, $Config.CritPercent))
    [void]$sb.AppendLine('※ 導通は ping、容量は WMI (Win32_LogicalDisk, DCOM 経由・不可の場合 WinRM) で取得。ping NG でも WMI で取得できた場合は容量を表示する。<br>')
    [void]$sb.AppendLine('※ 「認証エラー」はアクセス拒否・パスワード誤り・アカウントロックアウトのいずれか。ロックアウト防止のため WinRM への再試行は行わない。対象機器のアカウント状態とパスワードを確認すること。')
    [void]$sb.AppendLine('</div>')
    [void]$sb.AppendLine('</body>')
    [void]$sb.AppendLine('</html>')

    $fileName   = 'Disk-Management-{0:yyyyMMddHHmm}.HTML' -f $startTime
    $outputPath = Join-Path -Path $OutputDir -ChildPath $fileName
    [System.IO.File]::WriteAllText($outputPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding $true))

    Write-Host '------------------------------------------------------------'
    Write-Host ('  対象 {0} 台 / 導通OK {1} 台 / 容量取得OK {2} 台 / 異常 {3} 件 / 警告・危険 {4} 件' -f $results.Count, $pingOkCount, $queryOkCount, $ngRowCount, $warnRowCount)
    Write-Host ('  出力: {0}' -f $outputPath) -ForegroundColor Cyan
    Write-Host '------------------------------------------------------------'

    $phase2StepKeys = @('FolderResolved','PrevFileFound','TodayFileReady','MenuUpdated','CapacityUpdated','Saved')
    $phase2HasFail = @($phase2StepKeys | ForEach-Object { $phase2Result[$_] } | Where-Object { $_ -in @('FAIL', 'BLOCKED') }).Count -gt 0
    if ($ngRowCount -gt 0 -or $phase2HasFail) { exit 2 }
    exit 0
}
catch {
    Write-Host ('[ERROR] {0}' -f $_.Exception.Message) -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
    exit 1
}

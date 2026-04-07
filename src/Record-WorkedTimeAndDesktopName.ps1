<#
  .synopsis

  現在の時刻とデスクトップ名を取得し、タブ区切りテキストとしてファイルへ書き込む。

  .description

  書き込んだタブ区切りテキストを Excel に貼り付けると、以下のような表を作成できる。

    +----------------+----------+----------+----------+----------------------------+----------+
    |     年月日     | 開始時刻 | 終了時刻 | 作業時間 | 作業種別（デスクトップ名） |   説明   |
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(土) |  09:00   |  12:00   |   3.0 h  | Project A                  | <Start> ｜
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(土) |  13:00   |  15:00   |   2.0 h  | 教育                       |          |
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(土) |  15:00   |  17:30   |   2.5 h  | 事務処理全般               | <Finish> |
    +----------------+----------+----------+----------+----------------------------+----------+
  
  デスクトップ名を作業内容として利用するため、あらかじめ必要な仮想デスクトップを作成しておくこと。
  作業時間は本スクリプトで算出して出力するが、 Excel で計算したほうが良いかもしれない。
  
  なお、本スクリプトの実行には VirtualDesktop モジュールが必要となる。

  .parameter EventID

  イベントビューアーから取得したイベント ID 。

  .link

  https://www.powershellgallery.com/packages/VirtualDesktop/1.5.10

  .link

  https://github.com/MScholtes/PSVirtualDesktop

  .link

  https://learn.microsoft.com/ja-jp/microsoftteams/teams-powershell-install
#>
param(
    $EventID
)

#================================================================================
# デバッグ用関数。
#================================================================================

$IsDebugOn = $false
function Debug-Output($Path, $Value) {

    if ($IsDebugOn -eq $true) {

        Add-Content -Path $Path -Value "#[$(Split-Path $MyInvocation.ScriptName -leaf): $($MyInvocation.ScriptLineNumber)] $Value" -Encoding UTF8

    }

}

#================================================================================
# 列挙体の定義。
#================================================================================

#--------------------------------------------------------------
# イベントID で分岐を行えるように列挙体を定義する。
# 参考：https://laboradian.com/win-system-power-state-events/
#--------------------------------------------------------------
Enum EventIDs {
    FastShutdown = 187;        # 高速シャットダウン                   System/Kernel-Power: 187
    Shutdown = 109;            # シャットダウン                       System/Kernel-Power: 109
    PowerOn = 27;              # パワーオン                           System/Kernel-Boot:   27
    Hybernate = 109;           # 休止状態                             System/Kernel-Power: 109
    ModernStandbyStart = 506;  # Modern Standby 開始（画面消灯など）  System/Kernel-Power: 506
    ModernStandbyEnd = 507;    # Modern Standby 終了（画面点灯）      System/Kernel-Power: 507
}

#================================================================================
# 変数の宣言。
#================================================================================

#-------------------------------------------------------------------
# 出力先ファイルパス。
# プロジェクトフォルダ配下の log フォルダへログファイルを出力する。
# 別の場所に置きたい場合は好きなパスを指定すればよい。
#-------------------------------------------------------------------
$LogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.log"            # ログ記録先ファイルパス。
$BackupLogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.bk.log"   # ログ記録前に既存のログ記録先ファイルをバックアップするためのパス。
$DebugLogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\DebugRecord-WorkedTimeAndDesktopName.log"  # デバッグログの出力先ファイル。

#----------------------------------------------------------------------------------------------------
# 記録する日時。
# わかりやすさのため、日付部分だけを変数化しておく（半角スペースで分割し、日付部分だけを取り出す。）
#----------------------------------------------------------------------------------------------------
$CurrentDateTime = Get-Date
$CurrentDateTimeFormatted = $CurrentDateTime.ToString("yyyy/MM/dd HH:mm")
$CurrentDateTimeFormattedForBackupLog = $CurrentDateTime.ToString("yyyy-MM-dd_HHmmss")
$CurrentDate = (-split $CurrentDateTimeFormatted)[0]

#----------------------------------------------------------------------------------------------------
# 記録する日時とデスクトップ名。
# わかりやすさのため、日付部分だけを変数化しておく（半角スペースで分割し、日付部分だけを取り出す。）
#----------------------------------------------------------------------------------------------------
$CurrentDesktopName = Get-DesktopName

#--------------------------------------------------------------------------------------------------------------------
# PC を操作せずディスプレイがオフになった場合など、 Modern Standby 状態に遷移した後の記録に利用するデスクトップ名。
# ここに定める時間を超えても Modern Standby 状態である場合は放置中であることがわかるようにデスクトップ名を記録する。
#--------------------------------------------------------------------------------------------------------------------
$LeftDesktopName = '放置'
$LeavingLimitHours = 0.3  # 0.3h (= 18m).

#--------
#
#--------
# System/Kernel-Power から取得する Modern Standby の開始/終了状態。
# ID: 506 と ID: 507 を両方ともイベントログから取得し、
#   ・最新のログが ID: 506 であれば Modern Standby 開始状態
#   ・最新のログが ID: 507 であれば Modern Standby 終了状態
# であると判定できる。
$LastModernStandbyEvent = (Get-WinEvent -FilterHashtable @{
                        LogName = 'System';
                        ProviderName = 'Microsoft-Windows-Kernel-Power';
                        Id = 506, 507; } -MaxEvents 1)

# 最後の Modern Standby 開始イベントを取得する。
$LastModernStandbyStartEvent = (Get-WinEvent -FilterHashtable @{
                            LogName = 'System';
                            ProviderName = 'Microsoft-Windows-Kernel-Power';
                            Id = 506; } -MaxEvents 1)

# 最後の Modern Standby 開始イベントの作成日時を取得する。
$LastModernStandbyStartEventCreatedDate = $LastModernStandbyStartEvent.TimeCreated.ToString("yyyy/MM/dd HH:mm")

# 放置時間を算出する。現在時刻から Modern Standby 開始イベントの開始時刻を減算して求める。
# 放置時間は比較を行うため Int に型変換する。
$LeftDateTime = $CurrentDateTime - [DateTime]::ParseExact($LastModernStandbyStartEventCreatedDate, "yyyy/MM/dd HH:mm", $null)
$IntLeftHours = [Int]([Float]([String]::Format("{0:F1}", $LeftDateTime.TotalHours)) * 10) # "{0:F1}" -f xx.TotalHours とも書ける。

# 現在の稼働時間と放置判定する時間を比較するため、Int 型に置換する。
$IntLeavingLimitHours = [Int]([Float]$LeavingLimitHours * 10)

#----------------------------------------------------------------------------------------
# 開始・終了時に書き込む目印。
# 開始前・終了後にタイマータスクが呼ばれた場合には目印としてアップデートマークを付ける。
#----------------------------------------------------------------------------------------
$StartMark = "<Start>"
$FinishMark = "<Finish>"
$UpdateMark = "<Update-before-start-or-after-finish>"

#================================================================================
# 判定・記録の実行。
#================================================================================

#------------------------------------------------------------------------------
# 出力先ファイルが存在する場合は内容を更新し、
# 存在しない場合は初めてスクリプトを動かしたとみなしてファイルを新規作成する。
#------------------------------------------------------------------------------
if ((Test-Path $LogFilePath) -eq $true) {

    #------------------------------------------------------------------------------
    # 出力先ファイルの最終行を読み取って正規表現で各項目に分割し、内容を取得する。
    # 読み取れた内容に応じて書き込み内容を変更する。
    #------------------------------------------------------------------------------
    $IsMatched = (Get-Content -Tail 1 -Path $LogFilePath -Encoding UTF8) -match "^(?<Date>.+?)\t(?<StartedDateTime>.+?)\t(?<FinishedDateTime>.+?)\t(?<WorkedTime>.+?)\t(?<DesktopName>.*?)\t(?<StartFinishMark>.*?)$"
 
    #------------------------------------------------------------------------------
    # 最終行が正しく書き込まれている場合、正規表現と一致する。
    # 最終行の各項目の値が取得できているので、それを利用して必要な内容を書き込む。
    #
    # 最終行の内容が正規表現と一致しない場合、最終行は想定した記述になっていない。
    # その場合、仕方がないのでその行はあきらめて新しい行に開始の情報を追記する。
    #
    # 出力先ファイルのエンコードは UTF8NoBOM にしたかったが、
    # 古い PowerShell は UTF8NoBOM 非対応なので BOM 付きの UTF8 を利用する。
    #------------------------------------------------------------------------------
    if ($IsMatched -eq $true) {
 
        #----------------------------------------------------------------
        # 出力先ファイルをコピーし、バックアップファイルとして保存する。
        #----------------------------------------------------------------
        Copy-Item -Path $LogFilePath -Destination $BackupLogFilePath

        # 最終行以外の行を取得する。
        $Content = Get-Content -Path $LogFilePath | Select-Object -SkipLast 1 | Out-String
 
        # 作業時間を算出する。
        $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
        $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours とも書ける。

        #----------------
        # 開始時の処理。
        #----------------
        if ($EventID -eq [EventIDs]::PowerOn.Value__) {
            
            # 開始時の処理は連続で呼ばれることがある。
            # そのため、取得した最終行にスタートマークが存在する場合は何もしない。
            if ($Matches['StartFinishMark'] -like "${StartMark}*") {
                
                # 何もしない。
                Debug-Output -Path $DebugLogFilePath -Value "-- 2-1 PowerOn;`r`n"

            }
            # 最終行にスタートマークが存在しないのであれば、開始時の処理を続行する。
            else {

                # ログオン前にタスクが呼ばれた場合にデスクトップ名が取得できないことがある。
                # その場合、最後のデスクトップ名を利用して行を追加する。追加した行にはスタートマークを付与する。
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value " -- 2-2 PowerOn;`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}"
 
                }
                # デスクトップ名が取得できた場合は行を追加する。
                else {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "-- 2-3 PowerOn;`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}"
 
                }

            }
 
        }
        #----------------
        # 終了時の処理。
        #----------------
        elseif ($EventID -eq [EventIDs]::FastShutdown.Value__ -or
                $EventID -eq [EventIDs]::Shutdown.Value__ -or
                $EventID -eq [EventIDs]::Hybernate.Value__) {
 
            # 終了時の処理は連続で呼ばれることがある。
            # そのため、取得した最終行にフィニッシュマークが存在する場合は何もしない。
            if ($Matches['StartFinishMark'] -like "*${FinishMark}") {
                
                # 何もしない。
                Debug-Output -Path $DebugLogFilePath -Value "-- 3-1 PowerOff;`r`n"

            }
            # 最終行にフィニッシュマークが存在しないのであれば、終了時の処理を続行する。
            # 開始時から終了時まで同じデスクトップで作業していた場合、スタートマークとフィニッシュマークを両方書き込む形とする。
            else {

                # ログオフ後にタスクが呼ばれた場合にデスクトップ名が取得できないことがある。
                # その場合、最後のデスクトップ名を利用して行を更新する。更新した行にはフィニッシュマークを付与する。
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "-- 3-2 PowerOff;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}"
 
                }
                # デスクトップ名が取得できた場合は行を追加する。
                else {
 
                    # 同じデスクトップ名を取得出来た場合、作業を継続しているとみなして同じ行を更新する。
                    # 更新した行にはフィニッシュマークを付与する。
                    if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
                    
                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "-- 3-3 PowerOff;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}"
                        
                    }
                    # 異なるデスクトップ名を取得できた場合、最後に別の作業を行ったと判断して行を追加する。
                    # 追加した行にはフィニッシュマークを付与する。
                    else {

                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "-- 3-4 PowerOff;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}"

                    }
 
                }

            }
 
        }
        #------------------------------------------------------------------------------------------------------------------------------------------------
        # 電源ボタンを押下して画面を消灯する際の処理（PC によっては手動で電源ボタンを押下すると Modern Standby を開始することがある。）
        # この場合、作業を行っていないとみなし、現在のデスクトップ名を "放置" に設定して次の行を開始する。
        # また、電源ボタンを押下した場合は低消費電力状態に遷移し、さらに低消費電力状態にもなる場合があるようで、遷移後は定刻になってもタスクが実行されなくなり、
        # 画面復帰後に改めてタスクが実行されるケースを確認した。
        #------------------------------------------------------------------------------------------------------------------------------------------------ 
        elseif ($EventID -eq [EventIDs]::ModernStandbyStart.Value__) {
            
            # 遅れてタスクが実行され、Modern Standby 開始イベントがを契機にタスクを実行しているものの、最新のイベントは Modern Standby 終了イベントであるケースに対処する。
            # その場合、途中経過が一切記録されていないため、このタイミングで「放置」の行を作成する。
            # その際、既存行はそのまま、既存行の終了時刻から現在時刻までの期間を「放置」とする。
            if ($LastModernStandbyEvent.Id -eq [EventIDs]::ModernStandbyEnd.Value__) {

                $LeavedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['FinishedDateTime'], "yyyy/MM/dd HH:mm", $null)
                $LeavedTime = [String]::Format("{0:F1}", $LeavedHours.TotalHours) # "{0:F1}" -f xx.TotalHours とも書ける。

                Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t$($Matches['FinishedDateTime'])`t$($Matches['WorkedTime'])`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t$($Matches['FinishedDateTime'])`t${CurrentDateTimeFormatted}`t${LeavedTime}`t${LeftDesktopName}`t" -NoNewline -Encoding UTF8
                Debug-Output -Path $DebugLogFilePath -Value "-- 4-1 ModernStandbyStart; `$EventID: $EventID`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t$($Matches['FinishedDateTime'])`t$($Matches['WorkedTime'])`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t$($Matches['FinishedDateTime'])`t${CurrentDateTimeFormatted}`t${LeavedTime}`t${LeftDesktopName}`t"

            }
            # 想定通りに電源ボタン押下のタイミングで、もしくは自動消灯のタイミングで Modern Standby 開始イベントでタスクを起動できた場合、
            # 現在の行を完了して次の行に "放置" を記録する。
            else {

                Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${LeftDesktopName}`t" -NoNewline -Encoding UTF8
                Debug-Output -Path $DebugLogFilePath -Value "-- 4-2 ModernStandbyStart; `$EventID: $EventID`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t"

            }

        }
        #-----------------
        # 5 分ごとの処理。
        #-----------------
        else {

            # ログオン前、ログオフ後にこのルートに入る可能性がある。
            # その呼び出しを止めることはできないので、アップデートマークを付けて行を追加する。
            if ($Matches['StartFinishMark'] -like "*${FinishMark}") {
                
                Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${UpdateMark}" -NoNewline -Encoding UTF8
                Debug-Output -Path $DebugLogFilePath -Value "-- 5-1 Update;`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${UpdateMark}"

            }
            # フィニッシュマークが存在しないのであれば、通常の 5 分ごとの処理を続行する。
            else {
 
                # 「放置」状態の記録を更新する場合、許容時間以内であれば「放置」ではなかった扱いにする。
                # 許容時間を超えてから復帰した場合は「放置」のままにする。
                if ($Matches['DesktopName'] -eq $LeftDesktopName) {
                    
                    # Modern Standby 開始状態である場合は引き続き「放置」として記録する。
                    if ($LastModernStandbyEvent.Id -eq [EventIDs]::ModernStandbyStart.Value__) {
                        
                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${LeftDesktopName}`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "-- 5-2 ModernStandbyStart (506); `$LastModernStandbyEvent.Id: $($LastModernStandbyEvent.Id)"
                        Debug-Output -Path $DebugLogFilePath -Value "-- 5-2 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${LeftDesktopName}`t$($Matches['StartFinishMark'])"
    
                    }
                    # Modern Standby 終了状態であり、かつ放置時間が放置許容時間以上である場合は作業に復帰したとみなす。
                    # その際、放置許容時間以内である場合は本来のデスクトップ名で記録する。
                    # 放置時間が放置許容時間を超過する場合は「放置」したとみなす。
                    else {
                        
                        # 許容時間を超過してから復帰したとみなし、「放置」を記録して新しい行を始める。
                        if ($IntLeftHours -gt $IntLeavingLimitHours) {

                            #Debug-Output -Path $DebugLogFilePath -Value "-- 5-3 ModernStandbyEnd (507); `$LastModernStandbyEvent.Id: $($LastModernStandbyEvent.Id) and `$IntLeftHours > `$IntLeavingLimitHours: $IntLeftHours > $IntLeavingLimitHours"
                            #Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8

                            Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${LeftDesktopName}`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
                            Debug-Output -Path $DebugLogFilePath -Value "-- 5-3 ModernStandbyEnd (507); `$LastModernStandbyEvent.Id: $($LastModernStandbyEvent.Id) and `$IntLeftHours > `$IntLeavingLimitHours: $IntLeftHours > $IntLeavingLimitHours"
                            Debug-Output -Path $DebugLogFilePath -Value "-- 5-3 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${LeftDesktopName}`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t"

                        }
                        # 許容時間以内に復帰したとみなし、「放置」を本来の仮想デスクトップ名に置き換える。
                        else {
                            
                            # $CurrentDesktopName は本来の値を使用する。
                            # Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${CurrentDesktopName}`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                            # Debug-Output -Path $DebugLogFilePath -Value "-- 5-5 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${CurrentDesktopName}`t$($Matches['StartFinishMark'])"

                            Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${CurrentDesktopName}`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                            Debug-Output -Path $DebugLogFilePath -Value "-- 5-4 ModernStandbyEnd (507); `$LastModernStandbyEvent.Id: $($LastModernStandbyEvent.Id) and `$IntLeftHours <= `$IntLeavingLimitHours: $IntLeftHours <= $IntLeavingLimitHours"
                            Debug-Output -Path $DebugLogFilePath -Value "-- 5-4 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t${CurrentDesktopName}`t$($Matches['StartFinishMark'])"

                        }
    
                    }

                }
                # 同じデスクトップ名を取得出来た場合、作業を継続しているとみなして同じ行を更新する。
                elseif ($Matches['DesktopName'] -eq $CurrentDesktopName) {

                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "-- 5-6 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])"
                    
                }
                # 異なるデスクトップ名が取得できた場合、別作業を開始したとみなして次の行を開始する。
                else {
                    
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "-- 5-7 Update;`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t"

                }

            }
 
        }
    
    }
    # 最終行の記述が想定通りでない場合、バックアップファイルに日付時刻を付与してさらにバックアップする。
    # その後、"<Something wrong!>"、改行、「年月日」、「開始時刻」、「作業時間」、「作業種別（デスクトップ名）」を書き込む。
    else {
    
        $BackupLogFilePathNow = $BackupLogFilePath -replace "bk" , $CurrentDateTimeFormattedForBackupLog
        Copy-Item -Path $BackupLogFilePath -Destination $BackupLogFilePathNow
        
        Add-Content -Path $LogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
        Debug-Output -Path $DebugLogFilePath -Value "-- 6 something wrong;`r`n<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -Encoding UTF8
    
    }
}
# 出力先ファイルが存在しない場合、ファイルを新規作成して「年月日」、「開始時刻」、「終了時刻」「作業時間」、「作業種別（デスクトップ名）」を書き込む。
else {
    
    Set-Content -Path $LogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
    Debug-Output -Path $DebugLogFilePath -Value "-- 7 Create File;`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -Encoding UTF8
    
}

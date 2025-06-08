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

# イベントID で分岐を行えるように列挙体を定義する。
# 参考：https://laboradian.com/win-system-power-state-events/
Enum EventIDs {
    FastShutdown = 187;  # 高速シャットダウン System/Kernel-Power:187
    Shutdown = 109;      # シャットダウン     System/Kernel-Power:109
    PowerOn = 27;        # パワーオン         System/Kernel-Boot:27
    Hybernate = 109;     # 休止状態           System/Kernel-Power:109
}

# プロジェクトフォルダ配下の log フォルダへログファイルを出力する。
# 別の場所に置きたい場合は好きなパスを指定すればよい。
$LogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.log"
$DebugLogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\DebugRecord-WorkedTimeAndDesktopName.log"

# 記録する日時とデスクトップ名を取得する。
# わかりやすさのため、日付部分だけを変数化しておく（半角スペースで分割し、日付部分だけを取り出す。）
$CurrentDateTime = Get-Date
$CurrentDateTimeFormatted = $CurrentDateTime.ToString("yyyy/MM/dd HH:mm")
$CurrentDate = (-split $CurrentDateTimeFormatted)[0]
$CurrentDesktopName = Get-DesktopName

# 開始・終了時に書き込む目印を定義する。
# 開始前・終了後にタイマータスクが呼ばれた場合には目印としてアップデートマークを付ける。
$StartMark = "<Start>"
$FinishMark = "<Finish>"
$UpdateMark = "<Update-before-start-or-after-finish>"

# デバッグ用関数。
$IsDebugOn = $false
function Debug-Output($Path, $Value) {

    if ($IsDebugOn -eq $true) {

        Add-Content -Path $Path -Value $Value -Encoding UTF8

    }

}
 
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
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n:1-0 PowerOn"

            }
            # 最終行にスタートマークが存在しないのであれば、開始時の処理を続行する。
            else {

                # ログオン前にタスクが呼ばれた場合にデスクトップ名が取得できないことがある。
                # その場合、最後のデスクトップ名を利用して行を追加する。追加した行にはスタートマークを付与する。
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}:1-1 PowerOn"
 
                }
                # デスクトップ名が取得できた場合は行を追加する。
                else {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}:1-2 PowerOn"
 
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
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n:2-0 PowerOff"

            }
            # 最終行にフィニッシュマークが存在しないのであれば、終了時の処理を続行する。
            # 開始時から終了時まで同じデスクトップで作業していた場合、スタートマークとフィニッシュマークを両方書き込む形とする。
            else {

                # ログオフ後にタスクが呼ばれた場合にデスクトップ名が取得できないことがある。
                # その場合、最後のデスクトップ名を利用して行を更新する。更新した行にはフィニッシュマークを付与する。
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}:2-1 PowerOff"
 
                }
                # デスクトップ名が取得できた場合。は行を追加する。
                else {
 
                    # 同じデスクトップ名を取得出来た場合、作業を継続しているとみなして同じ行を更新する。
                    # 更新した行にはフィニッシュマークを付与する。
                    if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
                    
                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}:2-2 PowerOff"
                        
                    }
                    # 異なるデスクトップ名を取得できた場合、最後に別の作業を行ったと判断して行を追加する。
                    # 追加した行にはフィニッシュマークを付与する。
                    else {

                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}:2-3 PowerOff"

                    }
 
                }

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
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${UpdateMark}:3-0 Update"

            }
            # フィニッシュマークが存在しないのであれば、通常の 5 分ごとの処理を続行する。
            else {
 
                # 同じデスクトップ名を取得出来た場合、作業を継続しているとみなして同じ行を更新する。
                if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
                
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark']):3-1 Update"
                    
                }
                # 異なるデスクトップ名が取得できた場合、別作業を開始したとみなして次の行を開始する。
                # その際にマークは消す。
                else {
                    
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:3-2 Update"

                }

            }
 
        }
    
    }
    # 最終行の記述が想定通りでない場合、"<Something wrong!>"、改行、「年月日」、「開始時刻」、「作業時間」、「作業種別（デスクトップ名）」を書き込む。
    else {
    
        Add-Content -Path $LogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
        Debug-Output -Path $DebugLogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:5 something wrong" -Encoding UTF8
    
    }
}
# 出力先ファイルが存在しない場合、ファイルを新規作成して「年月日」、「開始時刻」、「終了時刻」「作業時間」、「作業種別（デスクトップ名）」を書き込む。
else {
    
    Set-Content -Path $LogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
    Debug-Output -Path $DebugLogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:6 Create File" -Encoding UTF8
    
}

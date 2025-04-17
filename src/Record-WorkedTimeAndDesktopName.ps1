<#
  .synopsis

  現在の時刻とデスクトップ名を取得し、タブ区切りテキストとしてファイルへ書き込む。

  .description

  書き込んだタブ区切りテキストを Excel に貼り付けると、以下のような表を作成できる。

    +----------------+----------+----------+----------+----------------------------+
    |     年月日     | 開始時刻 | 終了時刻 | 作業時間 | 作業種別（デスクトップ名） |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(土) |  09:00   |  12:00   |   3.0 h  | Project A                  |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(土) |  13:00   |  15:00   |   2.0 h  | 教育                       |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(土) |  15:00   |  17:30   |   2.5 h  | 事務処理全般               |
    +----------------+----------+----------+----------+----------------------------+
  
  デスクトップ名を作業内容として利用するため、あらかじめ必要なデスクトップを作成しておくこと。
  また、作業時間は本スクリプトでも出力するが、 Excel で計算したほうが良いかもしれない。
  
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
    FastShutdown = 187;
    Shutdown = 109;
    PowerOn = 27;
    Hybernate = 109;
}

# プロジェクトフォルダ配下の log フォルダへログファイルを出力する。
# 別の場所に置きたい場合は好きなパスを指定すればよい。
$LogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.log"


# 記録する日時とデスクトップ名を取得する。
# わかりやすさのため、日付部分だけを変数化しておく（半角スペースで分割し、日付部分だけを取り出す。）
$CurrentDateTime = Get-Date
$CurrentDateTimeFormatted = $CurrentDateTime.ToString("yyyy/MM/dd HH:mm")
$CurrentDate = (-split $CurrentDateTimeFormatted)[0]
$CurrentDesktopName = Get-DesktopName

#------------------------------------------------------------------------------
# 出力先ファイルが存在する場合は内容を更新し、
# 存在しない場合は初めてスクリプトを動かしたとみなしてファイルを新規作成する。
#------------------------------------------------------------------------------
if ((Test-Path $LogFilePath) -eq $true) {
    #------------------------------------------------------------------------------
    # 出力先ファイルの最終行を読み取って正規表現で各項目に分割し、内容を取得する。
    # 読み取れた内容に応じて書き込み内容を変更する。
    #------------------------------------------------------------------------------
    $IsMatched = (Get-Content -Tail 1 -Path $LogFilePath -Encoding UTF8) -match "^(?<Date>.+)\t(?<StartedDateTime>.+)\t(?<FinishedDateTime>.+)\t(?<WorkedTime>.+)\t(?<DesktopName>.+)$"

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

        if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
            # 最終行のデスクトップ名と現在のデスクトップ名が等しい場合、作業を継続しているとみなして終了時刻を現在時刻へ更新する。
            # ただし、 PC 起動時は作業の継続ではないとみなし、最終行を更新せずに次の行へ追記する。
            if ($EventID -eq [EventIDs]::PowerOn.Value__) {

                Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

            } else {
                $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
                $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours とも書ける。

                Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])" -NoNewline -Encoding UTF8
            }
        } else {

            # 最終行のデスクトップ名と現在のデスクトップ名が等しくない場合、別の作業を開始しているとみなす。
            # そのため、最終行の終了時刻に現在時刻を記録し、次の行に現在時刻で作業を開始した旨を追記する。
            $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
            $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours とも書ける。

            Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`r`n$($Matches['Date'])`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

        }
    } else {

        # 最終行の記述が想定通りでない場合、"<Something wrong!>"、改行、「年月日」、「開始時刻」、「作業時間」、「作業種別（デスクトップ名）」を書き込む。
        Add-Content -Path $LogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

    }
} else {

    # 出力先ファイルが存在しない場合、ファイルを新規作成して「年月日」、「開始時刻」、「終了時刻」「作業種別（デスクトップ名）」を書き込む。
    Set-Content -Path $LogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

}
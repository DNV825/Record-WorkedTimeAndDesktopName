# Record-WorkedTimeAndDesktopName オマケ

## Virtual Desktop モジュールに署名が付与されているか確認する手順

Virtual Desktop モジュールに署名が付与されているか確認したい場合、以下のようにコマンドを実行すればよい。署名は付与されていなかったが、これをもって信頼できないかというとそういうわけでもない。

```powershell
Save-Module -Name VirtualDesktop -Repository PSGallery -Path .\Saved

Get-ChildItem .\Saved\VirtualDesktop\*\*.ps*1 |
    Get-AuthenticodeSignature |
    Select-Object Path, Status, SignerCertificate | Format-Table -Auto

Path                                                                                                                Status SignerCertificate
----                                                                                                                ------ -----------------
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.ps1  NotSigned                  
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.psd1 NotSigned                  
C:\<ダウンロードした場所のパス>\Saved\VirtualDesktop\1.5.10\VirtualDesktop.psm1 NotSigned  
```

## 管理者権限でバッチファイルを起動するとカレントディレクトリが変わってしまう問題への対応

以下、カレントディレクトリをバッチファイルと同じ場所にするイディオム。
管理者権限ありでバッチファイルを起動すると、カレントディレクトリが別の場所に移動してしまうので、この処理が必要となる。

```bat
cd /d %~dp0
```

## タスクスケジューラーに関する情報

### タスクスケジューラーの起動モード確認手順

PowerShell で以下のコマンドを入力する。 ".exe" も入力しないといけない。

```powershell
PS C:\Users\xxx> sc.exe qc schedule
[SC] QueryServiceConfig SUCCESS

SERVICE_NAME: schedule
        TYPE               : 20  WIN32_SHARE_PROCESS
        START_TYPE         : 2   AUTO_START
        ERROR_CONTROL      : 1   NORMAL
        BINARY_PATH_NAME   : C:\WINDOWS\system32\svchost.exe -k netsvcs -p
        LOAD_ORDER_GROUP   : SchedulerGroup
        TAG                : 0
        DISPLAY_NAME       : Task Scheduler
        DEPENDENCIES       : RPCSS
                           : SystemEventsBroker
        SERVICE_START_NAME : LocalSystem
```

グループポリシーや最適化によって `START_TYPE` が `DelayAutoStart = True` で置き換えられていることがあるらしく、その場合は色々な手続きが終わってからタスクスケジュールが起動するため、イベントを拾えない可能性が高まる。組織や学校のアカウント、 Microsoft アカウントを使っている場合はログオンにも時間がかかるので、タスクスケジュールの起動がさらに遅れる。

タスクスケジュールの起動が遅れると、起動イベントが遠い過去に発生したことになってしまい、イベントが拾えなくなる。

### タスク実行時に使うユーザーアカウントとキャッチ可能なイベント / トリガー

| タスクの実行時に使うユーザーアカウント | キャッチ可能なイベント / トリガー |
| :--- | :--- |
| SYSTEM | BootTrigger, LogonTrigger, Microsoft-Windows-Power-Troubleshooter:1, User32:1074 |
| USERS | LogonTrigger, Microsoft-Windows-Power-Troubleshooter:1, User32:1074 |
| 個別のユーザー | LogonTrigger, Microsoft-Windows-Power-Troubleshooter:1, User32:1074 |

目的：

- BootTrigger ... PC の電源オンをキャッチする
- LogonTrigger ... ユーザーのログオンをキャッチする
- Microsoft-Windows-Power-Troubleshooter:1 ... 休止状態からの復帰をキャッチする
- User32:1074 ... GUIからのシャットダウンや休止状態への遷移をキャッチする

### SID を取得する方法

タスクスケジューラーのタスクは特定のユーザーが起動するように設定できるが、その際に SID と呼ばれる特別な ID を指定している。 GUI で設定する場合は自動的に SID を取得してくれるが、 XML ファイルに直接記述する場合は別の手段で SID を取得しなければならない。

PowerShell で SID を取得する手順は以下の通り。

```powershell
$ObjUser = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, $env:USERNAME)
$UserSid = $ObjUser.Translate([System.Security.Principal.SecurityIdentifier])
```

### Windows 11 のエディションを取得する方法

Windows 11 のエディションは以下のコマンドレットで取得できる。

```powershell
PS C:\Users\xxx> (Get-ComputerInfo).WindowsEditionID
```

取得できる値は以下の通り。なぜか Home は Core という値が設定されている。

- Core ＝ Home
- Professional ＝ Pro
- Enterprise ＝ Enterprise

### Windows 11 であるかを判定する方法

これがなかなか厄介で、今回は実装しなかった。後方互換性を確保するための処置（ [Windows 10 向けに作ったアプリが Windows 11 でも動くようにするため](https://answers.microsoft.com/en-us/windows/forum/all/why-does-windows-11-show-windows-10-as-os-version/2da93b57-951b-4771-a328-36ab5ff685d7?utm_source=chatgpt.com)）らしいが、Windows 11 なのに "Windows 10" の文言が取得できてしまう。

`OsName` には "Windows 11" の文言が含まれているのでこれを使えば良さそうであるが、どうやら `OsbuildNumber` で判別する方法が主流なようだ。

```powershell
PS C:\Users\xxx> (Get-ComputerInfo).WindowsProductName
Windows 10 Home

PS C:\Users\xxx> (Get-ComputerInfo).OsName
Microsoft Windows 11 Home

PS C:\Users\xxx> (Get-ComputerInfo).OsBuildNumber
26100
```

[22000 以降（Release 21H2）なら Windows 11、19044 以前なら Windows 10](https://learn.microsoft.com/ja-jp/windows/release-health/windows11-release-information) らしい。


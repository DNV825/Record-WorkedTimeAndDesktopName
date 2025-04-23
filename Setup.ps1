<#
  .synopsis

  Record-WorkedTimeAndDesktopName のインストール、またはアンインストールを実行する。

  .description
    
  インストールする場合は以下の手続きを実行する。

    1. Virtual Desktop モジュールをインストールする。
    2. タスクスケジューラーへタスク「Record-WorkedTimeAndDesktopName」を登録する。

  アンインストールする場合は以下の手続きを実行する。

    1. タスクスケジューラーからタスク「Record-WorkedTimeAndDesktopName」を削除する。
    2. Virtual Desktop モジュールをアンインストールする。

  コマンドプロンプトからの実行を想定しており、管理者権限ありでコマンドプロンプトを起動する必要がある。
  
  .notes

  VirtualDesktop モジュールのインストール・インポート時に以下の警告が表示されるが、どちらも Y を入力して進めること。

    続行するには NuGet プロバイダーが必要です
      PowerShellGet で NuGet ベースのリポジトリを操作するには、'2.8.5.201' 以降のバージョンの NuGet
      プロバイダーが必要です。NuGet プロバイダーは 'C:\Program Files\PackageManagement\ProviderAssemblies' または
      'C:\Users\xxx\AppData\Local\PackageManagement\ProviderAssemblies' に配置する必要があります。'Install-PackageProvider
      -Name NuGet -MinimumVersion 2.8.5.201 -Force' を実行して NuGet プロバイダーをインストールすることもできます。今すぐ
      PowerShellGet で NuGet プロバイダーをインストールしてインポートしますか?
      [Y] はい(Y)  [N] いいえ(N)  [S] 中断(S)  [?] ヘルプ (既定値は "Y"): Y

    信頼されていないリポジトリ
      信頼されていないリポジトリからモジュールをインストールしようとしています。このリポジトリを信頼する場合は、Set-PSReposit
      ory コマンドレットを実行して、リポジトリの InstallationPolicy の値を変更してください。'PSGallery'
      からモジュールをインストールしますか?
      [Y] はい(Y)  [A] すべて続行(A)  [N] いいえ(N)  [L] すべて無視(L)  [S] 中断(S)  [?] ヘルプ (既定値は "N"): Y

  NuGet はモジュールのインストールに必要なツールのため問題ない。
  Virtual Desktop モジュールは信頼して使うことにする。

  .parameter ActionType

  "Install" か "Uninstall" のいずれかを指定する。
#>
Param (
    [Parameter(Mandatory=$true)]
    [String]
    $ActionType
)

<#
    VirtualDesktop がインストールされているか。
#>
function IsVirtualDesktopModuleInstalled() {

    if ((Get-Module -ListAvailable -Name VirtualDesktop) -eq $null) {

        return $false

    } else {

        return $true

    }

}


<#
    タスクがタスクスケジューラーに登録されているか。
#>
function IsTaskRegistered($TaskName) {

    # タスクが登録されていない場合、標準出力にエラーが表示されるため、
    # 標準出力と標準エラー出力を抑止しつつタスクの存在を確認する。
    # 結果は標準出力で帰ってくるため、文字列を解析して真偽値を返す。
    $ErrorActionPreference = "silentlycontinue"
    ($TaskQueryString = (schtasks /query /tn $TaskName 2> $null)) | Out-Null
    $ErrorActionPreference = "continue"

    if ($TaskQueryString -match $TaskName) {

        return $true

    } else {

        return $false

    }

}

<#
    Record-WorkedTimeAndDesktopName をインストールする。
#>
function Install($TaskName, $Windows11Edition) {

    #=====================================================
    # (1/2) Virtual Desktop モジュールをインストールする。
    #=====================================================
    Write-Host "===================================================================================="
    Write-Host "  ${TaskName} のインストール手続きを開始します。"
    Write-Host "  管理者権限ありで実行してください。"
    Write-Host "===================================================================================="
    Write-Host "(1/2) Virtual Desktop モジュールのインストール"

    # Virtual Desktop がインストールされている場合、ここでは何もしない。
    # Virtual Desktop がインストールされていない場合、インストールを開始する。
    if (IsVirtualDesktopModuleInstalled -eq $true) {

        Write-Host "      Virtual Desktop モジュールはインストール済みでした。"

    } else {

        Write-Host "      Virtual Desktop モジュールをインストールします。"
        Write-Host "      インストールしたくない場合、 n を入力してインストールを終了してください。"
        Write-Host "      選択肢が表示されるまで、少々お待ちください。"
        Install-Module VirtualDesktop -Scope CurrentUser
        Import-Module VirtualDesktop
        $DesktopName = Get-DesktopName

        if ($DesktopName -eq $null) {

            # VirtualDesktop をインストールしなかった場合、インストール手続きを終了する。
            Write-Host "      Virtual Desktop モジュールをインストールしなかったか、もしくはインストールに失敗しました。"
            Write-Host "      インストール手続きを終了します。"
            return

        }

    }

    Write-Host ""

    #================================================================================
    # (2/2) タスクスケジューラーにタスク Record-WorkedTimeAndDesktopName を登録する。
    #================================================================================
    Write-Host "(2/2) タスクスケジューラーにタスク ${TaskName} を登録"
    Write-Host "      現在のフォルダのパスでタスクスケジューラーにタスクを登録します。"
    Write-Host "      インストール後にフォルダを移動する場合は再度インストールを実施してください。"

    # インストールに利用する xml ファイルの元ネタファイルを読み込み、スクリプトのパスをカレントディレクトリに置換する。
    # また、ユーザー SID をインストールを実行したユーザーの SID に置換する。
    $ObjUser = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, $env:USERNAME)
    $UserSid = $ObjUser.Translate([System.Security.Principal.SecurityIdentifier])

    $XmlContent = (Get-Content "./src/${TaskName}-${Windows11Edition}.xml.in" -Encoding Unicode).Replace("[@Install-Destination]", "$(Split-Path $PSCommandPath -Parent)").Replace("[@UserSID]", "${UserSid}")
    Set-Content -Path "./src/${TaskName}.xml" -Value $XmlContent -Encoding Unicode

    # タスクスケジューラーにタスク名 Record-WorkedTimeAndDesktopName が登録されている場合、
    # いったん削除して再登録するかユーザーに尋ね、ユーザーが登録を辞める選択をした場合はインストール手続きを終了する。
    # 登録されていない場合は単純にタスクをインポートする。
    if ((IsTaskRegistered $TaskName) -eq $true) {

        Write-Host "      すでにタスク名 ${TaskName} が登録されています。"
        $UserInput = Read-Host "      既存の ${TaskName} を削除し、新たに登録しなおしますか？ [y/N]"

        if ($UserInput -eq "y" -or $UserInput -eq "Y") {

            Write-Host "      タスク名 ${TaskName} を削除し、再登録します。"
            schtasks /end /tn $TaskName
            schtasks /delete /tn $TaskName
            schtasks /create /tn "${TaskName}" /xml ".\src\${TaskName}.xml"

        } else {
            
            Write-Host "      タスクの登録は行わず、インストール手続きを終了します。"

        }

    } else {

        Write-Host "      タスク名 ${TaskName} を登録します。"
        schtasks /create /tn "${TaskName}" /xml ".\src\${TaskName}.xml"

    }

    Write-Host "===================================================================================="
    Write-Host "  ${TaskName} をタスクスケジューラーへ登録しました。"

    if ($Windows11Edition -eq "Core") {

        Write-Host "  Windows 11 Home 向けのインストール手続きを完了完了しました。"

    } else {

        Write-Host "  Windows 11 ${Windows11Edition} をご利用の場合、引き続きローカルグループポリシーを起動し、"
        Write-Host "  PowerShell スクリプトを登録してください。登録方法は README.md をご参照願います。"

    }

    Write-Host "===================================================================================="

}

<#
    Record-WorkedTimeAndDesktopName をアンインストールする。
#>
function Uninstall($TaskName, $Windows11Edition) {
    
    #===================================================================================
    # (1/2) タスクスケジューラーからタスク Record-WorkedTimeAndDesktopName を削除する。
    #===================================================================================
    Write-Host "=========================================================================="
    Write-Host "  ${TaskName} のアンインストール手続きを開始します。"
    Write-Host "  管理者権限ありで実行してください。"
    Write-Host "=========================================================================="
    Write-Host "(1/2) タスクスケジューラーからタスク ${TaskName} を削除"

    # タスクスケジューラーにタスク名 Record-WorkedTimeAndDesktopName が登録されている場合はタスクを削除する。
    if ((IsTaskRegistered $TaskName) -eq $true) {

        Write-Host "      タスク名 ${TaskName} を削除します。"
        schtasks /end /tn $TaskName
        schtasks /delete /tn $TaskName

        if ((IsTaskRegistered $TaskName) -eq $true) {

            Write-Host "      タスク名 ${TaskName} の削除に失敗しました。お手数ですが、手動で削除してください。"
            Write-Host "      アンインストール手続きを終了します。"

        }

    } else {

        Write-Host "      タスク名 ${TaskName} は登録されていませんでした。"

    }

    Write-Host ""

    #=========================================================
    # (2/2) Virtual Desktop モジュールをアンインストールする。
    #=========================================================
    Write-Host "(2/2) Virtual Desktop モジュールのアンインストール"

    # Virtual Desktop がインストールされている場合、ここでは何もしない。
    # Virtual Desktop がインストールされていない場合、インストールを開始する。
    if (IsVirtualDesktopModuleInstalled -eq $true) {

        Write-Host "      Virtual Desktop モジュールをアンインストールします。"
        Write-Host "      アンインストールしたくない場合、 n を入力してアンインストールを終了してください。"
        $UserInput = Read-Host "      Virtual Desktop モジュールをアンインストールしますか？ [y/N]"

        if ($UserInput -eq "y"-or $UserInput -eq "Y") {

            Uninstall-Module VirtualDesktop

            if (IsVirtualDesktopModuleInstalled -eq $true) {

                # Virtual Desktop モジュールをアンインストールしなかった場合、アンインストール手続きを終了する。
                Write-Host "      Virtual Desktop モジュールをアンインストールしなかったか、もしくはアンインストールに失敗しました。"
                Write-Host "      アンインストール手続きを終了します。"
                return

            }
        } else {

            Write-Host "      Virtual Desktop モジュールはアンインストールしません。"

        }

    } else {

        Write-Host "      Virtual Desktop モジュールはインストールされていませんでした。"

    }

    Write-Host "=========================================================================="
    Write-Host "  ${TaskName} のアンインストール手続きを完了しました。"
    Write-Host "=========================================================================="

}

#=============================================
# インストール / アンインストールを実行する。
#=============================================
$TaskName = "Record-WorkedTimeAndDesktopName"
$Windows11Edition = (Get-ComputerInfo).WindowsEditionID

switch ($ActionType) {
    "Install" { Install $TaskName $Windows11Edition }
    "Uninstall" { Uninstall $TaskName $Windows11Edition }
    default { Write-Host "第1引数には Install か Uninstall を指定してください。" }
}


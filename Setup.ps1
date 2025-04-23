<#
  .synopsis

  Record-WorkedTimeAndDesktopName �̃C���X�g�[���A�܂��̓A���C���X�g�[�������s����B

  .description
    
  �C���X�g�[������ꍇ�͈ȉ��̎葱�������s����B

    1. Virtual Desktop ���W���[�����C���X�g�[������B
    2. �^�X�N�X�P�W���[���[�փ^�X�N�uRecord-WorkedTimeAndDesktopName�v��o�^����B

  �A���C���X�g�[������ꍇ�͈ȉ��̎葱�������s����B

    1. �^�X�N�X�P�W���[���[����^�X�N�uRecord-WorkedTimeAndDesktopName�v���폜����B
    2. Virtual Desktop ���W���[�����A���C���X�g�[������B

  �R�}���h�v�����v�g����̎��s��z�肵�Ă���A�Ǘ��Ҍ�������ŃR�}���h�v�����v�g���N������K�v������B
  
  .notes

  VirtualDesktop ���W���[���̃C���X�g�[���E�C���|�[�g���Ɉȉ��̌x�����\������邪�A�ǂ���� Y ����͂��Đi�߂邱�ƁB

    ���s����ɂ� NuGet �v���o�C�_�[���K�v�ł�
      PowerShellGet �� NuGet �x�[�X�̃��|�W�g���𑀍삷��ɂ́A'2.8.5.201' �ȍ~�̃o�[�W������ NuGet
      �v���o�C�_�[���K�v�ł��BNuGet �v���o�C�_�[�� 'C:\Program Files\PackageManagement\ProviderAssemblies' �܂���
      'C:\Users\xxx\AppData\Local\PackageManagement\ProviderAssemblies' �ɔz�u����K�v������܂��B'Install-PackageProvider
      -Name NuGet -MinimumVersion 2.8.5.201 -Force' �����s���� NuGet �v���o�C�_�[���C���X�g�[�����邱�Ƃ��ł��܂��B������
      PowerShellGet �� NuGet �v���o�C�_�[���C���X�g�[�����ăC���|�[�g���܂���?
      [Y] �͂�(Y)  [N] ������(N)  [S] ���f(S)  [?] �w���v (����l�� "Y"): Y

    �M������Ă��Ȃ����|�W�g��
      �M������Ă��Ȃ����|�W�g�����烂�W���[�����C���X�g�[�����悤�Ƃ��Ă��܂��B���̃��|�W�g����M������ꍇ�́ASet-PSReposit
      ory �R�}���h���b�g�����s���āA���|�W�g���� InstallationPolicy �̒l��ύX���Ă��������B'PSGallery'
      ���烂�W���[�����C���X�g�[�����܂���?
      [Y] �͂�(Y)  [A] ���ׂđ��s(A)  [N] ������(N)  [L] ���ׂĖ���(L)  [S] ���f(S)  [?] �w���v (����l�� "N"): Y

  NuGet �̓��W���[���̃C���X�g�[���ɕK�v�ȃc�[���̂��ߖ��Ȃ��B
  Virtual Desktop ���W���[���͐M�����Ďg�����Ƃɂ���B

  .parameter ActionType

  "Install" �� "Uninstall" �̂����ꂩ���w�肷��B
#>
Param (
    [Parameter(Mandatory=$true)]
    [String]
    $ActionType
)

<#
    VirtualDesktop ���C���X�g�[������Ă��邩�B
#>
function IsVirtualDesktopModuleInstalled() {

    if ((Get-Module -ListAvailable -Name VirtualDesktop) -eq $null) {

        return $false

    } else {

        return $true

    }

}


<#
    �^�X�N���^�X�N�X�P�W���[���[�ɓo�^����Ă��邩�B
#>
function IsTaskRegistered($TaskName) {

    # �^�X�N���o�^����Ă��Ȃ��ꍇ�A�W���o�͂ɃG���[���\������邽�߁A
    # �W���o�͂ƕW���G���[�o�͂�}�~���^�X�N�̑��݂��m�F����B
    # ���ʂ͕W���o�͂ŋA���Ă��邽�߁A���������͂��Đ^�U�l��Ԃ��B
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
    Record-WorkedTimeAndDesktopName ���C���X�g�[������B
#>
function Install($TaskName, $Windows11Edition) {

    #=====================================================
    # (1/2) Virtual Desktop ���W���[�����C���X�g�[������B
    #=====================================================
    Write-Host "===================================================================================="
    Write-Host "  ${TaskName} �̃C���X�g�[���葱�����J�n���܂��B"
    Write-Host "  �Ǘ��Ҍ�������Ŏ��s���Ă��������B"
    Write-Host "===================================================================================="
    Write-Host "(1/2) Virtual Desktop ���W���[���̃C���X�g�[��"

    # Virtual Desktop ���C���X�g�[������Ă���ꍇ�A�����ł͉������Ȃ��B
    # Virtual Desktop ���C���X�g�[������Ă��Ȃ��ꍇ�A�C���X�g�[�����J�n����B
    if (IsVirtualDesktopModuleInstalled -eq $true) {

        Write-Host "      Virtual Desktop ���W���[���̓C���X�g�[���ς݂ł����B"

    } else {

        Write-Host "      Virtual Desktop ���W���[�����C���X�g�[�����܂��B"
        Write-Host "      �C���X�g�[���������Ȃ��ꍇ�A n ����͂��ăC���X�g�[�����I�����Ă��������B"
        Write-Host "      �I�������\�������܂ŁA���X���҂����������B"
        Install-Module VirtualDesktop -Scope CurrentUser
        Import-Module VirtualDesktop
        $DesktopName = Get-DesktopName

        if ($DesktopName -eq $null) {

            # VirtualDesktop ���C���X�g�[�����Ȃ������ꍇ�A�C���X�g�[���葱�����I������B
            Write-Host "      Virtual Desktop ���W���[�����C���X�g�[�����Ȃ��������A�������̓C���X�g�[���Ɏ��s���܂����B"
            Write-Host "      �C���X�g�[���葱�����I�����܂��B"
            return

        }

    }

    Write-Host ""

    #================================================================================
    # (2/2) �^�X�N�X�P�W���[���[�Ƀ^�X�N Record-WorkedTimeAndDesktopName ��o�^����B
    #================================================================================
    Write-Host "(2/2) �^�X�N�X�P�W���[���[�Ƀ^�X�N ${TaskName} ��o�^"
    Write-Host "      ���݂̃t�H���_�̃p�X�Ń^�X�N�X�P�W���[���[�Ƀ^�X�N��o�^���܂��B"
    Write-Host "      �C���X�g�[����Ƀt�H���_���ړ�����ꍇ�͍ēx�C���X�g�[�������{���Ă��������B"

    # �C���X�g�[���ɗ��p���� xml �t�@�C���̌��l�^�t�@�C����ǂݍ��݁A�X�N���v�g�̃p�X���J�����g�f�B���N�g���ɒu������B
    # �܂��A���[�U�[ SID ���C���X�g�[�������s�������[�U�[�� SID �ɒu������B
    $ObjUser = New-Object System.Security.Principal.NTAccount($env:USERDOMAIN, $env:USERNAME)
    $UserSid = $ObjUser.Translate([System.Security.Principal.SecurityIdentifier])

    $XmlContent = (Get-Content "./src/${TaskName}-${Windows11Edition}.xml.in" -Encoding Unicode).Replace("[@Install-Destination]", "$(Split-Path $PSCommandPath -Parent)").Replace("[@UserSID]", "${UserSid}")
    Set-Content -Path "./src/${TaskName}.xml" -Value $XmlContent -Encoding Unicode

    # �^�X�N�X�P�W���[���[�Ƀ^�X�N�� Record-WorkedTimeAndDesktopName ���o�^����Ă���ꍇ�A
    # ��������폜���čēo�^���邩���[�U�[�ɐq�ˁA���[�U�[���o�^�����߂�I���������ꍇ�̓C���X�g�[���葱�����I������B
    # �o�^����Ă��Ȃ��ꍇ�͒P���Ƀ^�X�N���C���|�[�g����B
    if ((IsTaskRegistered $TaskName) -eq $true) {

        Write-Host "      ���łɃ^�X�N�� ${TaskName} ���o�^����Ă��܂��B"
        $UserInput = Read-Host "      ������ ${TaskName} ���폜���A�V���ɓo�^���Ȃ����܂����H [y/N]"

        if ($UserInput -eq "y" -or $UserInput -eq "Y") {

            Write-Host "      �^�X�N�� ${TaskName} ���폜���A�ēo�^���܂��B"
            schtasks /end /tn $TaskName
            schtasks /delete /tn $TaskName
            schtasks /create /tn "${TaskName}" /xml ".\src\${TaskName}.xml"

        } else {
            
            Write-Host "      �^�X�N�̓o�^�͍s�킸�A�C���X�g�[���葱�����I�����܂��B"

        }

    } else {

        Write-Host "      �^�X�N�� ${TaskName} ��o�^���܂��B"
        schtasks /create /tn "${TaskName}" /xml ".\src\${TaskName}.xml"

    }

    Write-Host "===================================================================================="
    Write-Host "  ${TaskName} ���^�X�N�X�P�W���[���[�֓o�^���܂����B"

    if ($Windows11Edition -eq "Core") {

        Write-Host "  Windows 11 Home �����̃C���X�g�[���葱���������������܂����B"

    } else {

        Write-Host "  Windows 11 ${Windows11Edition} �������p�̏ꍇ�A�����������[�J���O���[�v�|���V�[���N�����A"
        Write-Host "  PowerShell �X�N���v�g��o�^���Ă��������B�o�^���@�� README.md �����Q�Ɗ肢�܂��B"

    }

    Write-Host "===================================================================================="

}

<#
    Record-WorkedTimeAndDesktopName ���A���C���X�g�[������B
#>
function Uninstall($TaskName, $Windows11Edition) {
    
    #===================================================================================
    # (1/2) �^�X�N�X�P�W���[���[����^�X�N Record-WorkedTimeAndDesktopName ���폜����B
    #===================================================================================
    Write-Host "=========================================================================="
    Write-Host "  ${TaskName} �̃A���C���X�g�[���葱�����J�n���܂��B"
    Write-Host "  �Ǘ��Ҍ�������Ŏ��s���Ă��������B"
    Write-Host "=========================================================================="
    Write-Host "(1/2) �^�X�N�X�P�W���[���[����^�X�N ${TaskName} ���폜"

    # �^�X�N�X�P�W���[���[�Ƀ^�X�N�� Record-WorkedTimeAndDesktopName ���o�^����Ă���ꍇ�̓^�X�N���폜����B
    if ((IsTaskRegistered $TaskName) -eq $true) {

        Write-Host "      �^�X�N�� ${TaskName} ���폜���܂��B"
        schtasks /end /tn $TaskName
        schtasks /delete /tn $TaskName

        if ((IsTaskRegistered $TaskName) -eq $true) {

            Write-Host "      �^�X�N�� ${TaskName} �̍폜�Ɏ��s���܂����B���萔�ł����A�蓮�ō폜���Ă��������B"
            Write-Host "      �A���C���X�g�[���葱�����I�����܂��B"

        }

    } else {

        Write-Host "      �^�X�N�� ${TaskName} �͓o�^����Ă��܂���ł����B"

    }

    Write-Host ""

    #=========================================================
    # (2/2) Virtual Desktop ���W���[�����A���C���X�g�[������B
    #=========================================================
    Write-Host "(2/2) Virtual Desktop ���W���[���̃A���C���X�g�[��"

    # Virtual Desktop ���C���X�g�[������Ă���ꍇ�A�����ł͉������Ȃ��B
    # Virtual Desktop ���C���X�g�[������Ă��Ȃ��ꍇ�A�C���X�g�[�����J�n����B
    if (IsVirtualDesktopModuleInstalled -eq $true) {

        Write-Host "      Virtual Desktop ���W���[�����A���C���X�g�[�����܂��B"
        Write-Host "      �A���C���X�g�[���������Ȃ��ꍇ�A n ����͂��ăA���C���X�g�[�����I�����Ă��������B"
        $UserInput = Read-Host "      Virtual Desktop ���W���[�����A���C���X�g�[�����܂����H [y/N]"

        if ($UserInput -eq "y"-or $UserInput -eq "Y") {

            Uninstall-Module VirtualDesktop

            if (IsVirtualDesktopModuleInstalled -eq $true) {

                # Virtual Desktop ���W���[�����A���C���X�g�[�����Ȃ������ꍇ�A�A���C���X�g�[���葱�����I������B
                Write-Host "      Virtual Desktop ���W���[�����A���C���X�g�[�����Ȃ��������A�������̓A���C���X�g�[���Ɏ��s���܂����B"
                Write-Host "      �A���C���X�g�[���葱�����I�����܂��B"
                return

            }
        } else {

            Write-Host "      Virtual Desktop ���W���[���̓A���C���X�g�[�����܂���B"

        }

    } else {

        Write-Host "      Virtual Desktop ���W���[���̓C���X�g�[������Ă��܂���ł����B"

    }

    Write-Host "=========================================================================="
    Write-Host "  ${TaskName} �̃A���C���X�g�[���葱�����������܂����B"
    Write-Host "=========================================================================="

}

#=============================================
# �C���X�g�[�� / �A���C���X�g�[�������s����B
#=============================================
$TaskName = "Record-WorkedTimeAndDesktopName"
$Windows11Edition = (Get-ComputerInfo).WindowsEditionID

switch ($ActionType) {
    "Install" { Install $TaskName $Windows11Edition }
    "Uninstall" { Uninstall $TaskName $Windows11Edition }
    default { Write-Host "��1�����ɂ� Install �� Uninstall ���w�肵�Ă��������B" }
}


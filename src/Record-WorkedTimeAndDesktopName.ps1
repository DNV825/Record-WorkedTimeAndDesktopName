<#
  .synopsis

  ���݂̎����ƃf�X�N�g�b�v�����擾���A�^�u��؂�e�L�X�g�Ƃ��ăt�@�C���֏������ށB

  .description

  �������񂾃^�u��؂�e�L�X�g�� Excel �ɓ\��t����ƁA�ȉ��̂悤�ȕ\���쐬�ł���B

    +----------------+----------+----------+----------+----------------------------+----------+
    |     �N����     | �J�n���� | �I������ | ��Ǝ��� | ��Ǝ�ʁi�f�X�N�g�b�v���j |   ����   |
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(�y) |  09:00   |  12:00   |   3.0 h  | Project A                  | <Start> �b
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(�y) |  13:00   |  15:00   |   2.0 h  | ����                       |          |
    +----------------+----------+----------+----------+----------------------------+----------+
    | 2025/04/12(�y) |  15:00   |  17:30   |   2.5 h  | ���������S��               | <Finish> |
    +----------------+----------+----------+----------+----------------------------+----------+
  
  �f�X�N�g�b�v������Ɠ��e�Ƃ��ė��p���邽�߁A���炩���ߕK�v�ȉ��z�f�X�N�g�b�v���쐬���Ă������ƁB
  ��Ǝ��Ԃ͖{�X�N���v�g�ŎZ�o���ďo�͂��邪�A Excel �Ōv�Z�����ق����ǂ���������Ȃ��B
  
  �Ȃ��A�{�X�N���v�g�̎��s�ɂ� VirtualDesktop ���W���[�����K�v�ƂȂ�B

  .parameter EventID

  �C�x���g�r���[�A�[����擾�����C�x���g ID �B

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

# �C�x���gID �ŕ�����s����悤�ɗ񋓑̂��`����B
# �Q�l�Fhttps://laboradian.com/win-system-power-state-events/
Enum EventIDs {
    FastShutdown = 187;  # �����V���b�g�_�E�� System/Kernel-Power:187
    Shutdown = 109;      # �V���b�g�_�E��     System/Kernel-Power:109
    PowerOn = 27;        # �p���[�I��         System/Kernel-Boot:27
    Hybernate = 109;     # �x�~���           System/Kernel-Power:109
}

# �v���W�F�N�g�t�H���_�z���� log �t�H���_�փ��O�t�@�C�����o�͂���B
# �ʂ̏ꏊ�ɒu�������ꍇ�͍D���ȃp�X���w�肷��΂悢�B
$LogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.log"
$DebugLogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\DebugRecord-WorkedTimeAndDesktopName.log"

# �L�^��������ƃf�X�N�g�b�v�����擾����B
# �킩��₷���̂��߁A���t����������ϐ������Ă����i���p�X�y�[�X�ŕ������A���t�������������o���B�j
$CurrentDateTime = Get-Date
$CurrentDateTimeFormatted = $CurrentDateTime.ToString("yyyy/MM/dd HH:mm")
$CurrentDate = (-split $CurrentDateTimeFormatted)[0]
$CurrentDesktopName = Get-DesktopName

# �J�n�E�I�����ɏ������ޖڈ���`����B
# �J�n�O�E�I����Ƀ^�C�}�[�^�X�N���Ă΂ꂽ�ꍇ�ɂ͖ڈ�Ƃ��ăA�b�v�f�[�g�}�[�N��t����B
$StartMark = "<Start>"
$FinishMark = "<Finish>"
$UpdateMark = "<Update-before-start-or-after-finish>"

# �f�o�b�O�p�֐��B
$IsDebugOn = $false
function Debug-Output($Path, $Value) {

    if ($IsDebugOn -eq $true) {

        Add-Content -Path $Path -Value $Value -Encoding UTF8

    }

}
 
#------------------------------------------------------------------------------
# �o�͐�t�@�C�������݂���ꍇ�͓��e���X�V���A
# ���݂��Ȃ��ꍇ�͏��߂ăX�N���v�g�𓮂������Ƃ݂Ȃ��ăt�@�C����V�K�쐬����B
#------------------------------------------------------------------------------
if ((Test-Path $LogFilePath) -eq $true) {
   
    #------------------------------------------------------------------------------
    # �o�͐�t�@�C���̍ŏI�s��ǂݎ���Đ��K�\���Ŋe���ڂɕ������A���e���擾����B
    # �ǂݎ�ꂽ���e�ɉ����ď������ݓ��e��ύX����B
    #------------------------------------------------------------------------------
    $IsMatched = (Get-Content -Tail 1 -Path $LogFilePath -Encoding UTF8) -match "^(?<Date>.+?)\t(?<StartedDateTime>.+?)\t(?<FinishedDateTime>.+?)\t(?<WorkedTime>.+?)\t(?<DesktopName>.*?)\t(?<StartFinishMark>.*?)$"
 
    #------------------------------------------------------------------------------
    # �ŏI�s���������������܂�Ă���ꍇ�A���K�\���ƈ�v����B
    # �ŏI�s�̊e���ڂ̒l���擾�ł��Ă���̂ŁA����𗘗p���ĕK�v�ȓ��e���������ށB
    #
    # �ŏI�s�̓��e�����K�\���ƈ�v���Ȃ��ꍇ�A�ŏI�s�͑z�肵���L�q�ɂȂ��Ă��Ȃ��B
    # ���̏ꍇ�A�d�����Ȃ��̂ł��̍s�͂�����߂ĐV�����s�ɊJ�n�̏���ǋL����B
    #
    # �o�͐�t�@�C���̃G���R�[�h�� UTF8NoBOM �ɂ������������A
    # �Â� PowerShell �� UTF8NoBOM ��Ή��Ȃ̂� BOM �t���� UTF8 �𗘗p����B
    #------------------------------------------------------------------------------
    if ($IsMatched -eq $true) {
 
        # �ŏI�s�ȊO�̍s���擾����B
        $Content = Get-Content -Path $LogFilePath | Select-Object -SkipLast 1 | Out-String
 
        # ��Ǝ��Ԃ��Z�o����B
        $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
        $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours �Ƃ�������B
 
        #----------------
        # �J�n���̏����B
        #----------------
        if ($EventID -eq [EventIDs]::PowerOn.Value__) {
            
            # �J�n���̏����͘A���ŌĂ΂�邱�Ƃ�����B
            # ���̂��߁A�擾�����ŏI�s�ɃX�^�[�g�}�[�N�����݂���ꍇ�͉������Ȃ��B
            if ($Matches['StartFinishMark'] -like "${StartMark}*") {
                
                # �������Ȃ��B
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n:1-0 PowerOn"

            }
            # �ŏI�s�ɃX�^�[�g�}�[�N�����݂��Ȃ��̂ł���΁A�J�n���̏����𑱍s����B
            else {

                # ���O�I���O�Ƀ^�X�N���Ă΂ꂽ�ꍇ�Ƀf�X�N�g�b�v�����擾�ł��Ȃ����Ƃ�����B
                # ���̏ꍇ�A�Ō�̃f�X�N�g�b�v���𗘗p���čs��ǉ�����B�ǉ������s�ɂ̓X�^�[�g�}�[�N��t�^����B
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t$($Matches['DesktopName'])`t${StartMark}:1-1 PowerOn"
 
                }
                # �f�X�N�g�b�v�����擾�ł����ꍇ�͍s��ǉ�����B
                else {
 
                    Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${StartMark}:1-2 PowerOn"
 
                }

            }
 
        }
        #----------------
        # �I�����̏����B
        #----------------
        elseif ($EventID -eq [EventIDs]::FastShutdown.Value__ -or
                $EventID -eq [EventIDs]::Shutdown.Value__ -or
                $EventID -eq [EventIDs]::Hybernate.Value__) {
 
            # �I�����̏����͘A���ŌĂ΂�邱�Ƃ�����B
            # ���̂��߁A�擾�����ŏI�s�Ƀt�B�j�b�V���}�[�N�����݂���ꍇ�͉������Ȃ��B
            if ($Matches['StartFinishMark'] -like "*${FinishMark}") {
                
                # �������Ȃ��B
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n:2-0 PowerOff"

            }
            # �ŏI�s�Ƀt�B�j�b�V���}�[�N�����݂��Ȃ��̂ł���΁A�I�����̏����𑱍s����B
            # �J�n������I�����܂œ����f�X�N�g�b�v�ō�Ƃ��Ă����ꍇ�A�X�^�[�g�}�[�N�ƃt�B�j�b�V���}�[�N�𗼕��������ތ`�Ƃ���B
            else {

                # ���O�I�t��Ƀ^�X�N���Ă΂ꂽ�ꍇ�Ƀf�X�N�g�b�v�����擾�ł��Ȃ����Ƃ�����B
                # ���̏ꍇ�A�Ō�̃f�X�N�g�b�v���𗘗p���čs���X�V����B�X�V�����s�ɂ̓t�B�j�b�V���}�[�N��t�^����B
                if ($CurrentDesktopName -eq $null -or
                    $CurrentDesktopName -eq "") {
 
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}:2-1 PowerOff"
 
                }
                # �f�X�N�g�b�v�����擾�ł����ꍇ�B�͍s��ǉ�����B
                else {
 
                    # �����f�X�N�g�b�v�����擾�o�����ꍇ�A��Ƃ��p�����Ă���Ƃ݂Ȃ��ē����s���X�V����B
                    # �X�V�����s�ɂ̓t�B�j�b�V���}�[�N��t�^����B
                    if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
                    
                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])${FinishMark}:2-2 PowerOff"
                        
                    }
                    # �قȂ�f�X�N�g�b�v�����擾�ł����ꍇ�A�Ō�ɕʂ̍�Ƃ��s�����Ɣ��f���čs��ǉ�����B
                    # �ǉ������s�ɂ̓t�B�j�b�V���}�[�N��t�^����B
                    else {

                        Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}" -NoNewline -Encoding UTF8
                        Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${FinishMark}:2-3 PowerOff"

                    }
 
                }

            }
 
        }
        #-----------------
        # 5 �����Ƃ̏����B
        #-----------------
        else {

            # ���O�I���O�A���O�I�t��ɂ��̃��[�g�ɓ���\��������B
            # ���̌Ăяo�����~�߂邱�Ƃ͂ł��Ȃ��̂ŁA�A�b�v�f�[�g�}�[�N��t���čs��ǉ�����B
            if ($Matches['StartFinishMark'] -like "*${FinishMark}") {
                
                Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${UpdateMark}" -NoNewline -Encoding UTF8
                Debug-Output -Path $DebugLogFilePath -Value "--`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t${UpdateMark}:3-0 Update"

            }
            # �t�B�j�b�V���}�[�N�����݂��Ȃ��̂ł���΁A�ʏ�� 5 �����Ƃ̏����𑱍s����B
            else {
 
                # �����f�X�N�g�b�v�����擾�o�����ꍇ�A��Ƃ��p�����Ă���Ƃ݂Ȃ��ē����s���X�V����B
                if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
                
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark']):3-1 Update"
                    
                }
                # �قȂ�f�X�N�g�b�v�����擾�ł����ꍇ�A�ʍ�Ƃ��J�n�����Ƃ݂Ȃ��Ď��̍s���J�n����B
                # ���̍ۂɃ}�[�N�͏����B
                else {
                    
                    Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
                    Debug-Output -Path $DebugLogFilePath -Value "--`r`n$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`t$($Matches['StartFinishMark'])`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:3-2 Update"

                }

            }
 
        }
    
    }
    # �ŏI�s�̋L�q���z��ʂ�łȂ��ꍇ�A"<Something wrong!>"�A���s�A�u�N�����v�A�u�J�n�����v�A�u��Ǝ��ԁv�A�u��Ǝ�ʁi�f�X�N�g�b�v���j�v���������ށB
    else {
    
        Add-Content -Path $LogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
        Debug-Output -Path $DebugLogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:5 something wrong" -Encoding UTF8
    
    }
}
# �o�͐�t�@�C�������݂��Ȃ��ꍇ�A�t�@�C����V�K�쐬���āu�N�����v�A�u�J�n�����v�A�u�I�������v�u��Ǝ��ԁv�A�u��Ǝ�ʁi�f�X�N�g�b�v���j�v���������ށB
else {
    
    Set-Content -Path $LogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t" -NoNewline -Encoding UTF8
    Debug-Output -Path $DebugLogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}`t:6 Create File" -Encoding UTF8
    
}

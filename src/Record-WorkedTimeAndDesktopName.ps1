<#
  .synopsis

  ���݂̎����ƃf�X�N�g�b�v�����擾���A�^�u��؂�e�L�X�g�Ƃ��ăt�@�C���֏������ށB

  .description

  �������񂾃^�u��؂�e�L�X�g�� Excel �ɓ\��t����ƁA�ȉ��̂悤�ȕ\���쐬�ł���B

    +----------------+----------+----------+----------+----------------------------+
    |     �N����     | �J�n���� | �I������ | ��Ǝ��� | ��Ǝ�ʁi�f�X�N�g�b�v���j |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(�y) |  09:00   |  12:00   |   3.0 h  | Project A                  |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(�y) |  13:00   |  15:00   |   2.0 h  | ����                       |
    +----------------+----------+----------+----------+----------------------------+
    | 2025/04/12(�y) |  15:00   |  17:30   |   2.5 h  | ���������S��               |
    +----------------+----------+----------+----------+----------------------------+
  
  �f�X�N�g�b�v������Ɠ��e�Ƃ��ė��p���邽�߁A���炩���ߕK�v�ȃf�X�N�g�b�v���쐬���Ă������ƁB
  �܂��A��Ǝ��Ԃ͖{�X�N���v�g�ł��o�͂��邪�A Excel �Ōv�Z�����ق����ǂ���������Ȃ��B
  
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
    FastShutdown = 187;
    Shutdown = 109;
    PowerOn = 27;
    Hybernate = 109;
}

# �v���W�F�N�g�t�H���_�z���� log �t�H���_�փ��O�t�@�C�����o�͂���B
# �ʂ̏ꏊ�ɒu�������ꍇ�͍D���ȃp�X���w�肷��΂悢�B
$LogFilePath = "$(Split-Path $PSCommandPath -Parent)\..\log\Record-WorkedTimeAndDesktopName.log"


# �L�^��������ƃf�X�N�g�b�v�����擾����B
# �킩��₷���̂��߁A���t����������ϐ������Ă����i���p�X�y�[�X�ŕ������A���t�������������o���B�j
$CurrentDateTime = Get-Date
$CurrentDateTimeFormatted = $CurrentDateTime.ToString("yyyy/MM/dd HH:mm")
$CurrentDate = (-split $CurrentDateTimeFormatted)[0]
$CurrentDesktopName = Get-DesktopName

#------------------------------------------------------------------------------
# �o�͐�t�@�C�������݂���ꍇ�͓��e���X�V���A
# ���݂��Ȃ��ꍇ�͏��߂ăX�N���v�g�𓮂������Ƃ݂Ȃ��ăt�@�C����V�K�쐬����B
#------------------------------------------------------------------------------
if ((Test-Path $LogFilePath) -eq $true) {
    #------------------------------------------------------------------------------
    # �o�͐�t�@�C���̍ŏI�s��ǂݎ���Đ��K�\���Ŋe���ڂɕ������A���e���擾����B
    # �ǂݎ�ꂽ���e�ɉ����ď������ݓ��e��ύX����B
    #------------------------------------------------------------------------------
    $IsMatched = (Get-Content -Tail 1 -Path $LogFilePath -Encoding UTF8) -match "^(?<Date>.+)\t(?<StartedDateTime>.+)\t(?<FinishedDateTime>.+)\t(?<WorkedTime>.+)\t(?<DesktopName>.+)$"

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

        if ($Matches['DesktopName'] -eq $CurrentDesktopName) {
            # �ŏI�s�̃f�X�N�g�b�v���ƌ��݂̃f�X�N�g�b�v�����������ꍇ�A��Ƃ��p�����Ă���Ƃ݂Ȃ��ďI�����������ݎ����֍X�V����B
            # �������A PC �N�����͍�Ƃ̌p���ł͂Ȃ��Ƃ݂Ȃ��A�ŏI�s���X�V�����Ɏ��̍s�֒ǋL����B
            if ($EventID -eq [EventIDs]::PowerOn.Value__) {

                Add-Content -Path $LogFilePath -Value "`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

            } else {
                $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
                $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours �Ƃ�������B

                Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])" -NoNewline -Encoding UTF8
            }
        } else {

            # �ŏI�s�̃f�X�N�g�b�v���ƌ��݂̃f�X�N�g�b�v�����������Ȃ��ꍇ�A�ʂ̍�Ƃ��J�n���Ă���Ƃ݂Ȃ��B
            # ���̂��߁A�ŏI�s�̏I�������Ɍ��ݎ������L�^���A���̍s�Ɍ��ݎ����ō�Ƃ��J�n�����|��ǋL����B
            $ElapsedHours = $CurrentDateTime - [DateTime]::ParseExact($Matches['StartedDateTime'], "yyyy/MM/dd HH:mm", $null)
            $WorkedTime = [String]::Format("{0:F1}", $ElapsedHours.TotalHours) # "{0:F1}" -f xx.TotalHours �Ƃ�������B

            Set-Content -Path $LogFilePath -Value "${Content}$($Matches['Date'])`t$($Matches['StartedDateTime'])`t${CurrentDateTimeFormatted}`t${WorkedTime}`t$($Matches['DesktopName'])`r`n$($Matches['Date'])`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

        }
    } else {

        # �ŏI�s�̋L�q���z��ʂ�łȂ��ꍇ�A"<Something wrong!>"�A���s�A�u�N�����v�A�u�J�n�����v�A�u��Ǝ��ԁv�A�u��Ǝ�ʁi�f�X�N�g�b�v���j�v���������ށB
        Add-Content -Path $LogFilePath -Value "<Something wrong!>`r`n${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

    }
} else {

    # �o�͐�t�@�C�������݂��Ȃ��ꍇ�A�t�@�C����V�K�쐬���āu�N�����v�A�u�J�n�����v�A�u�I�������v�u��Ǝ�ʁi�f�X�N�g�b�v���j�v���������ށB
    Set-Content -Path $LogFilePath -Value "${CurrentDate}`t${CurrentDateTimeFormatted}`t${CurrentDateTimeFormatted}`t0.0`t${CurrentDesktopName}" -NoNewline -Encoding UTF8

}
#
# �Ζ��\���珬����ʔ�������쐬����Powershell
# 
# ���s�`�� : .\createInvoice.ps1 �Ζ��\Excel�t�@�C��
#
# �Ζ��\�̌`�� : <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
#

# �������Ȃ��ꍇ�A�����𒆒f����
if ([string]::IsNullorEmpty($Args[0])) {
    Write-Host "`r`n====== �����ɋΖ��\�t�@�C�����w�肵�Ă������� =====`r`n"
    exit
}

# �Ζ��\�t�@�C���̃t�@�C�������猎�����o��
$Args[0] -match "_�Ζ��\_(?<month>.*?)��" | Out-Null
$month = $Matches.month


if ( $month -match "[1-9]|1[12]") {
    start-sleep -milliSeconds 300

    try {
    # �Ζ��\�t�@�C���̃t���p�X�擾
    $kinmuhyouFullPath = Resolve-Path $Args[0] -ErrorAction Stop
    } catch [Exception] {
        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        Write-Host "�Ζ��\�t�@�C�������݂��܂���B`r`n�����𒆒f���܂�`r`n"
        exit
    }

    Write-Host "`r`n#######################################"
    Write-Host (' ' + $month + " ���̏�����ʔ�������쐬���܂��B`r`n���΂炭���҂����������B")
    Write-Host "#######################################`r`n"
    break
}
else {
    # �Ζ��\�t�@�C���̃t�H�[�}�b�g���Ⴄ�ꍇ�͏C��������
    start-sleep -milliSeconds 300
    Write-Host " ######### <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx �̌`���Ƀt�@�C�������C�����Ă������� #########`r`n"
}

# Excel���N������
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

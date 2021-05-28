#
# �Ζ��\���珬����ʔ�������쐬����Powershell
# 
# �O����� : ���Ypowershell�Ɠ����t�H���_�ɋ�́u<�Ј��ԍ�>_������ʔ�E�o������Z���׏�_m��_<����>.xlsx�v��1���݂��邱��
#
# ���s�`�� : .\createInvoice.ps1 �Ζ��\Excel�t�@�C���@����Excel�t�@�C��
#
# �Ζ��\�̌`�� : <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
#

# ����������Ȃ��ꍇ�A�����𒆒f����
if ([string]::IsNullorEmpty($Args[0])) {
    Write-Host "`r`n====== ����1�ڂɋΖ��\�t�@�C�����w�肵�Ă������� ======`r`n"
    exit
} elseif ([string]::IsNullorEmpty($Args[1])) {
    Write-Host "`r`n====== ������2�ڏ�����ʔ�����t�@�C�����w�肵�Ă������� ======`r`n"
    exit
}

# �Ζ��\�t�@�C���̃t�@�C�������猎�������o��
$Args[0] -match "_�Ζ��\_(?<month>.*?)��" | Out-Null
$month = $Matches.month


if ( $Args[0]  -match "[0-9]{3}_�Ζ��\_[1-9]|1[12]��_.+" ) {
    start-sleep -milliSeconds 300

    try {
    # �Ζ��\�t�@�C���̃t���p�X�擾
    $kinmuhyouFullPath = Resolve-Path $Args[0] -ErrorAction Stop
    } catch [Exception] {
        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        Write-Host "�Ζ��\�t�@�C�������݂��܂���B`r`n�����𒆒f���܂�`r`n"
        exit
    }

    try {
    # �����t�@�C���̃t���p�X�擾
    $koguchiFullPath = Resolve-Path $Args[1] -ErrorAction Stop
    } catch [Exception] {
        # �����t�@�C�������݂��Ă��邩�`�F�b�N
        Write-Host "�����t�@�C�������݂��܂���B`r`n�����𒆒f���܂�`r`n"
        exit
    }

    Write-Host "`r`n#######################################"
    Write-Host (' ' + $month + " ���̏�����ʔ�������쐬���܂��B`r`n���΂炭���҂����������B")
    Write-Host "#######################################`r`n"
}else {
    # �Ζ��\�t�@�C���̃t�H�[�}�b�g���Ⴄ�ꍇ�͏C��������
    Write-Host " ######### <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx �̌`���Ƀt�@�C�������C�����Ă������� #########`r`n"
    exit
}

# Excel���N������
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)

# �����u�b�N���J��
$koguchiBook = $excel.workbooks.open($koguchiFullPath)
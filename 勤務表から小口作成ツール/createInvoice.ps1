#
# �Ζ��\���珬����ʔ�������쐬����Powershell
# 
# ���s�`�� : .\createInvoice.ps1 �Ζ��\Excel�t�@�C��
#
# �Ζ��\�̌`�� : <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
#

# �������Ȃ��ꍇ�A�����𒆒f����
if ([string]::IsNullorEmpty($Args[0])){
    Write-Host "`r`n�����ɋΖ��\�t�@�C�����w�肵�Ă�������`r`n"
    exit
}

while ($true) {
    # �Θb�`���Łu�����̏�����ʔ�������쐬���邩�H�v����͂��Ă��炤
    $month = $input | Read-Host "�����̏�����ʔ�������쐬���܂����H(���p�����݂̂����)"    

    if ($month -match "[1-9]|1[12]"){
        start-sleep -milliSeconds 300

        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        if (! (Test-Path ($Args[0].fullName))) {
            Write-Host "�Ζ��\�t�@�C�������݂��܂���B`r`n�����𒆒f���܂�"
            exit
        }

        Write-Host "`r`n#######################################"
        Write-Host (' ' + $month + " ���̏�����ʔ�������쐬���܂��B`r`n���΂炭���҂����������B")
        Write-Host "#######################################`r`n"
        break
    } else {
        # �z��̓��͂ł͂Ȃ��ꍇ�A������x���͂�������
        start-sleep -milliSeconds 300
        Write-Host "`r`n���p�����ł�����x���͂��Ă�������`r`n"
    }
}


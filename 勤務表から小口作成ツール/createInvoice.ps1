#
# �Ζ��\���珬����ʔ�������쐬����Powershell
# 
# �O����� : ���Ypowershell�Ɠ����t�H���_�Ƀt�H�[�}�b�g�ƃn���R���L�ڂ��ꂽ ������ʔ�E�o������Z���׏�Excel�t�@�C�� ��1���݂��邱��
#
# ���s�`�� : .\createInvoice.ps1 �Ζ��\Excel�t�@�C���@����Excel�t�@�C��
#
# �Ζ��\�̌`�� : <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx
#

# ----------------- �֐���` ---------------------

# �Ζ��\�Ə�����ۑ������ɕ��āAExcel�𒆒f����֐�
function endExcel {
    # Excel�̏I��
    $excel.quit()
    # �g�p���Ă����v���Z�X�̉��
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    $koguchiBook = $null
    $koguchiSheet = $null
    $koguchiCell = $null
    # �������I������
    exit
}

########### ���ӏ�����\���B���Ȃ��ꍇ�ɂ�Enter����������B

# ����������Ȃ��ꍇ�A�����𒆒f����
if ([string]::IsNullorEmpty($Args[0])) {
    Write-Host "`r`n====== ����1�ڂɏ�����ʔ�����t�@�C�����w�肵�Ă������� ======`r`n" -ForegroundColor Red
    exit
}

# ���ݓ������擾����
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# ���ݓ�������쐬����ׂ��Ζ��\�̌����𔻒�
if ($today -le 24) {
    $kinmuhyouMonth = $thisMonth -1
} else {
    $kinmuhyouMonth = $thisMonth
}

# �Ζ��\�t�@�C�����擾
$kinmuhyou = Get-ChildItem -Recurse -File |? name -Match "[0-9]{3}_�Ζ��\_($kinmuhyouMonth)��_.+"

# �Y���Ζ��\�t�@�C���̌��m�F
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�������݂��܂���`r`n" -ForegroundColor Red
} elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n�Y������Ζ��\�t�@�C�����������܂�`r`n" -ForegroundColor Red
}

# �Ζ��\�t�@�C���̃t�@�C�������猎�������o��
$kinmuhyou -Match "_�Ζ��\_(?<month>.*?)��" | Out-Null
$month = $Matches.month


if ( $kinmuhyou.Name  -match "[0-9]{3}_�Ζ��\_[1-9]|1[12]��_.+" ) {
    Start-Sleep -milliSeconds 300

    try {
    # �Ζ��\�t�@�C���̃t���p�X�擾
    $kinmuhyouFullPath = Resolve-Path $kinmuhyou -ErrorAction Stop
    } catch [Exception] {
        # �Ζ��\�����݂��Ă��邩�`�F�b�N
        Write-Host "�Ζ��\�t�@�C�������݂��܂���B`r`n�����𒆒f���܂�`r`n" -ForegroundColor Red
        exit
    }

    try {
    # �����t�@�C���̃t���p�X�擾
    $koguchiFullPath = Resolve-Path $Args[0] -ErrorAction Stop
    } catch [Exception] {
        # �����t�@�C�������݂��Ă��邩�`�F�b�N
        Write-Host "�����t�@�C�������݂��܂���B`r`n�����𒆒f���܂�`r`n" -ForegroundColor Red
        exit
    }

    Write-Host "`r`n#######################################"
    Write-Host (' �@' + $month + " ���̏�����ʔ�������쐬���܂��B`r`n�@�@���΂炭���҂����������B")
    Write-Host "#######################################`r`n"
}else {
    # �Ζ��\�t�@�C���̃t�H�[�}�b�g���Ⴄ�ꍇ�͏C��������
    Write-Host " ######### <�Ј��ԍ�>_�Ζ��\_m��_<����>.xlsx �̌`���Ƀt�@�C�������C�����Ă������� #########`r`n" -ForegroundColor Red
    exit
}

# Excel���N������
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excel�����b�Z�[�W�_�C�A���O��\�����Ȃ��悤�ɂ���
$excel.DisplayAlerts = $false
$excel.visible = $true

# �Ζ��\�u�b�N���J��
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.sheets( "$month"+'��')

# �����u�b�N���J��
$koguchiBook = $excel.workbooks.open($koguchiFullPath)
$koguchiSheet = $koguchiBook.sheets(1)


# ------------- �Ζ��\�̒��g�������ɃR�s�[���� ----------------

# ------------- �l��񗓂̃R�s�[ --------------

# ���݂̔N���擾
$thisYear = (Get-Date).Year
# 1����12���̏�������낤�Ƃ��Ă�����N����N�߂�
if ($month -eq 1 -and (Get-Date).day -le 24) {
    $thisYear = (Get-Date).AddYears(-1).Year
}

# 1. �N�����̃R�s�[
$koguchiSheet.cells.item(60,4) = $thisYear
$koguchiSheet.cells.item(60,8) = $month

# ���̍ŏI������t���ɐݒ�
$koguchiSheet.cells.item(60,11) = (Get-Date "$thisYear/$month/1" -Day 1).AddMonths(1).AddDays(-1).Day

# 2. ���O�̃R�s�[
$koguchiSheet.cells.item(64,21) = $kinmuhyouSheet.cells.range("W7").text
# �Ζ��\�̖��O���󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(64,21).text -eq "") {
    Write-Host ("`r`n" + $month + "���̋Ζ��\�ɖ��O���L�ڂ���Ă��܂���`r`n�����𒆒f���܂�`r`n") -ForegroundColor Red
    endExcel
}

# 3. �����̃R�s�[
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "��" ���폜����
$affiliation -match "(?<affliationName>.+?)��" | Out-Null
$koguchiSheet.cells.item(62,6) = $Matches.affliationName
# �Ζ��\�̏������󔒂������ꍇ�����𒆒f����
if ($koguchiSheet.cells.item(62,6).text -eq "") {
    Write-Host ("`r`n" + $month + "���̋Ζ��\�ɏ������L�ڂ���Ă��܂���`r`n�����𒆒f���܂�`r`n") -ForegroundColor Red
    endExcel
}

# 4. ��ӂ̃R�s�[
# ��ӂ��Ȃ���������Ȃ��t���O
$haveNotStamp = $false
# �Ζ��\�̈�ӂ̂���Z�����N���b�v�{�[�h�ɃR�s�[
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# �����V�[�g�Ɉ�ӂ��y�[�X�g
$koguchiCell=$koguchiSheet.range("AD64")
$koguchiSheet.paste($koguchiCell)
# �y�[�X�g���ҏW
$koguchiSheet.range("AD64").formula = ""
$koguchiSheet.range("AD64").interior.colorindex = 0
# �r����ҏW���邽�߂̐錾
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
$koguchiSheet.range("AD64").borders.linestyle = $linestyle::xllinestylenone
if ($koguchiSheet.shapes.count -eq 68) {
    $haveNotStamp = $true
}



# �����F�̕ύX�i�S�����Ɂj

# ��ӂ��Ȃ���������Ȃ��ꍇ���ӊ��N
if ($haveNotStamp) {
    Write-Host "`r`n#################################################################################" -ForegroundColor Blue
    Write-Host "�@�@��ӂ��Ζ��\, ������ʔ�E�o������Z���׏��ɓ����Ă��Ȃ��\��������܂�`r`n�@�@�m�F���Ă�������" -ForegroundColor Blue
    Write-Host "#################################################################################`r`n" -ForegroundColor Blue
}
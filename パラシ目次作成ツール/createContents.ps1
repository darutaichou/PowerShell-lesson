#
# �p���V�̖ڎ����쐬����Powershell
#

# �p�C�v���C������p���V�������󂯎��
$parameterSheets = $input |? Name -Match '�p���V���ǂ�'

# Excel���N��
try {
    # �N������Excel�v���Z�X���擾
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

$excel.visible = $true

$parameterSheets |% {
    # �p���V�̃t���p�X���擾
    $fullPath = $_.fullName

    # �u�b�N���J��
    $book = $excel.workbooks.open($fullPath)

    # �ڎ��V�[�g���쐬
    $book.worksheets.add($book.sheets(3)) | Out-Null

    # $contentsSheet : �ڎ��V�[�g
    $contentsSheet = $book.sheets(3)

    # �ڎ��V�[�g�̖��O�ύX
    $contentsSheet.name = '�ڎ�'

    # �匩�o���̍s�𑾎��ɂ���
    $contentsSheet.cells.range("B1:B1000").font.bold = $true
    
    # �ڎ��V�[�g�̏c��J�E���^�[
    $countContentsRow = 2

    for ($i = 4; $i -le $book.worksheets.count; $i++) {
        # �匩�o����ڎ��ɃR�s�[
        $contentsSheet.cells.item($countContentsRow, 2) = $book.sheets($i).cells.item(2, 2)
        Write-Host ("�匩�o����" + $countContentsRow + "�ɃR�s�[���܂���")
        $countContentsRow++ 

        for ($j = 1; $j -le 1000; $j++){
            if ($book.sheets($i).cells.item($j,3).text -match "^[0-9]{1,2}-[0-9]{1,2}"){
                $contentsSheet.cells.item($countContentsRow,3) = $book.sheets($i).cells.item($j,3)
                Write-Host ("�����o����" + $countContentsRow + "�ɃR�s�[���܂���")
                $countContentsRow++
            }    
        }
    }

    # �u�b�N�̏㏑���ۑ�
    $book.save()

    # �u�b�N�����
    $book.close()
}

# Excel�����
$excel.quit()

# �ϐ����J������
$excel = $null
$book = $null
$contentsSheet = $null
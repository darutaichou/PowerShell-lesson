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

$parameterSheets |% {
    # �p���V�̃t���p�X���擾
    $fullPath = $_.fullName

    # �u�b�N���J��
    $book = $excel.workbooks.open($fullPath)

    # �ڎ��V�[�g���쐬
    
    # �ڎ��V�[�g�̏c��J�E���^�[
    $countContentsRow = 2

    for ($i = 4; $i -le $book.worksheets.count; $i++) {
        # �匩�o����ڎ��ɃR�s�[
        $parameterSheet.sheet(3).cells.item(2, 2) = $parameterSheet.sheet($i).cells.item
        $countContentsRow++ 
    }
}
# Excel���J��
$excel = New-Object -ComObject Excel.Application

# Excel��������悤�ɂ���
$excel.visible = $true

# �Ζ��\���J��
# $kinmuhyou = gci -Recurse |? name -cmatch ("^[0-9]{2,3}_�Ζ��\_")
# $kinmuhyou
$kinmuhyouBook = $excel.workbooks.open("D:\106_PowerShell\BVS\PowerShell-lesson\����ǂڂ�����\116_�Ζ��\_5��_�u��.xlsx")
$kinmuhyouSheet = $kinmuhyouBook.sheets('5��')

$allShapes = $kinmuhyouSheet.Shapes

foreach ($shape in $allShapes) {
    $shape.placement = 2
}
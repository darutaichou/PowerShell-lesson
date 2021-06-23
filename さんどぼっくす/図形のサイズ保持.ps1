# Excelを開く
$excel = New-Object -ComObject Excel.Application

# Excelを見えるようにする
$excel.visible = $true

# 勤務表を開く
# $kinmuhyou = gci -Recurse |? name -cmatch ("^[0-9]{2,3}_勤務表_")
# $kinmuhyou
$kinmuhyouBook = $excel.workbooks.open("D:\106_PowerShell\BVS\PowerShell-lesson\さんどぼっくす\116_勤務表_5月_志村.xlsx")
$kinmuhyouSheet = $kinmuhyouBook.sheets('5月')

$allShapes = $kinmuhyouSheet.Shapes

foreach ($shape in $allShapes) {
    $shape.placement = 2
}
#
# パラシの目次を作成するPowershell
#

# パイプラインからパラシだけを受け取る
$parameterSheets = $input |? Name -Match 'パラシもどき'

# Excelを起動
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

$parameterSheets |% {
    # パラシのフルパスを取得
    $fullPath = $_.fullName

    # ブックを開く
    $book = $excel.workbooks.open($fullPath)

    # 目次シートを作成
    
    # 目次シートの縦列カウンター
    $countContentsRow = 2

    for ($i = 4; $i -le $book.worksheets.count; $i++) {
        # 大見出しを目次にコピー
        $parameterSheet.sheet(3).cells.item(2, 2) = $parameterSheet.sheet($i).cells.item
        $countContentsRow++ 
    }
}
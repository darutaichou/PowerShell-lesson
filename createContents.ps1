#
# パラシから目次を作成するPowerShell作成
#

# パイプラインからパラシのみを受け取る
$parameterSheets = $input |? Name -Match 'パラシもどき'

# Excelの起動
try {
    # 起動中であれば起動中のプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

$parameterSheets |% {
    # ブックを開く
    $book = $excel.workbooks.open($_)

    # 目次シートの作成
    
    # 目次シートの縦列のカウンター
    $countContentsRow = 2

    for ($i = 4; $i -le $book.worksheets.count; $i++) {
        # 大見出しを目次へコピー
        $parameterSheet.sheet(3).cells.item(2, 2) = $parameterSheet.sheet(i).cells.item 
    }
}
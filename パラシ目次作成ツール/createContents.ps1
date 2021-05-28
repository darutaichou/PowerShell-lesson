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

$excel.visible = $true

$parameterSheets |% {
    # パラシのフルパスを取得
    $fullPath = $_.fullName

    # ブックを開く
    $book = $excel.workbooks.open($fullPath)

    # 目次シートを作成
    $book.worksheets.add($book.sheets(3)) | Out-Null

    # $contentsSheet : 目次シート
    $contentsSheet = $book.sheets(3)

    # 目次シートの名前変更
    $contentsSheet.name = '目次'

    # 大見出しの行を太字にする
    $contentsSheet.cells.range("B1:B1000").font.bold = $true
    
    # 目次シートの縦列カウンター
    $countContentsRow = 2

    for ($i = 4; $i -le $book.worksheets.count; $i++) {
        # 大見出しを目次にコピー
        $contentsSheet.cells.item($countContentsRow, 2) = $book.sheets($i).cells.item(2, 2)
        Write-Host ("大見出しを" + $countContentsRow + "にコピーしました")
        $countContentsRow++ 

        for ($j = 1; $j -le 1000; $j++){
            if ($book.sheets($i).cells.item($j,3).text -match "^[0-9]{1,2}-[0-9]{1,2}"){
                $contentsSheet.cells.item($countContentsRow,3) = $book.sheets($i).cells.item($j,3)
                Write-Host ("小見出しを" + $countContentsRow + "にコピーしました")
                $countContentsRow++
            }    
        }
    }

    # ブックの上書き保存
    $book.save()

    # ブックを閉じる
    $book.close()
}

# Excelを閉じる
$excel.quit()

# 変数を開放する
$excel = $null
$book = $null
$contentsSheet = $null
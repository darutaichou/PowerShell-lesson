#
# 勤務表から小口交通費請求書を作成するPowershell
# 
# 前提条件 : 当該powershellと同じフォルダに空の「<社員番号>_小口交通費・出張旅費精算明細書_m月_<氏名>.xlsx」が1つ存在すること
#
# 実行形式 : .\createInvoice.ps1 勤務表Excelファイル　小口Excelファイル
#
# 勤務表の形式 : <社員番号>_勤務表_m月_<氏名>.xlsx
#

# 引数が足りない場合、処理を中断する
if ([string]::IsNullorEmpty($Args[0])) {
    Write-Host "`r`n====== 引数1個目に小口交通費請求書ファイルを指定してください ======`r`n"
    exit
}

# 現在日時を取得する
$tomonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
if ($today -le 24) {
    $kinmuhyouMonth = $tomonth
} else {
    $kinmuhyouMonth = $tomonth + 1
}

# 勤務表ファイルのファイル名から月次を取り出す
$Args[0] -match "_勤務表_(?<month>.*?)月" | Out-Null
$month = $Matches.month


if ( $Args[0]  -match "[0-9]{3}_勤務表_[1-9]|1[12]月_.+" ) {
    start-sleep -milliSeconds 300

    try {
    # 勤務表ファイルのフルパス取得
    $kinmuhyouFullPath = Resolve-Path $Args[0] -ErrorAction Stop
    } catch [Exception] {
        # 勤務表が存在しているかチェック
        Write-Host "勤務表ファイルが存在しません。`r`n処理を中断します`r`n"
        exit
    }

    try {
    # 小口ファイルのフルパス取得
    $koguchiFullPath = Resolve-Path $Args[1] -ErrorAction Stop
    } catch [Exception] {
        # 小口ファイルが存在しているかチェック
        Write-Host "小口ファイルが存在しません。`r`n処理を中断します`r`n"
        exit
    }

    Write-Host "`r`n#######################################"
    Write-Host (' ' + $month + " 月の小口交通費請求書を作成します。`r`nしばらくお待ちください。")
    Write-Host "#######################################`r`n"
}else {
    # 勤務表ファイルのフォーマットが違う場合は修正させる
    Write-Host " ######### <社員番号>_勤務表_m月_<氏名>.xlsx の形式にファイル名を修正してください #########`r`n"
    exit
}

# Excelを起動する
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)

# 小口ブックを開く
$koguchiBook = $excel.workbooks.open($koguchiFullPath)
#
# 勤務表から小口交通費請求書を作成するPowershell
# 
# 前提条件 : 当該powershellと同じフォルダにフォーマットとハンコが記載された 小口交通費・出張旅費精算明細書Excelファイル が1つ存在すること
#
# 実行形式 : .\createInvoice.ps1 勤務表Excelファイル　小口Excelファイル
#
# 勤務表の形式 : <社員番号>_勤務表_m月_<氏名>.xlsx
#

# ----------------- 関数定義 ---------------------

# 勤務表と小口を保存せずに閉じて、Excelを中断する関数
function endExcel {
    # Excelの終了
    $excel.quit()
    # 使用していたプロセスの解放
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    $koguchiBook = $null
    $koguchiSheet = $null
    $koguchiCell = $null
    # 処理を終了する
    exit
}

###########
########################## 注意書きを表示。問題ない場合にはEnterを押させる。
#=========================================================================

# 引数が足りない場合、処理を中断する
if ([string]::IsNullorEmpty($Args[0])) {
    Write-Host "`r`n====== 引数1個目に小口交通費請求書ファイルを指定してください ======`r`n" -ForegroundColor Red
    exit
}

# 現在日時を取得する
$thisMonth = (Get-Date).Month
$today = (Get-Date).Day

# 現在日時から作成するべき勤務表の月次を判定
if ($today -le 24) {
    $kinmuhyouMonth = $thisMonth -1
} else {
    $kinmuhyouMonth = $thisMonth
}

# 勤務表ファイルを取得
$kinmuhyou = Get-ChildItem -Recurse -File |? name -Match "[0-9]{3}_勤務表_($kinmuhyouMonth)月_.+"

# 該当勤務表ファイルの個数確認
if ($kinmuhyou.Count -lt 1) {
    Write-Host "`r`n該当する勤務表ファイルが存在しません`r`n" -ForegroundColor Red
} elseif ($kinmuhyou.Count -gt 1) {
    Write-Host "`r`n該当する勤務表ファイルが多すぎます`r`n" -ForegroundColor Red
}

# 勤務表ファイルのファイル名から月次を取り出す
$kinmuhyou -Match "_勤務表_(?<month>.*?)月" | Out-Null
$month = $Matches.month


if ( $kinmuhyou.Name  -match "[0-9]{3}_勤務表_[1-9]|1[12]月_.+" ) {
    Start-Sleep -milliSeconds 300

    try {
    # 勤務表ファイルのフルパス取得
    $kinmuhyouFullPath = Resolve-Path $kinmuhyou -ErrorAction Stop
    } catch [Exception] {
        # 勤務表が存在しているかチェック
        Write-Host "勤務表ファイルが存在しません。`r`n処理を中断します`r`n" -ForegroundColor Red
        exit
    }

    try {
    # 小口ファイルのフルパス取得
    $koguchiFullPath = Resolve-Path $Args[0] -ErrorAction Stop
    } catch [Exception] {
        # 小口ファイルが存在しているかチェック
        Write-Host "小口ファイルが存在しません。`r`n処理を中断します`r`n" -ForegroundColor Red
        exit
    }

    Write-Host "`r`n#######################################"
    Write-Host (' 　' + $month + " 月の小口交通費請求書を作成します。`r`n　　しばらくお待ちください。")
    Write-Host "#######################################`r`n"
}else {
    # 勤務表ファイルのフォーマットが違う場合は修正させる
    Write-Host " ######### <社員番号>_勤務表_m月_<氏名>.xlsx の形式にファイル名を修正してください #########`r`n" -ForegroundColor Red
    exit
}

# Excelを起動する
try {
    # 起動中のExcelプロセスを取得
    $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    $excel = New-Object -ComObject "Excel.Application" 
}

# Excelがメッセージダイアログを表示しないようにする
$excel.DisplayAlerts = $false
$excel.visible = $true

# 勤務表ブックを開く
$kinmuhyouBook = $excel.workbooks.open($kinmuhyouFullPath)
$kinmuhyouSheet = $kinmuhyouBook.sheets( "$month"+'月')

# 小口ブックを開く
$koguchiBook = $excel.workbooks.open($koguchiFullPath)
$koguchiSheet = $koguchiBook.sheets(1)


# ------------- 勤務表の中身を小口にコピーする ----------------

# ------------- 個人情報欄のコピー --------------

# 現在の年を取得
$thisYear = (Get-Date).Year
# 1月に12月の小口を作ろうとしていたら年を一年戻す
if ($month -eq 1 -and (Get-Date).day -le 24) {
    $thisYear = (Get-Date).AddYears(-1).Year
}

# 1. 年月日のコピー
$koguchiSheet.cells.item(60,4) = $thisYear
$koguchiSheet.cells.item(60,8) = $month

# 月の最終日を日付欄に設定
$koguchiSheet.cells.item(60,11) = (Get-Date "$thisYear/$month/1").AddMonths(1).AddDays(-1).Day

# 2. 名前のコピー
$koguchiSheet.cells.item(64,21) = $kinmuhyouSheet.cells.range("W7").text
# 勤務表の名前が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(64,21).text -eq "") {
    Write-Host ("`r`n" + $month + "月の勤務表に名前が記載されていません`r`n処理を中断します`r`n") -ForegroundColor Red
    endExcel
}

# 3. 所属のコピー
$affiliation = $kinmuhyouSheet.cells.range("W6").text
# "部" を削除する
$affiliation -match "(?<affliationName>.+?)部" | Out-Null
$koguchiSheet.cells.item(62,6) = $Matches.affliationName
# 勤務表の所属が空白だった場合処理を中断する
if ($koguchiSheet.cells.item(62,6).text -eq "") {
    Write-Host ("`r`n" + $month + "月の勤務表に所属が記載されていません`r`n処理を中断します`r`n") -ForegroundColor Red
    endExcel
}

# 4. 印鑑のコピー
# 印鑑がないかもしれないフラグ
$haveNotStamp = $false
# 勤務表の印鑑のあるセルをクリップボードにコピー
$kinmuhyouSheet.range("AA7").copy() | Out-Null
# 小口シートに印鑑をペースト
$koguchiCell=$koguchiSheet.range("AD64")
$koguchiSheet.paste($koguchiCell)
# ペースト先を編集
$koguchiSheet.range("AD64").formula = ""
$koguchiSheet.range("AD64").interior.colorindex = 0
# 罫線を編集するための宣言
$LineStyle = "microsoft.office.interop.excel.xlLineStyle" -as [type]
# 罫線をなしにする
$koguchiSheet.range("AD64").borders.linestyle = $linestyle::xllinestylenone
# 印鑑（オブジェクト）が増えてなさそうなら、メッセージを表示する
if ($koguchiSheet.shapes.count -eq 68) {
    $haveNotStamp = $true
}



# 文字色の変更（全部黒に）

# 印鑑がないかもしれない場合注意喚起
if ($haveNotStamp) {
    Write-Host "`r`n#################################################################################" -ForegroundColor Blue
    Write-Host "　　印鑑が勤務表に入っていない、または既定のセルからずれている可能性があります`r`n　　確認してください"  -ForegroundColor Blue
    Write-Host "#################################################################################`r`n" -ForegroundColor Blue
}
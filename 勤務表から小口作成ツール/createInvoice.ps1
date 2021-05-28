#
# 勤務表から小口交通費請求書を作成するPowershell
# 
# 実行形式 : .\createInvoice.ps1 勤務表Excelファイル
#
# 勤務表の形式 : <社員番号>_勤務表_m月_<氏名>.xlsx
#

# 引数がない場合、処理を中断する
if ([string]::IsNullorEmpty($Args[0])){
    Write-Host "`r`n引数に勤務表ファイルを指定してください`r`n"
    exit
}

while ($true) {
    # 対話形式で「何月の小口交通費請求書を作成するか？」を入力してもらう
    $month = $input | Read-Host "何月の小口交通費請求書を作成しますか？(半角数字のみを入力)"    

    if ($month -match "[1-9]|1[12]"){
        start-sleep -milliSeconds 300

        # 勤務表が存在しているかチェック
        if (! (Test-Path ($Args[0].fullName))) {
            Write-Host "勤務表ファイルが存在しません。`r`n処理を中断します"
            exit
        }

        Write-Host "`r`n#######################################"
        Write-Host (' ' + $month + " 月の小口交通費請求書を作成します。`r`nしばらくお待ちください。")
        Write-Host "#######################################`r`n"
        break
    } else {
        # 想定の入力ではない場合、もう一度入力をさせる
        start-sleep -milliSeconds 300
        Write-Host "`r`n半角数字でもう一度入力してください`r`n"
    }
}


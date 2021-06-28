#
# メインのユーザ入力
#

# 勤務表を保存せずに閉じて、Excelを中断する関数
function breakExcel {
    # Bookを閉じる
    $kinmuhyouBook.close()
    # 使用していたプロセスの解放
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    # ガベージコレクト
    [GC]::Collect()
    # 処理を終了する
    exit
}


# フォームを作成する関数
# Args[0] : タイトルに表示する文字列
function drawForm {
    . {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "勤務地の情報を登録"
        $form.Size = New-Object System.Drawing.Size(650, 730)
        $form.StartPosition = "CenterScreen"
        $form.font = $font
        $form.formborderstyle = "FixedSingle"
    } | Out-Null
    return $form
}

# ラベルを作成する関数
# Args[0] : フォーム内の設定座標（横の位置）
# Args[1] : フォーム内の設定座標（縦の位置。高さ）
# Args[2] : ラベルを表示する幅
# Args[3] : ラベルを表示する高さ
# Args[4] : ラベルに表示する文字列
# Args[5] : ラベルを表示するフォーム
# Args[6] : ラベルのフォント
function drawLabel {
    . {
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
        $label.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
        $label.Text = $Args[4]
        $label.forecolor = "black"
        $label.font = $Args[6]
        $Args[5].Controls.Add($label)
    } | Out-Null
    return 
}

# OK/登録ボタンを作成する関数
# Args[0] : フォーム内の設定座標（横の位置）
# Args[1] : フォーム内の設定座標（縦の位置。高さ）
# Args[2] : ボタンを表示する横幅
# Args[3] : ボタンを表示する縦幅
# Args[4] : OK/登録ボタンに表示する文字列
# Args[5] : OK/登録ボタンを表示するフォーム
# result : OK
function drawOKButton {
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
    $OKButton.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
    $OKButton.Text = $Args[4]
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Args[5].AcceptButton = $OKButton
    $Args[5].Controls.Add($OKButton)
}

# 在宅ボタンを作成する関数
# Args[0] : フォーム内の設定座標（縦の位置。高さ）
# Args[1] : 在宅ボタンに表示する文字列
# Args[2] : 在宅ボタンを表示するフォーム
# result : Yes
function drawAtHomeButton {
    $AtHomeButton = New-Object System.Windows.Forms.Button
    $AtHomeButton.Location = New-Object System.Drawing.Point(10, $Args[0])
    $AtHomeButton.Size = New-Object System.Drawing.Size(300, 30)
    $AtHomeButton.Text = $Args[1]
    $AtHomeButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $AtHomeButton.Backcolor = "paleturquoise"
    $AtHomeButton.Forecolor = "Blue"
    $Args[2].Controls.Add($AtHomeButton)
}

# 戻るボタンを作成する関数
# Args[0] : フォーム内の設定座標（縦の位置。高さ）
# Args[1] : 戻るボタンに表示する文字列
# Args[2] : 戻るボタンを表示するフォーム
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(500, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(90, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    # 1番目のフォームではボタンを非活性にする
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }
    else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# 登録済み勤務地から選択ボタンを作成する関数
# result : No
function drawRegisteredButton {
    $registeredButton = New-Object System.Windows.Forms.Button
    $registeredButton.Location = New-Object System.Drawing.Point(320, $Args[0])
    $registeredButton.Size = New-Object System.Drawing.Size(300, 30)
    $registeredButton.Text = $Args[1]
    $registeredButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $registeredButton.Backcolor = "palegreen"
    $registeredButton.Forecolor = "darkgreen"
    # ツール用引数.txt が存在していない or 中身が空の時はボタンを非活性にする
    if (!(Test-Path $infoTextFileFullpath) -or ($argumentText.Length -eq 0)) {
        $registeredButton.Enabled = $false; 
    }
    else {
        $registeredButton.Enabled = $True;
    }
    $Args[2].Controls.Add($registeredButton)
}


# テキストボックスを作成する関数
# Args[0] : フォーム内の設定座標（横の位置）
# Args[1] : フォーム内の設定座標（縦の位置。高さ）
# Args[2] : テキストボックスの横幅
# Args[3] : テキストボックスの高さ
# Args[4] : テキストボックスを表示するフォーム
function drawTextBox {
    . {
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point($Args[0], $Args[1])
        $textBox.Size = New-Object System.Drawing.Size($Args[2], $Args[3])
        $textBox.BackColor = "white"
        $Args[4].Controls.Add($textBox)
    } | Out-Null
    return $textBox
}


#
# ���C���̃��[�U����
#

# �Ζ��\��ۑ������ɕ��āAExcel�𒆒f����֐�
function breakExcel {
    # Book�����
    $kinmuhyouBook.close()
    # �g�p���Ă����v���Z�X�̉��
    $excel = $null
    $kinmuhyouBook = $null
    $kinmuhyouSheet = $null
    # �K�x�[�W�R���N�g
    [GC]::Collect()
    # �������I������
    exit
}


# �t�H�[�����쐬����֐�
# Args[0] : �^�C�g���ɕ\�����镶����
function drawForm {
    . {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "�Ζ��n�̏���o�^"
        $form.Size = New-Object System.Drawing.Size(650, 730)
        $form.StartPosition = "CenterScreen"
        $form.font = $font
        $form.formborderstyle = "FixedSingle"
    } | Out-Null
    return $form
}

# ���x�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : ���x����\�����镝
# Args[3] : ���x����\�����鍂��
# Args[4] : ���x���ɕ\�����镶����
# Args[5] : ���x����\������t�H�[��
# Args[6] : ���x���̃t�H���g
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

# OK/�o�^�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : �{�^����\�����鉡��
# Args[3] : �{�^����\������c��
# Args[4] : OK/�o�^�{�^���ɕ\�����镶����
# Args[5] : OK/�o�^�{�^����\������t�H�[��
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

# �ݑ�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �ݑ�{�^���ɕ\�����镶����
# Args[2] : �ݑ�{�^����\������t�H�[��
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

# �߂�{�^�����쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[1] : �߂�{�^���ɕ\�����镶����
# Args[2] : �߂�{�^����\������t�H�[��
# result : Retry
function drawReturnButton {
    $ReturnButton = New-Object System.Windows.Forms.Button
    $ReturnButton.Location = New-Object System.Drawing.Point(500, $Args[0])
    $ReturnButton.Size = New-Object System.Drawing.Size(90, 30)
    $ReturnButton.Text = $Args[1]
    $ReturnButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    # 1�Ԗڂ̃t�H�[���ł̓{�^����񊈐��ɂ���
    if ($i -eq 0) {
        $ReturnButton.Enabled = $false; 
    }
    else {
        $ReturnButton.Enabled = $True;
    }
    $Args[2].Controls.Add($ReturnButton)
}

# �o�^�ς݋Ζ��n����I���{�^�����쐬����֐�
# result : No
function drawRegisteredButton {
    $registeredButton = New-Object System.Windows.Forms.Button
    $registeredButton.Location = New-Object System.Drawing.Point(320, $Args[0])
    $registeredButton.Size = New-Object System.Drawing.Size(300, 30)
    $registeredButton.Text = $Args[1]
    $registeredButton.DialogResult = [System.Windows.Forms.DialogResult]::No
    $registeredButton.Backcolor = "palegreen"
    $registeredButton.Forecolor = "darkgreen"
    # �c�[���p����.txt �����݂��Ă��Ȃ� or ���g����̎��̓{�^����񊈐��ɂ���
    if (!(Test-Path $infoTextFileFullpath) -or ($argumentText.Length -eq 0)) {
        $registeredButton.Enabled = $false; 
    }
    else {
        $registeredButton.Enabled = $True;
    }
    $Args[2].Controls.Add($registeredButton)
}


# �e�L�X�g�{�b�N�X���쐬����֐�
# Args[0] : �t�H�[�����̐ݒ���W�i���̈ʒu�j
# Args[1] : �t�H�[�����̐ݒ���W�i�c�̈ʒu�B�����j
# Args[2] : �e�L�X�g�{�b�N�X�̉���
# Args[3] : �e�L�X�g�{�b�N�X�̍���
# Args[4] : �e�L�X�g�{�b�N�X��\������t�H�[��
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


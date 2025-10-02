# Windows Forms アセンブリをロード
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------------------------------------------------------
# メインのファイル処理を行う関数 (この関数自体に変更はありません)
# ----------------------------------------------------------------------------
function Process-File {
    param(
        [string]$SourceFile,
        [string]$DestinationFolder,
        [System.Windows.Forms.TextBox]$LogTextBox
    )

    # ログ出力用のヘルパー関数
    function Write-Log {
        param([string]$Message)
        $LogTextBox.AppendText("$Message`r`n")
        $LogTextBox.Update()
    }

    # ファイルが存在しない場合はスキップ
    if (-not (Test-Path -Path $SourceFile -PathType Leaf)) {
        Write-Log "[警告] ファイルが見つかりません。スキップします: $SourceFile"
        return
    }

    $FileName = Split-Path -Path $SourceFile -Leaf
    $TempFilePath = Join-Path -Path $env:TEMP -ChildPath $FileName

    try {
        if (-not (Test-Path -Path $DestinationFolder -PathType Container)) {
            Write-Log "移動先フォルダが存在しないため、作成します: $DestinationFolder"
            New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        }

        Write-Log "ファイルを一時フォルダにコピーしています..."
        Write-Log "  コピー元: $SourceFile"
        Copy-Item -Path $SourceFile -Destination $TempFilePath -Force

        Write-Log "代替データストリームを削除しています..."
        Remove-Item -Path $TempFilePath -Stream * -ErrorAction Stop
        Write-Log "代替データストリームの削除が完了しました。"
        
        $FinalPath = Join-Path -Path $DestinationFolder -ChildPath $FileName
        Write-Log "処理後のファイルを移動しています..."
        Write-Log "  移動先: $FinalPath"
        Move-Item -Path $TempFilePath -Destination $FinalPath -Force
        
    } catch {
        Write-Log "[エラー] '$FileName' の処理に失敗しました: $($_.Exception.Message)"
    } finally {
        if (Test-Path -Path $TempFilePath) {
            Write-Log "クリーンアップ処理: 一時ファイルを削除します。"
            Remove-Item -Path $TempFilePath -Force
        }
    }
}

# ----------------------------------------------------------------------------
# GUIの作成 (リストボックスなど、UI要素を大幅に変更)
# ----------------------------------------------------------------------------

# メインフォーム
$form = New-Object System.Windows.Forms.Form
$form.Text = "代替データストリーム削除ツール (複数ファイル対応)"
$form.Size = New-Object System.Drawing.Size(620, 550)
$form.StartPosition = 'CenterScreen'

# ラベル: 処理するファイル
$labelSource = New-Object System.Windows.Forms.Label
$labelSource.Text = "処理するファイルリスト:"
$labelSource.Location = New-Object System.Drawing.Point(20, 15)
$labelSource.Size = New-Object System.Drawing.Size(150, 20)
$form.Controls.Add($labelSource)

# リストボックス: 処理するファイル一覧
$listBoxSourceFiles = New-Object System.Windows.Forms.ListBox
$listBoxSourceFiles.Location = New-Object System.Drawing.Point(20, 40)
$listBoxSourceFiles.Size = New-Object System.Drawing.Size(460, 150)
$listBoxSourceFiles.SelectionMode = "MultiExtended" # ShiftやCtrlキーでの複数選択を許可
$form.Controls.Add($listBoxSourceFiles)

# ボタン: ファイル追加
$buttonAddFiles = New-Object System.Windows.Forms.Button
$buttonAddFiles.Text = "ファイル追加..."
$buttonAddFiles.Location = New-Object System.Drawing.Point(490, 40)
$buttonAddFiles.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($buttonAddFiles)

# ボタン: 選択を削除
$buttonRemoveFile = New-Object System.Windows.Forms.Button
$buttonRemoveFile.Text = "選択を削除"
$buttonRemoveFile.Location = New-Object System.Drawing.Point(490, 75)
$buttonRemoveFile.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($buttonRemoveFile)

# ボタン: リストをクリア
$buttonClearList = New-Object System.Windows.Forms.Button
$buttonClearList.Text = "リストをクリア"
$buttonClearList.Location = New-Object System.Drawing.Point(490, 110)
$buttonClearList.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($buttonClearList)

# ラベル: 保存先フォルダ
$labelDest = New-Object System.Windows.Forms.Label
$labelDest.Text = "保存先フォルダ:"
$labelDest.Location = New-Object System.Drawing.Point(20, 210)
$labelDest.Size = New-Object System.Drawing.Size(120, 20)
$form.Controls.Add($labelDest)

# テキストボックス: 保存先フォルダパス
$textBoxDest = New-Object System.Windows.Forms.TextBox
$textBoxDest.Location = New-Object System.Drawing.Point(150, 207)
$textBoxDest.Size = New-Object System.Drawing.Size(330, 20)
$form.Controls.Add($textBoxDest)

# ボタン: フォルダ選択
$buttonBrowseDest = New-Object System.Windows.Forms.Button
$buttonBrowseDest.Text = "参照..."
$buttonBrowseDest.Location = New-Object System.Drawing.Point(490, 205)
$buttonBrowseDest.Size = New-Object System.Drawing.Size(100, 23)
$form.Controls.Add($buttonBrowseDest)

# ボタン: 実行
$buttonRun = New-Object System.Windows.Forms.Button
$buttonRun.Text = "実行"
$buttonRun.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
$buttonRun.Location = New-Object System.Drawing.Point(250, 245)
$buttonRun.Size = New-Object System.Drawing.Size(120, 35)
$form.Controls.Add($buttonRun)

# テキストボックス: ログ/ステータス表示用
$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(20, 295)
$logBox.Size = New-Object System.Drawing.Size(570, 200)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$logBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$form.Controls.Add($logBox)

# ----------------------------------------------------------------------------
# GUIのイベント処理
# ----------------------------------------------------------------------------

# ファイル追加ボタンのクリックイベント
$buttonAddFiles.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "処理するファイルを選択してください（複数選択可）"
    $openFileDialog.Multiselect = $true # ★★★ 複数選択を有効にする ★★★

    if ($openFileDialog.ShowDialog() -eq 'OK') {
        # 選択された全てのファイル名をリストボックスに追加
        foreach ($file in $openFileDialog.FileNames) {
            # 重複を避ける
            if (-not $listBoxSourceFiles.Items.Contains($file)) {
                $listBoxSourceFiles.Items.Add($file)
            }
        }
    }
})

# 選択を削除ボタンのクリックイベント
$buttonRemoveFile.Add_Click({
    # 選択されている項目をリストから削除（逆順にループするのが安全）
    for ($i = $listBoxSourceFiles.SelectedIndices.Count - 1; $i -ge 0; $i--) {
        $listBoxSourceFiles.Items.RemoveAt($listBoxSourceFiles.SelectedIndices[$i])
    }
})

# リストをクリアボタンのクリックイベント
$buttonClearList.Add_Click({
    $listBoxSourceFiles.Items.Clear()
})

# フォルダ選択ボタンのクリックイベント
$buttonBrowseDest.Add_Click({
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "保存先のフォルダを選択してください"
    if ($folderBrowserDialog.ShowDialog() -eq 'OK') {
        $textBoxDest.Text = $folderBrowserDialog.SelectedPath
    }
})

# 実行ボタンのクリックイベント
$buttonRun.Add_Click({
    # 入力チェック
    if ($listBoxSourceFiles.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("処理するファイルをリストに追加してください。", "入力エラー", "OK", "Error")
        return
    }
    if ([string]::IsNullOrWhiteSpace($textBoxDest.Text)) {
        [System.Windows.Forms.MessageBox]::Show("保存先フォルダを指定してください。", "入力エラー", "OK", "Error")
        return
    }
    
    # 処理中はボタンを無効化し、ログをクリア
    $buttonRun.Enabled = $false
    $logBox.Clear()
    $form.Update()

    # ★★★ リスト内の全てのファイルに対して処理を実行するループ ★★★
    $filesToProcess = $listBoxSourceFiles.Items
    $logBox.AppendText("処理を開始します... 対象: $($filesToProcess.Count) ファイル`r`n")
    $logBox.AppendText("----------------------------------------`r`n")
    
    foreach ($file in $filesToProcess) {
        $logBox.AppendText("`r`n>>> 処理中: $(Split-Path $file -Leaf)`r`n")
        $form.Update()

        Process-File -SourceFile $file -DestinationFolder $textBoxDest.Text -LogTextBox $logBox
    }

    $logBox.AppendText("`r`n----------------------------------------`r`n")
    $logBox.AppendText("全ての処理が完了しました。`r`n")

    # 処理完了後にボタンを再度有効化
    $buttonRun.Enabled = $true
})


# フォームを表示
$form.ShowDialog() | Out-Null
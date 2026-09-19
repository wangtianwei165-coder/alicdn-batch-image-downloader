#Requires -Version 5.1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:settingsDirectory = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'AlibabaImageDownloader'
$script:settingsFile = Join-Path $script:settingsDirectory 'settings.json'

function Get-UserSettings {
    try {
        if (Test-Path -LiteralPath $script:settingsFile) {
            return (Get-Content -LiteralPath $script:settingsFile -Raw -Encoding UTF8 | ConvertFrom-Json)
        }
    } catch {}
    return $null
}

function Save-UserSettings {
    param([string]$DownloadFolder, [int]$NamingIndex, [int]$RetryCount)
    try {
        if (-not (Test-Path -LiteralPath $script:settingsDirectory)) {
            New-Item -ItemType Directory -Path $script:settingsDirectory -Force | Out-Null
        }
        [pscustomobject]@{
            DownloadFolder = $DownloadFolder
            NamingIndex = $NamingIndex
            RetryCount = $RetryCount
        } | ConvertTo-Json | Set-Content -LiteralPath $script:settingsFile -Encoding UTF8
    } catch {}
}

$savedSettings = Get-UserSettings

function New-Label {
    param([string]$Text, [int]$X, [int]$Y, [int]$Width = 150, [int]$Height = 24)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Location = New-Object System.Drawing.Point($X, $Y)
    $label.Size = New-Object System.Drawing.Size($Width, $Height)
    $label.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
    return $label
}

function Get-ImageLinks {
    param([string]$Text)

    $matches = [regex]::Matches($Text, 'https?://[^\s<>"'']+', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $seen = @{}
    $links = New-Object System.Collections.Generic.List[string]
    foreach ($match in $matches) {
        $url = $match.Value.Trim().TrimEnd(')', ']', '}', ',', ';', '，', '；')
        if (-not $seen.ContainsKey($url)) {
            $seen[$url] = $true
            $links.Add($url)
        }
    }
    return $links.ToArray()
}

function Get-SafeFileName {
    param([string]$Name)
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($char in $invalid) {
        $Name = $Name.Replace([string]$char, '_')
    }
    $Name = $Name.Trim().TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'image' }
    if ($Name.Length -gt 150) { $Name = $Name.Substring(0, 150) }
    return $Name
}

function Get-ExtensionFromUrl {
    param([string]$Url)
    try {
        $uri = [Uri]$Url
        $extension = [System.IO.Path]::GetExtension($uri.AbsolutePath).ToLowerInvariant()
        if ($extension -in @('.jpg', '.jpeg', '.png', '.webp', '.gif', '.bmp', '.tif', '.tiff', '.avif')) {
            return $extension
        }
    } catch {}
    return '.jpg'
}

function Get-UniquePath {
    param([string]$Folder, [string]$BaseName, [string]$Extension)
    $candidate = Join-Path $Folder ($BaseName + $Extension)
    $index = 2
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $Folder (('{0}_{1}{2}' -f $BaseName, $index, $Extension))
        $index++
    }
    return $candidate
}

$form = New-Object System.Windows.Forms.Form
$form.Text = '阿里国际站图片批量下载工具 v1.2'
$form.StartPosition = 'CenterScreen'
$form.Size = New-Object System.Drawing.Size(900, 720)
$form.MinimumSize = New-Object System.Drawing.Size(900, 620)
$form.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(247, 249, 252)

$title = New-Label '阿里国际站图片批量下载工具' 22 18 420 34
$title.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 17, [System.Drawing.FontStyle]::Bold)
$title.ForeColor = [System.Drawing.Color]::FromArgb(30, 64, 175)
$form.Controls.Add($title)

$tip = New-Label '把 Excel 中的一整列图片链接粘贴到下面；也支持从混合文字中自动提取链接。' 24 55 780 25
$tip.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
$form.Controls.Add($tip)

$linkLabel = New-Label '图片链接（每行一条）' 24 91 220 24
$form.Controls.Add($linkLabel)

$txtLinks = New-Object System.Windows.Forms.TextBox
$txtLinks.Multiline = $true
$txtLinks.ScrollBars = 'Both'
$txtLinks.AcceptsReturn = $true
$txtLinks.AcceptsTab = $true
$txtLinks.WordWrap = $false
$txtLinks.Location = New-Object System.Drawing.Point(24, 118)
$txtLinks.Size = New-Object System.Drawing.Size(835, 235)
$txtLinks.Anchor = 'Top,Left,Right'
$txtLinks.Font = New-Object System.Drawing.Font('Consolas', 9.5)
$form.Controls.Add($txtLinks)

$btnPaste = New-Object System.Windows.Forms.Button
$btnPaste.Text = '从剪贴板粘贴'
$btnPaste.Location = New-Object System.Drawing.Point(24, 364)
$btnPaste.Size = New-Object System.Drawing.Size(130, 32)
$form.Controls.Add($btnPaste)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = '清空'
$btnClear.Location = New-Object System.Drawing.Point(164, 364)
$btnClear.Size = New-Object System.Drawing.Size(78, 32)
$form.Controls.Add($btnClear)

$lblCount = New-Label '已识别 0 条链接' 260 370 220 24
$lblCount.ForeColor = [System.Drawing.Color]::FromArgb(75, 85, 99)
$form.Controls.Add($lblCount)

$folderLabel = New-Label '保存位置' 24 414 100 24
$form.Controls.Add($folderLabel)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(24, 440)
$txtFolder.Size = New-Object System.Drawing.Size(690, 28)
$txtFolder.Anchor = 'Top,Left,Right'
if ($savedSettings -and -not [string]::IsNullOrWhiteSpace([string]$savedSettings.DownloadFolder)) {
    $txtFolder.Text = [string]$savedSettings.DownloadFolder
} else {
    $txtFolder.Text = [Environment]::GetFolderPath('MyPictures')
}
$form.Controls.Add($txtFolder)

$btnFolder = New-Object System.Windows.Forms.Button
$btnFolder.Text = '选择文件夹'
$btnFolder.Location = New-Object System.Drawing.Point(724, 437)
$btnFolder.Size = New-Object System.Drawing.Size(135, 33)
$btnFolder.Anchor = 'Top,Right'
$form.Controls.Add($btnFolder)

$nameLabel = New-Label '文件命名' 24 484 80 24
$form.Controls.Add($nameLabel)

$cmbNaming = New-Object System.Windows.Forms.ComboBox
$cmbNaming.DropDownStyle = 'DropDownList'
$cmbNaming.Location = New-Object System.Drawing.Point(115, 481)
$cmbNaming.Size = New-Object System.Drawing.Size(245, 28)
[void]$cmbNaming.Items.Add('保留阿里原始文件名')
[void]$cmbNaming.Items.Add('按顺序自动编号 001、002…')
if ($savedSettings -and ([int]$savedSettings.NamingIndex -in @(0, 1))) {
    $cmbNaming.SelectedIndex = [int]$savedSettings.NamingIndex
} else {
    $cmbNaming.SelectedIndex = 0
}
$form.Controls.Add($cmbNaming)

$retryLabel = New-Label '失败重试' 382 484 75 24
$form.Controls.Add($retryLabel)

$numRetry = New-Object System.Windows.Forms.NumericUpDown
$numRetry.Location = New-Object System.Drawing.Point(459, 481)
$numRetry.Size = New-Object System.Drawing.Size(58, 28)
$numRetry.Minimum = 0
$numRetry.Maximum = 10
if ($savedSettings -and ([int]$savedSettings.RetryCount -ge 0) -and ([int]$savedSettings.RetryCount -le 10)) {
    $numRetry.Value = [int]$savedSettings.RetryCount
} else {
    $numRetry.Value = 3
}
$form.Controls.Add($numRetry)

$retryUnit = New-Label '次' 522 484 35 24
$form.Controls.Add($retryUnit)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = '一键开始下载'
$btnStart.Location = New-Object System.Drawing.Point(575, 477)
$btnStart.Size = New-Object System.Drawing.Size(180, 40)
$btnStart.Anchor = 'Top,Right'
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(37, 99, 235)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'
$btnStart.Font = New-Object System.Drawing.Font('Microsoft YaHei UI', 10, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnStart)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = '取消'
$btnCancel.Location = New-Object System.Drawing.Point(765, 477)
$btnCancel.Size = New-Object System.Drawing.Size(94, 40)
$btnCancel.Anchor = 'Top,Right'
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(24, 532)
$progress.Size = New-Object System.Drawing.Size(835, 24)
$progress.Anchor = 'Top,Left,Right'
$form.Controls.Add($progress)

$lblStatus = New-Label '等待开始' 24 562 835 24
$lblStatus.Anchor = 'Top,Left,Right'
$form.Controls.Add($lblStatus)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.Location = New-Object System.Drawing.Point(24, 590)
$txtLog.Size = New-Object System.Drawing.Size(835, 80)
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$txtLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtLog)

$script:isDownloading = $false
$script:cancelRequested = $false

$updateCount = {
    $count = (Get-ImageLinks $txtLinks.Text).Count
    $lblCount.Text = "已识别 $count 条链接（自动去重）"
}

$txtLinks.Add_TextChanged($updateCount)
$btnPaste.Add_Click({
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $txtLinks.Text = [System.Windows.Forms.Clipboard]::GetText()
    }
})
$btnClear.Add_Click({ $txtLinks.Clear() })

$btnFolder.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = '选择图片保存文件夹'
    if (Test-Path -LiteralPath $txtFolder.Text) { $dialog.SelectedPath = $txtFolder.Text }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dialog.SelectedPath
        Save-UserSettings $txtFolder.Text $cmbNaming.SelectedIndex ([int]$numRetry.Value)
    }
})

$btnStart.Add_Click({
    $links = @(Get-ImageLinks $txtLinks.Text)
    if ($links.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('没有识别到有效链接，请先粘贴图片 URL。', '提示', 'OK', 'Warning') | Out-Null
        return
    }

    $folder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($folder)) {
        [System.Windows.Forms.MessageBox]::Show('请选择保存文件夹。', '提示', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        $testFile = Join-Path $folder ('.write_test_' + [Guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($testFile, '')
        Remove-Item -LiteralPath $testFile -Force
    } catch {
        [System.Windows.Forms.MessageBox]::Show('保存文件夹无法写入，请重新选择。', '提示', 'OK', 'Error') | Out-Null
        return
    }

    Save-UserSettings $folder $cmbNaming.SelectedIndex ([int]$numRetry.Value)

    $progress.Value = 0
    $txtLog.Clear()
    $btnStart.Enabled = $false
    $btnCancel.Enabled = $true
    $txtLinks.Enabled = $false
    $btnFolder.Enabled = $false
    $script:isDownloading = $true
    $script:cancelRequested = $false

    $retryCount = [int]$numRetry.Value
    $numbered = ($cmbNaming.SelectedIndex -eq 1)
    $success = 0
    $failed = 0
    $failures = New-Object System.Collections.Generic.List[string]
    $fatalError = $null

    try {
        for ($i = 0; $i -lt $links.Count; $i++) {
            [System.Windows.Forms.Application]::DoEvents()
            if ($script:cancelRequested) { break }

            $url = $links[$i]
            $lblStatus.Text = "正在下载 $($i + 1)/$($links.Count)：$url"
            [System.Windows.Forms.Application]::DoEvents()

            $extension = Get-ExtensionFromUrl $url
            if ($numbered) {
                $digits = [Math]::Max(3, $links.Count.ToString().Length)
                $baseName = ($i + 1).ToString(('D' + $digits))
            } else {
                try {
                    $baseName = [System.IO.Path]::GetFileNameWithoutExtension(([Uri]$url).AbsolutePath)
                } catch {
                    $baseName = 'image_' + ($i + 1)
                }
                $baseName = Get-SafeFileName $baseName
            }

            $target = Get-UniquePath $folder $baseName $extension
            $temp = $target + '.part'
            $downloaded = $false
            $lastError = ''

            for ($attempt = 0; $attempt -le $retryCount; $attempt++) {
                [System.Windows.Forms.Application]::DoEvents()
                if ($script:cancelRequested) { break }
                $client = $null
                try {
                    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force }
                    $client = New-Object System.Net.WebClient
                    $client.Headers.Add('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36')
                    $client.Headers.Add('Referer', 'https://www.alibaba.com/')
                    $client.DownloadFile($url, $temp)
                    $client.Dispose()
                    if ((Test-Path -LiteralPath $temp) -and ((Get-Item -LiteralPath $temp).Length -gt 0)) {
                        Move-Item -LiteralPath $temp -Destination $target -Force
                        $downloaded = $true
                        break
                    }
                    throw '下载文件为空'
                } catch {
                    $lastError = $_.Exception.Message
                    if ($client) { $client.Dispose() }
                    if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
                    if ($attempt -lt $retryCount -and -not $script:cancelRequested) {
                        $waitUntil = (Get-Date).AddSeconds([Math]::Min(2 + $attempt, 5))
                        while ((Get-Date) -lt $waitUntil) {
                            [System.Windows.Forms.Application]::DoEvents()
                            if ($script:cancelRequested) { break }
                            Start-Sleep -Milliseconds 100
                        }
                    }
                }
            }

            if ($script:cancelRequested) {
                if (Test-Path -LiteralPath $temp) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
                break
            }

            if ($downloaded) {
                $success++
                $logMessage = "成功：$([System.IO.Path]::GetFileName($target))"
            } else {
                $failed++
                $failures.Add("$url`t$lastError")
                $logMessage = "失败：$url"
            }

            $progress.Value = [int]((($i + 1) / $links.Count) * 100)
            $lblStatus.Text = "进度 $($i + 1)/$($links.Count)    成功 $success    失败 $failed"
            $txtLog.AppendText($logMessage + [Environment]::NewLine)
            $txtLog.SelectionStart = $txtLog.Text.Length
            $txtLog.ScrollToCaret()
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($failures.Count -gt 0) {
            $failurePath = Join-Path $folder ('下载失败链接_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.txt')
            [System.IO.File]::WriteAllLines($failurePath, $failures, (New-Object System.Text.UTF8Encoding($true)))
        }
    } catch {
        $fatalError = $_.Exception.Message
    } finally {
        $script:isDownloading = $false
        $btnStart.Enabled = $true
        $btnCancel.Enabled = $false
        $txtLinks.Enabled = $true
        $btnFolder.Enabled = $true
    }

    if ($fatalError) {
        $lblStatus.Text = '运行出错'
        [System.Windows.Forms.MessageBox]::Show($fatalError, '运行出错', 'OK', 'Error') | Out-Null
    } elseif ($script:cancelRequested) {
        $lblStatus.Text = "下载已取消：成功 $success，失败 $failed"
        [System.Windows.Forms.MessageBox]::Show('下载已取消，已经成功下载的图片会保留。', '已取消', 'OK', 'Information') | Out-Null
    } else {
        $lblStatus.Text = "下载完成：成功 $success，失败 $failed"
        $message = "下载完成！`r`n`r`n成功：$success 张`r`n失败：$failed 张`r`n保存位置：$folder"
        if ($failed -gt 0) { $message += "`r`n`r`n失败链接已保存为 txt 文件。" }
        [System.Windows.Forms.MessageBox]::Show($message, '下载完成', 'OK', 'Information') | Out-Null
    }
})

$btnCancel.Add_Click({
    $script:cancelRequested = $true
    $btnCancel.Enabled = $false
    $lblStatus.Text = '将在当前图片处理结束后取消…'
})

$form.Add_FormClosing({
    param($sender, $e)
    Save-UserSettings $txtFolder.Text $cmbNaming.SelectedIndex ([int]$numRetry.Value)
    if ($script:isDownloading) {
        $answer = [System.Windows.Forms.MessageBox]::Show('图片还在下载，确定要退出吗？', '确认退出', 'YesNo', 'Question')
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) {
            $e.Cancel = $true
        } else {
            $script:cancelRequested = $true
        }
    }
})

[void]$form.ShowDialog()

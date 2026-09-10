# ============================================
# Outlookメール保存ツール（Wordセクション対策版）
# ============================================

# --- 設定 ---
$CONFIG = @{
    DesktopPath = [Environment]::GetFolderPath("Desktop")
    EdgePaths = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe"
    )
    MaxSubjectLength = 200
    MaxSenderLength = 50
    MaxPathLength = 240
    PdfTimeout = 30
    OpenFolderAfterProcess = $true
    EnableLogging = $true
    # PDF生成設定
    PdfMinFileSize = 100
    PdfRetryCount = 10
    PdfRetryInterval = 500
    # Outlook設定
    OutlookMailItemClass = 43
}

# --- グローバル変数 ---
$script:LogFilePath = $null
$script:LogFileWriteWarned = $false

# ============================================
# ユーティリティ関数
# ============================================

function Initialize-LogFile {
    # 実行単位のログファイルを確定する。失敗してもコンソール出力のみで処理は継続する。
    try {
        $logDir = Join-Path $CONFIG.DesktopPath "MailExporter_Logs"
        if (-not [System.IO.Directory]::Exists($logDir)) {
            [System.IO.Directory]::CreateDirectory($logDir) | Out-Null
        }
        $logFileName = "MailExporter_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
        $script:LogFilePath = Join-Path $logDir $logFileName
        [System.IO.File]::WriteAllText($script:LogFilePath, "", [System.Text.Encoding]::UTF8)
    } catch {
        $script:LogFilePath = $null
        Write-Host "ログファイルの初期化に失敗しました。コンソール出力のみで処理を継続します: $_" -ForegroundColor Yellow
    }
}

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($CONFIG.EnableLogging) {
        if ($script:LogFilePath) {
            try {
                [System.IO.File]::AppendAllText($script:LogFilePath, "$logEntry`r`n", [System.Text.Encoding]::UTF8)
            } catch {
                # ログ保存に失敗してもメール処理全体は止めず、コンソール出力のみ継続する
                if (-not $script:LogFileWriteWarned) {
                    $script:LogFileWriteWarned = $true
                    Write-Host "ログファイルへの書き込みに失敗しました。以降はコンソール出力のみになります: $_" -ForegroundColor Yellow
                }
            }
        }
    }

    $colors = @{
        'Info' = 'White'
        'Warning' = 'Yellow'
        'Error' = 'Red'
        'Success' = 'Green'
    }

    Write-Host $Message -ForegroundColor $colors[$Level]
}

function Write-ErrorDetail {
    param(
        [string]$MailSubject,
        [string]$Stage,
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    $ex = $ErrorRecord.Exception
    $invocation = $ErrorRecord.InvocationInfo

    Write-Log "  [エラー詳細]" -Level Error
    Write-Log "    発生日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Error
    Write-Log "    処理対象メール: $MailSubject" -Level Error
    Write-Log "    処理段階: $Stage" -Level Error
    Write-Log "    例外種別: $($ex.GetType().FullName)" -Level Error
    Write-Log "    Exception.Message: $($ex.Message)" -Level Error
    if ($invocation) {
        Write-Log "    発生位置: $($invocation.ScriptName):$($invocation.ScriptLineNumber) $($invocation.Line.Trim())" -Level Error
    }
    if ($ErrorRecord.ScriptStackTrace) {
        Write-Log "    ScriptStackTrace: $($ErrorRecord.ScriptStackTrace)" -Level Error
    }
}

function Get-SafeFilename {
    param(
        [string]$str,
        [int]$maxLength = 50
    )
    
    if ([string]::IsNullOrEmpty($str)) { return "NoName" }
    
    $replaceMap = @{
        '\' = '￥'; '/' = '／'; ':' = '：'; '*' = '＊'
        '?' = '？'; '"' = '"'; '<' = '＜'; '>' = '＞'; '|' = '｜'
    }
    
    $safe = $str
    foreach ($key in $replaceMap.Keys) {
        $safe = $safe -replace [regex]::Escape($key), $replaceMap[$key]
    }
    
    $safe = $safe -replace '[\r\n\t\x00-\x1F]', '' -replace '\s+', ' '
    $safe = $safe.Trim()
    
    if ($safe.Length -gt $maxLength) { 
        $safe = $safe.Substring(0, $maxLength).TrimEnd()
    }
    
    return $(if ([string]::IsNullOrEmpty($safe)) { "NoName" } else { $safe })
}

function Get-SafeFolderPath {
    param(
        [string]$basePath,
        [string]$folderName,
        [int]$maxLength = 240
    )
    
    $fullPath = Join-Path $basePath $folderName
    
    if ($fullPath.Length -le $maxLength) {
        return @{ Path = $fullPath; Truncated = $false }
    }
    
    if ($folderName -match '^(\d{8})_(.+)$') {
        $datePrefix = $Matches[1]
        $subject = $Matches[2]
        $availableLength = $maxLength - $basePath.Length - $datePrefix.Length - 11
        
        if ($availableLength -gt 10) {
            $truncatedSubject = $subject.Substring(0, $availableLength)
            $fullPath = Join-Path $basePath "${datePrefix}_${truncatedSubject}"
            return @{ Path = $fullPath; Truncated = $true }
        }
    }
    
    if ($folderName -match '^(\d{8})') {
        $fullPath = Join-Path $basePath $Matches[1]
    }
    
    return @{ Path = $fullPath; Truncated = $true }
}

function Release-Ref {
    param($ref)
    
    if ($ref -ne $null) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) | Out-Null
        } catch {
            # 既に解放済みの場合は無視
        }
        Remove-Variable ref -ErrorAction SilentlyContinue
    }
}

function Find-EdgePath {
    foreach ($path in $CONFIG.EdgePaths) {
        if (Test-Path $path) { return $path }
    }
    return $null
}

function Escape-HtmlText {
    param([string]$text)
    
    if ([string]::IsNullOrEmpty($text)) { return "" }
    
    return $text -replace '&', '&amp;' `
                -replace '<', '&lt;' `
                -replace '>', '&gt;' `
                -replace '"', '&quot;' `
                -replace "'", '&#39;'
}

# ============================================
# メール情報取得関数
# ============================================

function Get-MailMetadata {
    param($mail)
    
    try {
        $receivedDate = $mail.ReceivedTime
        $dateStr = $receivedDate.ToString("yyyyMMdd")
        $dateTimeStr = $receivedDate.ToString("yyyy年MM月dd日 HH:mm")
    } catch {
        $receivedDate = Get-Date
        $dateStr = $receivedDate.ToString("yyyyMMdd")
        $dateTimeStr = $receivedDate.ToString("yyyy年MM月dd日 HH:mm")
        Write-Log "  受信日取得失敗。現在日時を使用" -Level Warning
    }
    
    $subject = if ([string]::IsNullOrEmpty($mail.Subject)) { "(件名なし)" } else { $mail.Subject }
    $senderName = if ([string]::IsNullOrEmpty($mail.SenderName)) { "(不明)" } else { $mail.SenderName }
    $senderEmail = if ([string]::IsNullOrEmpty($mail.SenderEmailAddress)) { "" } else { $mail.SenderEmailAddress }
    $senderInfo = if ($senderEmail) { "$senderName <$senderEmail>" } else { $senderName }
    $toAddr = if ([string]::IsNullOrEmpty($mail.To)) { "(なし)" } else { $mail.To }
    $ccAddr = if ([string]::IsNullOrEmpty($mail.CC)) { "(なし)" } else { $mail.CC }
    
    return @{
        Subject = $subject
        SenderName = $senderName
        SenderInfo = $senderInfo
        ToAddr = $toAddr
        CcAddr = $ccAddr
        DateStr = $dateStr
        DateTimeStr = $dateTimeStr
    }
}

# ============================================
# フォルダ・ファイル処理関数
# ============================================

function New-MailFolder {
    param(
        [string]$dateStr,
        [string]$subject
    )
    
    $safeSubject = Get-SafeFilename $subject -maxLength $CONFIG.MaxSubjectLength
    $baseFolderName = "${dateStr}_${safeSubject}"
    $pathInfo = Get-SafeFolderPath -basePath $CONFIG.DesktopPath `
                                    -folderName $baseFolderName `
                                    -maxLength $CONFIG.MaxPathLength
    
    if ($pathInfo.Truncated) {
        Write-Log "  パス長制限のためフォルダ名を短縮" -Level Warning
    }
    
    $outputDir = $pathInfo.Path
    $counter = 1
    $originalPath = $outputDir

    # ワイルドカード解釈を避けるため、.NET のパス存在確認/作成APIを直接使用する
    # (Split-Path は -Resolve を付けない限りファイルシステムに触れない文字列処理のため、
    #  -Leaf 指定時はワイルドカード展開の影響を受けない)
    while ([System.IO.Directory]::Exists($outputDir) -or [System.IO.File]::Exists($outputDir)) {
        $counter++
        $folderName = Split-Path -Path $originalPath -Leaf
        $outputDir = Join-Path $CONFIG.DesktopPath "${folderName}_${counter}"
    }

    [System.IO.Directory]::CreateDirectory($outputDir) | Out-Null
    Write-Log "  フォルダ: $(Split-Path -Path $outputDir -Leaf)" -Level Success

    return $outputDir
}

function Save-MailAttachments {
    param(
        $mail,
        [string]$outputDir
    )
    
    $attachmentNames = @()
    
    if ($mail.Attachments.Count -eq 0) {
        return $attachmentNames
    }
    
    Write-Log "  添付処理: $($mail.Attachments.Count) 件" -Level Info
    
    foreach ($att in $mail.Attachments) {
        try {
            $attName = $att.FileName
            if ([string]::IsNullOrEmpty($attName)) { continue }
            
            $safeAttName = Get-SafeFilename $attName -maxLength 200
            $savePath = Join-Path $outputDir $safeAttName

            if ([System.IO.File]::Exists($savePath)) {
                $ext = [System.IO.Path]::GetExtension($safeAttName)
                $base = [System.IO.Path]::GetFileNameWithoutExtension($safeAttName)
                $dupCounter = 1
                do {
                    $savePath = Join-Path $outputDir "${base}_${dupCounter}${ext}"
                    $dupCounter++
                } while ([System.IO.File]::Exists($savePath))
            }
            
            $att.SaveAsFile($savePath)
            $attachmentNames += $safeAttName
            
        } catch {
            Write-Log "  添付エラー: $attName - $_" -Level Warning
        }
    }
    
    Write-Log "  添付保存: $($attachmentNames.Count) 件完了" -Level Success
    return $attachmentNames
}

# ============================================
# HTML・PDF生成関数（Wordセクション対策強化）
# ============================================

function Get-MailHeaderHtml {
    param($metadata)
    
    # 改ページ対策CSS
    return @"
<div style="border: 1px solid #000; padding: 10px; margin-bottom: 10px; font-family: 'BIZ UDGothic', 'Meiryo', 'Yu Gothic', sans-serif; page-break-inside: avoid; break-inside: avoid; page-break-after: avoid; break-after: avoid;">
    <table style="width: 100%; border-collapse: collapse; font-size: 10pt;">
        <tr>
            <td style="padding: 1px; width: 90px; font-weight: bold; border-bottom: 1px solid #ccc;">件名</td>
            <td style="padding: 1px; border-bottom: 1px solid #ccc;">$(Escape-HtmlText $metadata.Subject)</td>
        </tr>
        <tr>
            <td style="padding: 1px; font-weight: bold; border-bottom: 1px solid #ccc;">差出人</td>
            <td style="padding: 1px; border-bottom: 1px solid #ccc;">$(Escape-HtmlText $metadata.SenderInfo)</td>
        </tr>
        <tr>
            <td style="padding: 1px; font-weight: bold; border-bottom: 1px solid #ccc;">宛先</td>
            <td style="padding: 1px; border-bottom: 1px solid #ccc;">$(Escape-HtmlText $metadata.ToAddr)</td>
        </tr>
        <tr>
            <td style="padding: 1px; font-weight: bold; border-bottom: 1px solid #ccc;">CC</td>
            <td style="padding: 1px; border-bottom: 1px solid #ccc;">$(Escape-HtmlText $metadata.CcAddr)</td>
        </tr>
        <tr>
            <td style="padding: 1px; font-weight: bold;">受信日時</td>
            <td style="padding: 1px;">$($metadata.DateTimeStr)</td>
        </tr>
    </table>
</div>
"@
}

function Get-MailFooterHtml {
    param([array]$attachmentNames)
    
    $attListStr = if ($attachmentNames.Count -gt 0) {
        ($attachmentNames | ForEach-Object { "• $_" }) -join "<br>"
    } else {
        "(添付ファイルなし)"
    }
    
    return @"
<div style="border-top: 1px solid #000; padding: 10px; margin-top: 10px; font-family: 'BIZ UDGothic', 'Meiryo', sans-serif; page-break-inside: avoid; break-inside: avoid; page-break-before: avoid; break-before: avoid;">
    <div style="font-size: 9pt;">
        <strong>添付ファイル:</strong><br>
        <div style="margin-top: 5px; padding-left: 5px;">$attListStr</div>
    </div>
    <div style="margin-top: 10px; padding-top: 5px; border-top: 1px solid #ccc; font-size: 8pt; color: #666; text-align: right;">
        保存日時: $(Get-Date -Format 'yyyy年MM月dd日 HH:mm:ss')
    </div>
</div>
"@
}

function Get-HtmlDocumentStyle {
    # 【Wordセクション対策】WordSection1に対する強力なリセット
    return @"
<style>
    /* 用紙の余白設定：上部を10mmに縮小 */
    @page { margin-top: 10mm; margin-bottom: 15mm; margin-left: 15mm; margin-right: 15mm; }
    
    /* 本文外側の余白を削除 */
    body { margin: 0; padding: 0; page-break-inside: auto; }
    
    /* 強制改ページの無効化 */
    p, div, table, h1, h2, h3, h4, h5, h6 {
        page-break-before: auto !important;
        break-before: auto !important;
    }

    /* WordSection1対策 */
    div.WordSection1 {
        page: auto !important;
        width: 100% !important;
        margin: 0 !important;
        padding: 0 !important;
    }
</style>
"@
}

function New-MailHtml {
    param(
        $mail,
        $metadata,
        [array]$attachmentNames
    )
    
    $headerHtml = Get-MailHeaderHtml -metadata $metadata
    $footerHtml = Get-MailFooterHtml -attachmentNames $attachmentNames
    $styleHtml = Get-HtmlDocumentStyle
    
    $htmlBody = $mail.HTMLBody
    
    if ([string]::IsNullOrEmpty($htmlBody)) {
        # テキストメール処理
        $bodyText = if ([string]::IsNullOrEmpty($mail.Body)) { "(本文なし)" } else { $mail.Body }
        $bodyEscaped = (Escape-HtmlText $bodyText) -replace "`r`n", "<br>" -replace "`n", "<br>"
        
        return @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$(Escape-HtmlText $metadata.Subject)</title>
    $styleHtml
</head>
<body style="font-family: 'Meiryo', 'MS Gothic', sans-serif;">
$headerHtml
<div style="background-color: white; padding: 10px; border: 1px solid #ddd; white-space: pre-wrap; line-height: 1.6; page-break-inside: auto;">
$bodyEscaped
</div>
$footerHtml
</body>
</html>
"@
    } else {
        # HTMLメール処理
        
        # 【最重要】HTML文字列内のWordページ設定定義を無効化（置換）
        # @page WordSection1 {...} のような定義を無効化
        $htmlBody = $htmlBody -replace "@page\s+WordSection1", "@page WordSection1_Disabled"
        # page: WordSection1; というプロパティ指定を無効化
        $htmlBody = $htmlBody -replace "page:\s*WordSection1;?", "page: auto;"
        
        # 文字コード設定
        if ($htmlBody -notmatch "charset\s*=\s*['""]?utf-8") {
            if ($htmlBody -match "<head[^>]*>") {
                $htmlBody = $htmlBody -replace "(<head[^>]*>)", "`$1<meta charset='utf-8'>$styleHtml"
            } else {
                $htmlBody = "<html><head><meta charset='utf-8'>$styleHtml</head>" + $htmlBody
            }
        } else {
            if ($htmlBody -match "<head[^>]*>") {
                $htmlBody = $htmlBody -replace "(<head[^>]*>)", "`$1$styleHtml"
            }
        }
        
        # ラッパー注入
        if ($htmlBody -match "<body[^>]*>") {
            $htmlBody = $htmlBody -replace "(<body[^>]*>)", "`$1<div style='padding:10px; page-break-inside: auto;'>$headerHtml"
            $htmlBody = $htmlBody -replace "</body>", "$footerHtml</div></body>"
        } else {
            $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    $styleHtml
</head>
<body style="margin: 10px;">
<div style="page-break-inside: auto;">
$headerHtml
$htmlBody
$footerHtml
</div>
</body>
</html>
"@
        }
        
        return $htmlBody
    }
}

function Convert-HtmlToPdf {
    param(
        [string]$HtmlPath,
        [string]$PdfPath,
        [string]$EdgePath
    )
    
    try {
        $argList = @(
            "--headless"
            "--disable-gpu"
            "--disable-software-rasterizer"
            "--disable-dev-shm-usage"
            "--run-all-compositor-stages-before-draw"
            "--log-level=3"
            "--disable-logging"
            "--no-sandbox"
            "--print-to-pdf=`"$PdfPath`""
            "`"$HtmlPath`""
        )

        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $EdgePath
        $processInfo.Arguments = $argList -join " "
        $processInfo.UseShellExecute = $false
        $processInfo.RedirectStandardOutput = $true
        $processInfo.RedirectStandardError = $true
        $processInfo.CreateNoWindow = $true
        
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null
        
        $exited = $process.WaitForExit($CONFIG.PdfTimeout * 1000)
        
        if (-not $exited) {
            $process.Kill()
            throw "PDF変換がタイムアウトしました（$($CONFIG.PdfTimeout)秒）"
        }
        
        # PDF生成確認
        $retryCount = 0
        while ($retryCount -lt $CONFIG.PdfRetryCount) {
            if ([System.IO.File]::Exists($PdfPath)) {
                Start-Sleep -Milliseconds 200
                $fileInfo = Get-Item -LiteralPath $PdfPath -ErrorAction SilentlyContinue
                if ($fileInfo -and $fileInfo.Length -gt $CONFIG.PdfMinFileSize) {
                    return $true
                }
            }
            Start-Sleep -Milliseconds $CONFIG.PdfRetryInterval
            $retryCount++
        }
        
        throw "PDFファイルが正しく生成されませんでした"
        
    } catch {
        throw "PDF変換エラー: $_"
    } finally {
        if ($process -and -not $process.HasExited) {
            $process.Kill()
        }
    }
}

# ============================================
# メール処理関数
# ============================================

function Process-SingleMail {
    param(
        $mail,
        [int]$index,
        [int]$total,
        [string]$edgePath
    )

    $stage = "メール種別確認"
    $subjectForLog = "(不明)"

    try {
        if ($mail.Class -ne $CONFIG.OutlookMailItemClass) {
            Write-Log "[$index/$total] スキップ: メールアイテム以外" -Level Warning
            return $null
        }

        # メタデータ取得
        $stage = "メタデータ取得"
        $metadata = Get-MailMetadata -mail $mail
        $subjectForLog = $metadata.Subject
        Write-Log "[$index/$total] 処理中: $($metadata.Subject)" -Level Info

        # フォルダ作成
        $stage = "フォルダ作成"
        $outputDir = New-MailFolder -dateStr $metadata.DateStr -subject $metadata.Subject

        # 添付ファイル保存
        $stage = "添付ファイル保存"
        $attachmentNames = Save-MailAttachments -mail $mail -outputDir $outputDir

        # HTML生成
        $stage = "HTML生成"
        $finalHtml = New-MailHtml -mail $mail -metadata $metadata -attachmentNames $attachmentNames

        # 一時HTMLファイル保存（件名由来のパスに [ ] 等を含んでもワイルドカード解釈されないよう -LiteralPath を使用）
        $stage = "一時HTMLファイル保存"
        $tempHtml = Join-Path $outputDir "_temp_mail.html"
        Set-Content -LiteralPath $tempHtml -Value $finalHtml -Encoding UTF8 -Force

        # PDF変換
        $stage = "PDF変換"
        $safeSender = Get-SafeFilename $metadata.SenderName -maxLength $CONFIG.MaxSenderLength
        $pdfName = "$($metadata.DateStr)_${safeSender}mail.pdf"
        $pdfPath = Join-Path $outputDir $pdfName

        Write-Log "  PDF変換中..." -Level Info

        try {
            $pdfResult = Convert-HtmlToPdf -HtmlPath $tempHtml -PdfPath $pdfPath -EdgePath $edgePath

            if ($pdfResult) {
                Write-Log "  PDF作成: $pdfName" -Level Success
            } else {
                Write-Log "  PDF生成失敗" -Level Error
                return $null
            }
        } finally {
            if ([System.IO.File]::Exists($tempHtml)) {
                Remove-Item -LiteralPath $tempHtml -Force -ErrorAction SilentlyContinue
            }
        }

        Write-Log "  完了 ✓" -Level Success
        Write-Log "--------------------------------------------" -Level Info

        return $outputDir

    } catch {
        Write-ErrorDetail -MailSubject $subjectForLog -Stage $stage -ErrorRecord $_
        return $null
    }
}

# ============================================
# メイン処理
# ============================================

$outlook = $null
$explorer = $null
$selection = $null
$processedCount = 0
$errorCount = 0
$processedFolders = @()

try {
    Initialize-LogFile

    Write-Log "============================================" -Level Info
    Write-Log "Outlookメール保存ツール 開始" -Level Info
    Write-Log "============================================" -Level Info

    # Edge確認
    $edgePath = Find-EdgePath
    if (-not $edgePath) {
        Write-Log "Microsoft Edgeが見つかりません。" -Level Error
        Write-Log "Edgeをインストールするか、パスを確認してください。" -Level Error
        exit 1
    }
    Write-Log "Edge: $edgePath" -Level Info
    
    # Outlook接続
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Write-Log "Outlook接続: 成功" -Level Success
    } catch {
        Write-Log "Outlookが起動していません。" -Level Error
        Write-Log "Outlookを起動してメールを選択してから実行してください。" -Level Error
        exit 1
    }

    $explorer = $outlook.ActiveExplorer()
    $selection = $explorer.Selection

    if ($selection.Count -eq 0) {
        Write-Log "メールが選択されていません。" -Level Warning
        Write-Log "Outlookでメールを選択してから実行してください。" -Level Warning
        exit 1
    }

    Write-Log "選択メール: $($selection.Count) 件" -Level Info
    Write-Log "--------------------------------------------" -Level Info

    # メール処理ループ
    $mailIndex = 0
    foreach ($mail in $selection) {
        $mailIndex++
        
        try {
            $outputDir = Process-SingleMail -mail $mail -index $mailIndex -total $selection.Count -edgePath $edgePath
            
            if ($outputDir) {
                $processedFolders += $outputDir
                $processedCount++
            } else {
                $errorCount++
            }
            
        } catch {
            Write-ErrorDetail -MailSubject "(不明)" -Stage "メール処理(予期しないエラー)" -ErrorRecord $_
            $errorCount++
        }
    }

    # 処理完了サマリー
    Write-Log "============================================" -Level Info
    Write-Log "処理完了" -Level Success
    Write-Log "  成功: $processedCount 件" -Level Success
    if ($errorCount -gt 0) {
        Write-Log "  失敗: $errorCount 件" -Level Warning
    }
    Write-Log "============================================" -Level Info

    Write-Host ""
    Write-Host "処理完了" -ForegroundColor Green
    Write-Host "  成功: $processedCount 件" -ForegroundColor Green
    Write-Host "  失敗: $errorCount 件" -ForegroundColor $(if ($errorCount -gt 0) { 'Yellow' } else { 'Green' })
    if ($script:LogFilePath) {
        Write-Host "  ログ: $script:LogFilePath" -ForegroundColor Green
    }

    # フォルダ自動オープン
    if ($CONFIG.OpenFolderAfterProcess -and $processedFolders.Count -gt 0) {
        Start-Sleep -Milliseconds 500
        if ($processedFolders.Count -eq 1) {
            Write-Log "保存フォルダを開きます..." -Level Info
            Invoke-Item -LiteralPath $processedFolders[0]
        } else {
            Write-Log "デスクトップを開きます..." -Level Info
            Invoke-Item -LiteralPath $CONFIG.DesktopPath
        }
    }

} catch {
    Write-ErrorDetail -MailSubject "(不明)" -Stage "スクリプト全体(予期しないエラー)" -ErrorRecord $_
    exit 1
} finally {
    # COM解放
    Write-Log "クリーンアップ中..." -Level Info
    Release-Ref $selection
    Release-Ref $explorer
    Release-Ref $outlook
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# 終了待機
Write-Host "`n処理が完了しました。" -ForegroundColor Green
# Write-Host "何かキーを押すと終了します..." -ForegroundColor Cyan
# $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
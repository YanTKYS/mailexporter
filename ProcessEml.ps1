# ============================================
# Outlookメール保存ツール（Word文書直接生成版）
# ============================================

# --- 設定 ---
$CONFIG = @{
    DesktopPath = [Environment]::GetFolderPath("Desktop")
    MaxSubjectLength = 200
    MaxSenderLength = 50
    MaxPathLength = 240
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
# ログは「メールごとの保存フォルダ内」に保存する方式を基本とする。
# 起動時のOutlook接続確認・選択メール件数などは実行全体で共通のログとして
# CommonLogBuffer に保持し続け(メールが切り替わっても消去しない)、
# 現在処理中メール固有のログは MailLogBuffer に保持する。
# 保存フォルダが確定した時点で「共通ログ + そのメールのログ」をまとめて
# ログファイルへ書き出してから逐次追記に切り替える。
# 保存フォルダを確定できないまま処理が失敗した場合のみ、両方のバッファを
# デスクトップ直下へ退避する。
$script:CommonLogBuffer = @()
$script:MailLogBuffer = @()
$script:CurrentLogFilePath = $null
$script:InMailContext = $false

# ログレベルごとのコンソール表示色
$script:LogColors = @{
    'Info'    = 'White'
    'Warning' = 'Yellow'
    'Error'   = 'Red'
    'Success' = 'Green'
}

# ============================================
# ユーティリティ関数
# ============================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error', 'Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    if ($CONFIG.EnableLogging) {
        if ($script:CurrentLogFilePath) {
            try {
                [System.IO.File]::AppendAllText($script:CurrentLogFilePath, "$logEntry`r`n", [System.Text.Encoding]::UTF8)
            } catch {
                # ログ保存に失敗してもメール処理全体は止めず、コンソール出力のみ継続する。
                # 同じI/Oエラーを繰り返さないよう、以降のファイル書き込み自体を無効化する。
                Write-Host "ログファイルへの書き込みに失敗しました。以降はコンソール出力のみになります: $_" -ForegroundColor Yellow
                $script:CurrentLogFilePath = $null
            }
        } elseif ($script:InMailContext) {
            $script:MailLogBuffer += $logEntry
        } else {
            $script:CommonLogBuffer += $logEntry
        }
    }

    Write-Host $Message -ForegroundColor $script:LogColors[$Level]
}

function Start-MailLogContext {
    # 1通のメール処理を開始する際に呼ぶ。共通ログ(CommonLogBuffer)は消去せず、
    # そのメール固有の状態だけをリセットする。
    $script:CurrentLogFilePath = $null
    $script:MailLogBuffer = @()
    $script:InMailContext = $true
}

function Stop-MailLogContext {
    # 1通のメール処理を終える際に呼ぶ。以降のログ(次のメールの開始前など)は
    # 共通ログ側へ積まれるようにする。
    $script:CurrentLogFilePath = $null
    $script:InMailContext = $false
}

function Initialize-MailLog {
    param([string]$OutputDir)

    if (-not $CONFIG.EnableLogging) { return }

    # 保存フォルダ内に処理ログファイルを作成し、共通ログ+そのメールのログを書き出したうえで
    # 以降は逐次追記に切り替える。フォルダへログを書けない場合はデスクトップへ退避する。
    try {
        $logFileName = "処理ログ_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        $logFilePath = Join-Path $OutputDir $logFileName
        $combined = @($script:CommonLogBuffer) + @($script:MailLogBuffer)
        $initialContent = if ($combined.Count -gt 0) { ($combined -join "`r`n") + "`r`n" } else { "" }
        [System.IO.File]::WriteAllText($logFilePath, $initialContent, [System.Text.Encoding]::UTF8)
        $script:CurrentLogFilePath = $logFilePath
        $script:MailLogBuffer = @()
    } catch {
        Write-Log "  処理ログの保存に失敗しました（保存先フォルダ: $OutputDir）: $_" -Level Warning
        Save-PendingLogToDesktop -Reason "保存フォルダへのログ書き込みに失敗(フォルダ: $OutputDir)"
    }
}

function Save-PendingLogToDesktop {
    param(
        [string]$Reason
    )

    if (-not $CONFIG.EnableLogging) {
        $script:MailLogBuffer = @()
        return
    }

    # 保存フォルダを確定できなかった場合の退避先。共通ログ+そのメールのログをまとめて書き出す。
    # 同時刻の衝突を避けるため連番を付与する。
    try {
        $baseName = "MailExporter_Error_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        $errorLogPath = Join-Path $CONFIG.DesktopPath "$baseName.txt"
        $suffix = 1
        while ([System.IO.File]::Exists($errorLogPath)) {
            $suffix++
            $errorLogPath = Join-Path $CONFIG.DesktopPath "${baseName}_$suffix.txt"
        }

        $header = @(
            "============================================"
            "MailExporter エラーログ"
            "発生日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            "理由: $Reason"
            "============================================"
        )
        $body = $header + @($script:CommonLogBuffer) + @($script:MailLogBuffer)
        [System.IO.File]::WriteAllText($errorLogPath, (($body -join "`r`n") + "`r`n"), [System.Text.Encoding]::UTF8)
        Write-Host "エラーログを保存しました: $errorLogPath" -ForegroundColor Yellow
    } catch {
        Write-Host "デスクトップへのエラーログ保存にも失敗しました: $_" -ForegroundColor Red
    } finally {
        # 共通ログは以降のメールでも引き続き使うため消去しない。そのメール分だけ消去する。
        $script:MailLogBuffer = @()
    }
}

function Write-ErrorDetail {
    param(
        [string]$MailSubject,
        [string]$Stage,
        [System.Management.Automation.ErrorRecord]$ErrorRecord,
        # 指定した場合、保存フォルダのログへ書けていなければデスクトップへ退避する
        [string]$FallbackReason
    )

    # エラー記録処理自体が例外で落ちて元のエラー情報を失わないよう、
    # 記録しやすい情報から順に出力しつつ全体を防御する
    try {
        Write-Log "  [エラー詳細]" -Level Error
        Write-Log "    発生日時: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Level Error
        Write-Log "    処理対象メール: $MailSubject" -Level Error
        Write-Log "    処理段階: $Stage" -Level Error

        $ex = $ErrorRecord.Exception
        Write-Log "    例外種別: $($ex.GetType().FullName)" -Level Error
        Write-Log "    エラー内容: $($ex.Message)" -Level Error

        $invocation = $ErrorRecord.InvocationInfo
        if ($invocation) {
            $lineText = if ($invocation.Line) { $invocation.Line.Trim() } else { "" }
            Write-Log "    発生位置: $($invocation.ScriptName):$($invocation.ScriptLineNumber) $lineText" -Level Error
        }

        if ($ErrorRecord.ScriptStackTrace) {
            Write-Log "    ScriptStackTrace: $($ErrorRecord.ScriptStackTrace)" -Level Error
        }
    } catch {
        Write-Host "エラー詳細の記録処理自体が失敗しました: $_" -ForegroundColor Red
    }

    if ($FallbackReason -and -not $script:CurrentLogFilePath) {
        Save-PendingLogToDesktop -Reason $FallbackReason
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
        '?' = '？'; '"' = '＂'; '<' = '＜'; '>' = '＞'; '|' = '｜'
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
        return $fullPath
    }

    Write-Log "  パス長制限のためフォルダ名を短縮" -Level Warning

    # 「yyyyMMdd_件名」の件名部分だけを、上限に収まる長さまで切り詰める。
    # 使用済みの文字数 = 保存先フォルダ + 区切りの"\" + 日付8文字 + 区切りの"_"
    #                    + 同名フォルダがあった場合に付く連番("_99"を想定して3文字)
    if ($folderName -match '^(\d{8})_(.+)$') {
        $datePrefix = $Matches[1]
        $subject = $Matches[2]
        $usedLength = $basePath.Length + 1 + $datePrefix.Length + 1 + 3
        $availableLength = $maxLength - $usedLength

        if ($availableLength -gt 10) {
            $truncatedSubject = $subject.Substring(0, [Math]::Min($availableLength, $subject.Length))
            return (Join-Path $basePath "${datePrefix}_${truncatedSubject}")
        }
    }

    # 件名を削っても収まらない場合は、日付だけのフォルダ名にする
    if ($folderName -match '^(\d{8})') {
        return (Join-Path $basePath $Matches[1])
    }

    return $fullPath
}

function Release-Ref {
    param($ref)
    
    if ($ref -ne $null) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ref) | Out-Null
        } catch {
            # 既に解放済みの場合は無視
        }
    }
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
    $outputDir = Get-SafeFolderPath -basePath $CONFIG.DesktopPath `
                                    -folderName "${dateStr}_${safeSubject}" `
                                    -maxLength $CONFIG.MaxPathLength

    # 同じ日付・件名のフォルダが既にある場合は上書きせず _2, _3 ... と連番を付ける
    $baseDir = $outputDir
    $counter = 1
    while ([System.IO.Directory]::Exists($outputDir) -or [System.IO.File]::Exists($outputDir)) {
        $counter++
        $outputDir = "${baseDir}_${counter}"
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
        $attName = "(不明)"
        try {
            $attName = $att.FileName
            if ([string]::IsNullOrEmpty($attName)) { continue }

            $saveName = Get-SafeFilename $attName -maxLength 200
            $savePath = Join-Path $outputDir $saveName

            # 同名の添付が複数ある場合は上書きせず _2, _3 ... と連番を付ける
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($saveName)
            $ext = [System.IO.Path]::GetExtension($saveName)
            $counter = 1
            while ([System.IO.File]::Exists($savePath)) {
                $counter++
                $saveName = "${baseName}_${counter}${ext}"
                $savePath = Join-Path $outputDir $saveName
            }

            $att.SaveAsFile($savePath)
            # 添付一覧(PDFフッター)には実際に保存したファイル名を載せる
            $attachmentNames += $saveName

        } catch {
            Write-Log "  添付エラー: $attName - $_" -Level Warning
        }
    }

    Write-Log "  添付保存: $($attachmentNames.Count) 件完了" -Level Success
    return $attachmentNames
}

# ============================================
# Word文書生成関数（Wordオブジェクトモデルで直接組み立てる）
# ============================================
# HTMLをWordに読み込ませる方式(Documents.Open)は、メールのHTML内容によっては
# Wordの変換フィルタが極端に長時間応答しなくなることが実機で確認されたため採用しない。
# 元メールのHTML装飾の完全な再現は目的とせず、件名・差出人・宛先・CC・受信日時・
# 本文・添付一覧という文字情報を、Documents.Add()で作成した新規文書へ
# Selection.TypeText/TypeParagraphで直接書き込むことで、安定した速度でPDF化する。

function ConvertTo-WordParagraphText {
    # TypeTextはChr(13)(`r)を含む文字列を渡すと、そこで改段落しつつ1回のCOM呼び出しで
    # まとめて挿入できる。行ごとにTypeText/TypeParagraphを呼び返す方式は、
    # 長文メールで往復回数が増えて遅くなるため採らない。
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return (($Text -replace "`r`n", "`r") -replace "`n", "`r")
}

function Write-MailHeaderLine {
    # ラベル部分だけ太字にして1行で書き込む(表は使わず、安定して確実なAPIのみで組み立てる)
    param($Selection, [string]$Label, [string]$Value)

    $safeValue = ($Value -replace '[\r\n]+', ' ')

    $Selection.Font.Bold = $true
    $Selection.TypeText("${Label}: ")
    $Selection.Font.Bold = $false
    $Selection.TypeText($safeValue)
    $Selection.TypeParagraph()
}

function New-MailWordDocument {
    # Word.Applicationインスタンス(呼び出し元で使い回している1つのインスタンス)上に
    # 新規文書を1つ作成し、指定様式の内容を書き込んで返す。呼び出し元でClose/Export/移動を行う。
    param(
        $WordApp,
        $Metadata,
        [string]$BodyText,
        [array]$AttachmentNames
    )

    $doc = $WordApp.Documents.Add()

    try {
        $doc.PageSetup.TopMargin = $WordApp.CentimetersToPoints(2.0)
        $doc.PageSetup.BottomMargin = $WordApp.CentimetersToPoints(2.0)
        $doc.PageSetup.LeftMargin = $WordApp.CentimetersToPoints(2.5)
        $doc.PageSetup.RightMargin = $WordApp.CentimetersToPoints(2.5)
    } catch {
        # 余白調整はレイアウト上の見た目の問題に留まるため、失敗しても変換自体は継続する
    }

    $sel = $WordApp.Selection
    $sel.Font.Name = 'Yu Gothic'
    $sel.Font.Size = 10.5
    $sel.Font.Bold = $false

    # 件名（本文より大きく太字の見出しとして扱う）
    $subjectLine = ($Metadata.Subject -replace '[\r\n]+', ' ')
    $sel.Font.Size = 15
    $sel.Font.Bold = $true
    $sel.TypeText($subjectLine)
    $sel.TypeParagraph()

    $sel.Font.Size = 10.5
    $sel.Font.Bold = $false
    $sel.TypeParagraph()

    # ヘッダー情報（差出人・宛先・CC・受信日時を上部に整理）
    Write-MailHeaderLine -Selection $sel -Label '差出人' -Value $Metadata.SenderInfo
    Write-MailHeaderLine -Selection $sel -Label '宛先' -Value $Metadata.ToAddr
    Write-MailHeaderLine -Selection $sel -Label 'CC' -Value $Metadata.CcAddr
    Write-MailHeaderLine -Selection $sel -Label '受信日時' -Value $Metadata.DateTimeStr

    # ヘッダー情報と本文の区切り
    $sel.TypeParagraph()
    $sel.Font.Size = 8
    $sel.TypeText([string]::new([char]0x2015, 50))
    $sel.TypeParagraph()
    $sel.Font.Size = 10.5
    $sel.TypeParagraph()

    # 本文（原則としてMailItem.Bodyをそのまま使用し、装飾再現より文字情報の保持を優先する）
    $bodyToWrite = if ([string]::IsNullOrEmpty($BodyText)) { "(本文なし)" } else { $BodyText }
    $sel.TypeText((ConvertTo-WordParagraphText $bodyToWrite))
    $sel.TypeParagraph()
    $sel.TypeParagraph()

    # 添付ファイル一覧（本文末尾）
    $sel.Font.Bold = $true
    $sel.Font.Size = 9.5
    $sel.TypeText("添付ファイル:")
    $sel.TypeParagraph()
    $sel.Font.Bold = $false
    $sel.Font.Size = 10

    if ($AttachmentNames.Count -gt 0) {
        $attachmentText = ($AttachmentNames | ForEach-Object { "・$_" }) -join "`r"
        $sel.TypeText((ConvertTo-WordParagraphText $attachmentText))
        $sel.TypeParagraph()
    } else {
        $sel.TypeText("(添付ファイルなし)")
        $sel.TypeParagraph()
    }

    return $doc
}

function Test-WordAppAlive {
    # Word.ApplicationのCOMが生きているかどうかの簡易確認。
    # 1通のメールでPDF変換に失敗した場合に、Word全体を再利用してよいか判断するために使う
    # (COM状態が壊れている場合にのみ以降の処理を安全に打ち切るため)。
    param($WordApp)

    try {
        $null = $WordApp.Documents.Count
        return $true
    } catch {
        return $false
    }
}

function Convert-MailToPdf {
    # 引数で渡されたWord.Applicationインスタンス(処理全体で使い回す1つのインスタンス)上に
    # 新規文書を作成してPDFへ変換する。Word自体の起動・終了はここでは行わない。
    param(
        $Metadata,
        [string]$BodyText,
        [array]$AttachmentNames,
        [string]$FinalPdfPath,
        $WordApp
    )

    # 完成したPDFのみを、利用者由来の長い/特殊文字を含み得る最終保存フォルダへ移動する。
    # 変換自体は %TEMP%\MailExporter\{GUID}\ の短い一時フォルダ内だけで完結させる。
    $tempWorkDir = Join-Path ([System.IO.Path]::GetTempPath()) "MailExporter\$([System.Guid]::NewGuid().ToString('N'))"
    $doc = $null

    try {
        [System.IO.Directory]::CreateDirectory($tempWorkDir) | Out-Null
        $tempPdfPath = Join-Path $tempWorkDir "mail.pdf"

        Write-Log "    Word文書生成開始" -Level Info
        $doc = New-MailWordDocument -WordApp $WordApp -Metadata $Metadata -BodyText $BodyText -AttachmentNames $AttachmentNames
        Write-Log "    Word文書生成完了" -Level Info

        Write-Log "    PDF出力開始" -Level Info
        $doc.ExportAsFixedFormat($tempPdfPath, 17)
        Write-Log "    PDF出力完了" -Level Info

        # 書き込み完了直後のファイルシステム反映待ちのため、ファイルができるまで一定回数だけ待つ
        $pdfSize = -1
        $pdfReady = $false
        for ($retryCount = 0; $retryCount -lt $CONFIG.PdfRetryCount; $retryCount++) {
            $pdfFile = Get-Item -LiteralPath $tempPdfPath -ErrorAction SilentlyContinue
            if ($pdfFile) {
                $pdfSize = $pdfFile.Length
                if ($pdfSize -gt $CONFIG.PdfMinFileSize) {
                    $pdfReady = $true
                    break
                }
            }
            Start-Sleep -Milliseconds $CONFIG.PdfRetryInterval
        }

        if (-not $pdfReady) {
            Write-Log "  PDF生成に失敗しました（出力ファイルが確認できません）" -Level Error
            Write-Log "  PDF出力ファイルの存在: $([System.IO.File]::Exists($tempPdfPath))" -Level Error
            Write-Log "  PDF出力ファイルサイズ: $pdfSize" -Level Error
            return $false
        }

        [System.IO.File]::Move($tempPdfPath, $FinalPdfPath)
        return $true

    } catch {
        Write-Log "  PDF変換エラー: $_" -Level Error
        return $false
    } finally {
        if ($doc) {
            try { $doc.Close(0) } catch {}
            Release-Ref $doc
        }
        try {
            if ([System.IO.Directory]::Exists($tempWorkDir)) {
                [System.IO.Directory]::Delete($tempWorkDir, $true)
            }
        } catch {
            Write-Log "  一時作業フォルダの削除に失敗しました（$tempWorkDir）: $_" -Level Warning
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
        $WordApp
    )

    # 前のメールのログ状態(バッファ/ログファイル)を引き継がない
    Start-MailLogContext

    $stage = "メタデータ取得"
    $subjectForLog = "(不明)"

    try {
        # メタデータ取得
        $metadata = Get-MailMetadata -mail $mail
        $subjectForLog = $metadata.Subject
        Write-Log "[$index/$total] 処理中: $($metadata.Subject)" -Level Info

        # フォルダ作成
        $stage = "フォルダ作成"
        $outputDir = New-MailFolder -dateStr $metadata.DateStr -subject $metadata.Subject

        # フォルダが確定したので、以降のログはこのフォルダ内へ保存する
        Initialize-MailLog -OutputDir $outputDir

        # 添付ファイル保存
        $stage = "添付ファイル保存"
        $attachmentNames = Save-MailAttachments -mail $mail -outputDir $outputDir

        # PDF変換（Word文書の作成〜変換は短い一時フォルダ内で完結させ、完成後に保存フォルダへ移動する）
        $stage = "PDF変換"
        $safeSender = Get-SafeFilename $metadata.SenderName -maxLength $CONFIG.MaxSenderLength
        $pdfName = "$($metadata.DateStr)_${safeSender}mail.pdf"
        $finalPdfPath = Join-Path $outputDir $pdfName

        Write-Log "  PDF変換中..." -Level Info

        $pdfResult = Convert-MailToPdf -Metadata $metadata -BodyText $mail.Body -AttachmentNames $attachmentNames `
                                        -FinalPdfPath $finalPdfPath -WordApp $WordApp

        if ($pdfResult) {
            Write-Log "  PDF作成: $pdfName" -Level Success
        } else {
            Write-Log "  PDF生成失敗" -Level Error
            return $null
        }

        Write-Log "  完了 ✓" -Level Success
        Write-Log "--------------------------------------------" -Level Info

        return $outputDir

    } catch {
        # 保存フォルダのログに書けていればそこへ、書けていなければデスクトップへ退避する
        Write-ErrorDetail -MailSubject $subjectForLog -Stage $stage -ErrorRecord $_ `
            -FallbackReason "メール処理中にエラー(段階: $stage)"
        return $null
    } finally {
        # 次のメール(または呼び出し元)のログが、このメールのコンテキストへ混入しないようにする
        Stop-MailLogContext
    }
}

# ============================================
# メイン処理
# ============================================

$outlook = $null
$explorer = $null
$selection = $null
$word = $null
$processedCount = 0
$errorCount = 0
$skippedCount = 0
$processedFolders = @()

try {
    Write-Log "============================================" -Level Info
    Write-Log "Outlookメール保存ツール 開始" -Level Info
    Write-Log "============================================" -Level Info

    # Outlook接続
    try {
        $outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        Write-Log "Outlook接続: 成功" -Level Success
    } catch {
        Write-Log "Outlookが起動していません。" -Level Error
        Write-Log "Outlookを起動してメールを選択してから実行してください。" -Level Error
        Save-PendingLogToDesktop -Reason "起動時エラー: Outlookに接続できない"
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

    # Word起動（処理全体で1つのインスタンスを使い回す。ここで初めて実際に生成を試み、
    # 失敗した場合はここでエラーとする。以降メールごとにWordを起動・終了することはしない）
    try {
        $word = New-Object -ComObject Word.Application
    } catch {
        Write-Log "Microsoft Wordを起動できませんでした。" -Level Error
        Write-Log "Microsoft Wordがインストールされていることを確認してください。" -Level Error
        Save-PendingLogToDesktop -Reason "起動時エラー: Wordを起動できない"
        exit 1
    }
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # マクロ実行等の確認ダイアログが表示される余地を無くすため、強制的に無効化する(念のための対策)
    try { $word.AutomationSecurity = 3 } catch {}
    Write-Log "Word起動: 成功" -Level Success
    Write-Log "--------------------------------------------" -Level Info

    # メール処理ループ（各メールのログはそのメールの保存フォルダ内へ個別に保存される）
    $mailIndex = 0
    foreach ($mail in $selection) {
        $mailIndex++

        try {
            # 会議出席依頼や連絡先など、メール以外のアイテムは処理対象外(失敗ではない)
            # 実行全体の共通情報ではないため、後続メールの処理ログへ混入しないよう
            # 共通ログには積まず、コンソール表示のみとする
            if ($mail.Class -ne $CONFIG.OutlookMailItemClass) {
                Write-Host "[$mailIndex/$($selection.Count)] スキップ: メール以外のアイテム" -ForegroundColor Yellow
                $skippedCount++
                continue
            }

            $outputDir = Process-SingleMail -mail $mail -index $mailIndex -total $selection.Count -WordApp $word

            if ($outputDir) {
                $processedFolders += $outputDir
                $processedCount++
            } else {
                $errorCount++

                # PDF変換に失敗した場合のみ、Word自体が引き続き使える状態かを確認する。
                # COM状態が壊れていると判断できる場合だけ、以降のメール処理を安全に打ち切る
                # (単純な1通の変換失敗であれば、残りのメールの処理は継続する)。
                if (-not (Test-WordAppAlive -WordApp $word)) {
                    Write-Log "Wordとの接続が失われたため、以降のメール処理を中止します" -Level Error
                    break
                }
            }

        } catch {
            Write-ErrorDetail -MailSubject "(不明)" -Stage "メール処理(予期しないエラー)" -ErrorRecord $_ `
                -FallbackReason "メール処理中に予期しないエラー"
            $errorCount++

            if (-not (Test-WordAppAlive -WordApp $word)) {
                Write-Log "Wordとの接続が失われたため、以降のメール処理を中止します" -Level Error
                break
            }
        }
    }

    # 処理完了サマリー（ログは各メールの保存フォルダ内に個別にあるため、ここはコンソール表示のみ）
    Stop-MailLogContext
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "処理完了" -ForegroundColor Green
    Write-Host "  成功: $processedCount 件" -ForegroundColor Green
    Write-Host "  失敗: $errorCount 件" -ForegroundColor $(if ($errorCount -gt 0) { 'Yellow' } else { 'Green' })
    if ($skippedCount -gt 0) {
        Write-Host "  スキップ(メール以外): $skippedCount 件" -ForegroundColor Yellow
    }
    Write-Host "============================================" -ForegroundColor Green

    # フォルダ自動オープン
    if ($CONFIG.OpenFolderAfterProcess -and $processedFolders.Count -gt 0) {
        Start-Sleep -Milliseconds 500
        if ($processedFolders.Count -eq 1) {
            Invoke-Item -LiteralPath $processedFolders[0]
        } else {
            Invoke-Item -LiteralPath $CONFIG.DesktopPath
        }
    }

} catch {
    # メール処理ループの外で起きた予期しないエラー。特定のメールの保存フォルダに紐づかないため
    # 必ずデスクトップへ退避する。ここまでに蓄積した共通ログ(Outlook接続確認・選択メール件数等)を
    # 消さずに残すため、CurrentLogFilePathだけを念のため無効化する(Start-MailLogContextは呼ばない)。
    $script:CurrentLogFilePath = $null
    Write-ErrorDetail -MailSubject "(不明)" -Stage "スクリプト全体(予期しないエラー)" -ErrorRecord $_ `
        -FallbackReason "スクリプト全体で予期しないエラー"
    exit 1
} finally {
    # Word終了処理（開いているDocumentがあれば閉じてからQuitし、COMを解放する。
    # mailexporter自身がNew-Object -ComObjectで生成したインスタンスのみを対象とし、
    # 利用者が別途起動しているWordには一切触れない）
    if ($word) {
        try {
            $openDocs = @($word.Documents)
            foreach ($d in $openDocs) {
                try { $d.Close(0) } catch {}
                Release-Ref $d
            }
        } catch {}
        try { $word.Quit() } catch {}
        Release-Ref $word
    }

    # COM解放
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
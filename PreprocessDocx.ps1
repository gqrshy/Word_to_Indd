<#
.SYNOPSIS
    Word文書をInDesignインポート用に前処理するスクリプト

.DESCRIPTION
    以下の処理を行い、InDesignでのクラッシュを防ぎます：
    - SVG拡張形式をPNGフォールバックのみに変換
    - 変更履歴（Track Changes）を確定
    - コメントを削除
    - テキストボックスの問題要素を簡略化

.PARAMETER InputFile
    入力するdocxファイルのパス

.PARAMETER OutputFile
    出力するdocxファイルのパス（省略時は _cleaned を付加）

.EXAMPLE
    .\PreprocessDocx.ps1 -InputFile "Chapter7.docx"
    .\PreprocessDocx.ps1 -InputFile "Chapter7.docx" -OutputFile "Chapter7_fixed.docx"
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [Parameter(Mandatory=$false)]
    [string]$OutputFile
)

# 出力ファイル名が指定されていない場合は自動生成
if (-not $OutputFile) {
    $directory = [System.IO.Path]::GetDirectoryName($InputFile)
    $filename = [System.IO.Path]::GetFileNameWithoutExtension($InputFile)
    $extension = [System.IO.Path]::GetExtension($InputFile)
    $OutputFile = [System.IO.Path]::Combine($directory, "${filename}_cleaned${extension}")
}

# 入力ファイルの存在確認
if (-not (Test-Path $InputFile)) {
    Write-Error "入力ファイルが見つかりません: $InputFile"
    exit 1
}

# 絶対パスに変換
$InputFile = [System.IO.Path]::GetFullPath($InputFile)
$OutputFile = [System.IO.Path]::GetFullPath($OutputFile)

Write-Host "=== Word文書前処理スクリプト ===" -ForegroundColor Cyan
Write-Host "入力: $InputFile"
Write-Host "出力: $OutputFile"
Write-Host ""

# .NET のZIP機能を読み込み
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 一時フォルダを作成
$tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.Guid]::NewGuid().ToString())
New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

Write-Host "[1/5] docxファイルを解凍中..." -ForegroundColor Yellow

try {
    # docxファイルを解凍
    [System.IO.Compression.ZipFile]::ExtractToDirectory($InputFile, $tempDir)

    $documentXmlPath = [System.IO.Path]::Combine($tempDir, "word", "document.xml")
    $relsPath = [System.IO.Path]::Combine($tempDir, "word", "_rels", "document.xml.rels")

    if (-not (Test-Path $documentXmlPath)) {
        throw "document.xml が見つかりません。有効なdocxファイルではない可能性があります。"
    }

    # XMLを読み込み（BOMなしUTF-8で読み込む）
    $xmlContent = [System.IO.File]::ReadAllText($documentXmlPath, [System.Text.Encoding]::UTF8)

    # 統計情報
    $stats = @{
        SvgRemoved = 0
        DelRemoved = 0
        InsProcessed = 0
        CommentsRemoved = 0
        TextBoxSimplified = 0
    }

    Write-Host "[2/5] SVG拡張形式を処理中..." -ForegroundColor Yellow

    # SVG拡張（asvg:svgBlip）を含むa:ext要素を削除
    # パターン: <a:ext uri="{96DAC541-...SVG...}">...<asvg:svgBlip.../></a:ext>
    $svgExtPattern = '<a:ext\s+uri="\{96DAC541-7B7A-43D3-8B79-37D633B846F1\}"[^>]*>.*?</a:ext>'
    $matches = [regex]::Matches($xmlContent, $svgExtPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $stats.SvgRemoved = $matches.Count
    $xmlContent = [regex]::Replace($xmlContent, $svgExtPattern, '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    # 空になったa:extLst要素を削除
    $xmlContent = [regex]::Replace($xmlContent, '<a:extLst>\s*</a:extLst>', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    Write-Host "  - $($stats.SvgRemoved) 個のSVG拡張を削除" -ForegroundColor Gray

    Write-Host "[3/5] 変更履歴を処理中..." -ForegroundColor Yellow

    # 削除マーク（w:del）を完全に削除（中身ごと削除）
    $delPattern = '<w:del\s[^>]*>.*?</w:del>'
    $matches = [regex]::Matches($xmlContent, $delPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $stats.DelRemoved = $matches.Count
    $xmlContent = [regex]::Replace($xmlContent, $delPattern, '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    # 挿入マーク（w:ins）のタグを削除し、中身を残す
    $insPattern = '<w:ins\s[^>]*>(.*?)</w:ins>'
    $matches = [regex]::Matches($xmlContent, $insPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $stats.InsProcessed = $matches.Count
    $xmlContent = [regex]::Replace($xmlContent, $insPattern, '$1', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    # 段落プロパティの変更履歴（w:pPrChange）を削除
    $pPrChangePattern = '<w:pPrChange\s[^>]*>.*?</w:pPrChange>'
    $xmlContent = [regex]::Replace($xmlContent, $pPrChangePattern, '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    # 文字プロパティの変更履歴（w:rPrChange）を削除
    $rPrChangePattern = '<w:rPrChange\s[^>]*>.*?</w:rPrChange>'
    $xmlContent = [regex]::Replace($xmlContent, $rPrChangePattern, '', [System.Text.RegularExpressions.RegexOptions]::Singleline)

    # rsidDel, rsidR等の属性を削除（オプション：ファイルサイズ削減）
    $xmlContent = [regex]::Replace($xmlContent, '\s+w:rsidDel="[^"]*"', '')

    Write-Host "  - $($stats.DelRemoved) 個の削除マークを除去" -ForegroundColor Gray
    Write-Host "  - $($stats.InsProcessed) 個の挿入マークを確定" -ForegroundColor Gray

    Write-Host "[4/5] コメントを処理中..." -ForegroundColor Yellow

    # コメント参照を削除
    $commentPatterns = @(
        '<w:commentRangeStart[^/]*/>'
        '<w:commentRangeStart[^>]*>[^<]*</w:commentRangeStart>'
        '<w:commentRangeEnd[^/]*/>'
        '<w:commentRangeEnd[^>]*>[^<]*</w:commentRangeEnd>'
        '<w:commentReference[^/]*/>'
        '<w:commentReference[^>]*>[^<]*</w:commentReference>'
    )

    foreach ($pattern in $commentPatterns) {
        $matches = [regex]::Matches($xmlContent, $pattern)
        $stats.CommentsRemoved += $matches.Count
        $xmlContent = [regex]::Replace($xmlContent, $pattern, '')
    }

    Write-Host "  - $($stats.CommentsRemoved) 個のコメント参照を削除" -ForegroundColor Gray

    # コメント関連ファイルを削除
    $commentFiles = @(
        "word/comments.xml"
        "word/commentsExtended.xml"
        "word/commentsExtensible.xml"
        "word/commentsIds.xml"
    )

    foreach ($file in $commentFiles) {
        $filePath = [System.IO.Path]::Combine($tempDir, $file)
        if (Test-Path $filePath) {
            Remove-Item $filePath -Force
            Write-Host "  - $file を削除" -ForegroundColor Gray
        }
    }

    Write-Host "[5/5] テキストボックスを簡略化中..." -ForegroundColor Yellow

    # mc:AlternateContent内のmc:Fallbackを優先（mc:Choiceを削除）
    # これにより、新しい形式のテキストボックスではなく互換形式が使われる
    $altContentPattern = '<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>.*?</mc:Choice>\s*(<mc:Fallback[^>]*>.*?</mc:Fallback>)\s*</mc:AlternateContent>'
    $matches = [regex]::Matches($xmlContent, $altContentPattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $stats.TextBoxSimplified = $matches.Count

    # Fallbackの中身だけを残す
    $xmlContent = [regex]::Replace(
        $xmlContent,
        $altContentPattern,
        { param($m)
            $fallback = $m.Groups[1].Value
            # mc:Fallbackタグ自体も削除して中身だけ返す
            $fallback -replace '<mc:Fallback[^>]*>(.*)</mc:Fallback>', '$1'
        },
        [System.Text.RegularExpressions.RegexOptions]::Singleline
    )

    Write-Host "  - $($stats.TextBoxSimplified) 個のAlternateContentを簡略化" -ForegroundColor Gray

    # 修正したXMLを保存
    [System.IO.File]::WriteAllText($documentXmlPath, $xmlContent, (New-Object System.Text.UTF8Encoding($false)))

    # Content_Types.xmlからコメント関連のエントリを削除
    $contentTypesPath = [System.IO.Path]::Combine($tempDir, "[Content_Types].xml")
    if (Test-Path $contentTypesPath) {
        $ctContent = [System.IO.File]::ReadAllText($contentTypesPath, [System.Text.Encoding]::UTF8)
        $ctContent = [regex]::Replace($ctContent, '<Override[^>]*comments[^>]*/>', '')
        [System.IO.File]::WriteAllText($contentTypesPath, $ctContent, (New-Object System.Text.UTF8Encoding($false)))
    }

    # document.xml.relsからコメント関連のリレーションを削除
    if (Test-Path $relsPath) {
        $relsContent = [System.IO.File]::ReadAllText($relsPath, [System.Text.Encoding]::UTF8)
        $relsContent = [regex]::Replace($relsContent, '<Relationship[^>]*comments[^>]*/>', '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        [System.IO.File]::WriteAllText($relsPath, $relsContent, (New-Object System.Text.UTF8Encoding($false)))
    }

    # 出力ファイルが既に存在する場合は削除
    if (Test-Path $OutputFile) {
        Remove-Item $OutputFile -Force
    }

    # 新しいdocxファイルとして再パッケージ
    Write-Host ""
    Write-Host "docxファイルを再パッケージ中..." -ForegroundColor Yellow
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputFile, [System.IO.Compression.CompressionLevel]::Optimal, $false)

    Write-Host ""
    Write-Host "=== 処理完了 ===" -ForegroundColor Green
    Write-Host "出力ファイル: $OutputFile" -ForegroundColor Green
    Write-Host ""
    Write-Host "処理サマリー:" -ForegroundColor Cyan
    Write-Host "  - SVG拡張削除: $($stats.SvgRemoved) 個"
    Write-Host "  - 削除マーク除去: $($stats.DelRemoved) 個"
    Write-Host "  - 挿入マーク確定: $($stats.InsProcessed) 個"
    Write-Host "  - コメント参照削除: $($stats.CommentsRemoved) 個"
    Write-Host "  - AlternateContent簡略化: $($stats.TextBoxSimplified) 個"
    Write-Host ""
    Write-Host "このファイルをInDesignでインポートしてください。" -ForegroundColor Cyan

} catch {
    Write-Error "エラーが発生しました: $_"
    exit 1
} finally {
    # 一時フォルダを削除
    if (Test-Path $tempDir) {
        Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

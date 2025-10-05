# MEDHOT 医薬品コードデータ ダウンロードスクリプト
# https://medhot.medd.jp/view_download から全件CSVをダウンロード

param(
    [string]$OutputDir = "data\raw"
)

# 出力ディレクトリの作成
$OutputDir = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# メタデータディレクトリ作成
$metaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $metaDir)) {
    New-Item -ItemType Directory -Path $metaDir -Force | Out-Null
}

Write-Host "=== MEDHOT 医薬品コードデータ ダウンロード ===" -ForegroundColor Cyan
Write-Host ""

# ダウンロードページのURL
$downloadPageUrl = "https://medhot.medd.jp/view_download"

try {
    # ダウンロードページを取得
    Write-Host "ダウンロードページを取得中..." -ForegroundColor Yellow
    $response = Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing
    
    # CSVファイルのリンクを抽出
    $csvLinks = $response.Links | Where-Object { 
        $_.href -match 'csv/A_\d{8}_[12]\.txt'
    }
    
    if ($csvLinks.Count -eq 0) {
        Write-Error "CSVファイルのリンクが見つかりませんでした"
        exit 1
    }
    
    Write-Host "見つかったファイル: $($csvLinks.Count) 件" -ForegroundColor Green
    
    # 日付を抽出（最初のファイルから）
    $dateString = ""
    if ($csvLinks[0].href -match 'A_(\d{8})_') {
        $dateString = $matches[1]
        Write-Host "データ日付: $dateString" -ForegroundColor Green
    }
    else {
        $dateString = Get-Date -Format "yyyyMMdd"
        Write-Host "日付が取得できませんでした。現在日付を使用: $dateString" -ForegroundColor Yellow
    }
    
    Write-Host ""
    
    # ダウンロードしたファイル情報を格納
    $downloadedFiles = @()
    
    # 各CSVファイルをダウンロード
    foreach ($link in $csvLinks) {
        $csvUrl = $link.href
        
        # セッションIDを含むURLの場合、クリーンアップ
        if ($csvUrl -match '^https?://') {
            # 絶対URLの場合
            $csvUrl = $csvUrl -replace ';jsessionid=[^/]*', ''
        }
        else {
            # 相対URLの場合
            $csvUrl = "https://medhot.medd.jp" + ($csvUrl -replace ';jsessionid=[^/]*', '')
        }
        
        # ファイル名を抽出
        if ($csvUrl -match 'csv/(A_\d{8}_[12]\.txt)') {
            $originalFileName = $matches[1]
        }
        else {
            Write-Warning "ファイル名が取得できませんでした: $csvUrl"
            continue
        }
        
        # ファイルタイプを判定
        if ($originalFileName -match '_1\.txt$') {
            $fileType = "販売名・調剤包装単位コード"
            $outputFileName = "medhot_hanbaime_chouzai_${dateString}.txt"
        }
        elseif ($originalFileName -match '_2\.txt$') {
            $fileType = "調剤・販売・元梱包装単位コード"
            $outputFileName = "medhot_chouzai_hanbai_motokon_${dateString}.txt"
        }
        else {
            $fileType = "不明"
            $outputFileName = "medhot_${dateString}_${originalFileName}"
        }
        
        $outputPath = Join-Path $OutputDir $outputFileName
        
        Write-Host "ダウンロード中: $fileType" -ForegroundColor Yellow
        Write-Host "  URL: $csvUrl" -ForegroundColor Gray
        Write-Host "  保存先: $outputFileName" -ForegroundColor Gray
        
        try {
            # ダウンロード実行
            Invoke-WebRequest -Uri $csvUrl -OutFile $outputPath -UseBasicParsing
            
            $fileInfo = Get-Item $outputPath
            $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
            
            Write-Host "  ✓ 完了: $fileSizeMB MB" -ForegroundColor Green
            Write-Host ""
            
            # ダウンロード情報を記録
            $downloadedFiles += [PSCustomObject]@{
                FileType     = $fileType
                OriginalName = $originalFileName
                SavedName    = $outputFileName
                Url          = $csvUrl
                Size         = $fileInfo.Length
                SizeMB       = $fileSizeMB
                DownloadDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            }
        }
        catch {
            Write-Error "ファイルのダウンロードに失敗しました: $csvUrl"
            Write-Error $_.Exception.Message
        }
    }
    
    # メタデータを保存
    if ($downloadedFiles.Count -gt 0) {
        $metadataPath = Join-Path $metaDir "medhot_download_history.csv"
        
        if (Test-Path $metadataPath) {
            $downloadedFiles | Export-Csv -Path $metadataPath -Append -NoTypeInformation -Encoding UTF8
        }
        else {
            $downloadedFiles | Export-Csv -Path $metadataPath -NoTypeInformation -Encoding UTF8
        }
        
        Write-Host "メタデータを保存: $metadataPath" -ForegroundColor Green
        
        # 最新のダウンロード情報をJSONでも保存
        $jsonMetadata = @{
            DownloadDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            DataDate     = $dateString
            SourceUrl    = $downloadPageUrl
            Files        = $downloadedFiles
            TotalFiles   = $downloadedFiles.Count
            TotalSizeMB  = ($downloadedFiles | Measure-Object -Property SizeMB -Sum).Sum
        }
        
        $jsonMetadataPath = Join-Path $metaDir "medhot_latest_download.json"
        $jsonMetadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonMetadataPath -Encoding UTF8
        
        Write-Host "JSON メタデータを保存: $jsonMetadataPath" -ForegroundColor Green
    }
    
    Write-Host ""
    Write-Host "=== ダウンロード完了 ===" -ForegroundColor Green
    Write-Host "  データ日付: $dateString" -ForegroundColor White
    Write-Host "  ダウンロード件数: $($downloadedFiles.Count) ファイル" -ForegroundColor White
    Write-Host "  合計サイズ: $(($downloadedFiles | Measure-Object -Property SizeMB -Sum).Sum) MB" -ForegroundColor White
    Write-Host "  保存先: $OutputDir" -ForegroundColor White
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

# 医薬品HOTコードマスター ダウンロードスクリプト
# MEDIS-DC 医薬品HOTコードマスターサイトから最新データをダウンロード

param(
    [string]$OutputDir = "data\raw"
)

# 出力ディレクトリの作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# メタデータディレクトリ作成
$metaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $metaDir)) {
    New-Item -ItemType Directory -Path $metaDir -Force | Out-Null
}

Write-Host "=== 医薬品HOTコードマスター ダウンロード ===" -ForegroundColor Cyan

# HOTコードマスターページのURL
$baseUrl = "https://www2.medis.or.jp/hcode/"

try {
    # ページを取得してダウンロードリンクを探す
    Write-Host "HOTコードマスターページを取得中..." -ForegroundColor Yellow
    $response = Invoke-WebRequest -Uri $baseUrl -UseBasicParsing
    
    # ZIPファイルのリンクを抽出（全件マスター）
    $zipLinks = $response.Links | Where-Object { 
        $_.href -match 'moto_data/h\d{8}\.zip$' 
    }
    
    if ($zipLinks.Count -eq 0) {
        Write-Error "HOTコードマスターのZIPファイルが見つかりませんでした"
        exit 1
    }
    
    # 最新のZIPファイルを取得（日付が最も新しいもの）
    $latestZip = $zipLinks | Sort-Object { 
        if ($_.href -match 'h(\d{8})\.zip') { [int]$matches[1] } else { 0 }
    } -Descending | Select-Object -First 1
    
    $zipUrl = $latestZip.href
    if (-not $zipUrl.StartsWith("http")) {
        $zipUrl = "https://www2.medis.or.jp" + $zipUrl
    }
    
    # 日付を抽出
    if ($zipUrl -match 'h(\d{8})\.zip') {
        $dateString = $matches[1]
        Write-Host "最新版の日付: $dateString" -ForegroundColor Green
    }
    else {
        $dateString = (Get-Date -Format "yyyyMMdd")
    }
    
    # ダウンロード
    $zipFileName = "hotcode_master_$dateString.zip"
    $zipPath = Join-Path $OutputDir $zipFileName
    
    Write-Host "ダウンロード中: $zipUrl" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath
    
    $zipInfo = Get-Item $zipPath
    Write-Host "  ダウンロード完了: $($zipInfo.Length) bytes" -ForegroundColor Green
    
    # ZIPファイルを展開
    Write-Host "ZIPファイルを展開中..." -ForegroundColor Yellow
    $extractDir = Join-Path $OutputDir "hotcode_extract_$dateString"
    
    if (Test-Path $extractDir) {
        Remove-Item $extractDir -Recurse -Force
    }
    
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
    
    # 展開されたファイルを確認（サブディレクトリ含む）
    $extractedFiles = Get-ChildItem -Path $extractDir -File -Recurse
    Write-Host "  展開完了: $($extractedFiles.Count) ファイル" -ForegroundColor Green
    
    foreach ($file in $extractedFiles) {
        Write-Host "    - $($file.Name) ($($file.Length) bytes)" -ForegroundColor White
        
        # TXTファイルを直接data\rawにコピー（ファイル名を保持）
        if ($file.Extension -eq ".txt" -or $file.Extension -eq ".TXT") {
            # 元のファイル名から拡張子を除いて日付を追加
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $destFileName = "${baseName}_${dateString}.txt"
            $destPath = Join-Path $OutputDir $destFileName
            Copy-Item $file.FullName -Destination $destPath -Force
            Write-Host "      → $destPath にコピー" -ForegroundColor Cyan
        }
    }
    
    # メタデータを保存
    $metadata = @{
        DownloadDate   = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
        SourceUrl      = $zipUrl
        DataDate       = $dateString
        ZipFile        = $zipFileName
        ZipSize        = $zipInfo.Length
        ExtractedFiles = $extractedFiles.Count
        Files          = $extractedFiles | ForEach-Object { 
            @{
                Name      = $_.Name
                Size      = $_.Length
                Extension = $_.Extension
            }
        }
    }
    
    $metadataPath = Join-Path $metaDir "hotcode_download_history.csv"
    $metadataEntry = [PSCustomObject]@{
        DownloadDate   = $metadata.DownloadDate
        SourceUrl      = $metadata.SourceUrl
        DataDate       = $metadata.DataDate
        ZipFile        = $metadata.ZipFile
        ZipSize        = $metadata.ZipSize
        ExtractedFiles = $metadata.ExtractedFiles
    }
    
    if (Test-Path $metadataPath) {
        $metadataEntry | Export-Csv -Path $metadataPath -Append -NoTypeInformation -Encoding UTF8
    }
    else {
        $metadataEntry | Export-Csv -Path $metadataPath -NoTypeInformation -Encoding UTF8
    }
    
    Write-Host "`nメタデータを保存: $metadataPath" -ForegroundColor Green
    
    # サマリをJSON形式でも保存
    $jsonMetadataPath = Join-Path $metaDir "hotcode_latest_download.json"
    $metadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $jsonMetadataPath -Encoding UTF8
    
    Write-Host "`n=== ダウンロード完了 ===" -ForegroundColor Green
    Write-Host "  ZIPファイル: $zipPath" -ForegroundColor White
    Write-Host "  展開ディレクトリ: $extractDir" -ForegroundColor White
    Write-Host "  データ日付: $dateString" -ForegroundColor White
}
catch {
    Write-Error "ダウンロード中にエラーが発生しました: $($_.Exception.Message)"
    exit 1
}

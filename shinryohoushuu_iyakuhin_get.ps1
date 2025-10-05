# 診療報酬システム医薬品マスター取得スクリプト
# 厚生労働省の診療報酬システムから医薬品マスターCSVファイルを取得する

param(
    [string]$OutputDir = "data\raw",
    [string]$BaseUrl = "https://shinryohoshu.mhlw.go.jp/shinryohoshu/downloadMenu/",
    [switch]$ForceUpdate,     # 更新チェックをスキップして強制ダウンロード
    [switch]$CheckOnly,       # 更新チェックのみ実行（ダウンロードしない）
    [switch]$CleanupDuplicates, # 重複ファイルを整理
    [switch]$FixFileFormats,  # 誤った拡張子のファイルを修正
    [switch]$ShowHelp
)

# ヘルプ表示
if ($ShowHelp) {
    Write-Host @"
診療報酬システム医薬品マスター取得スクリプト

使用方法:
  .\shinryohoushuu_iyakuhin_get.ps1 [オプション]

オプション:
  -OutputDir <path>      出力ディレクトリ (デフォルト: data\raw)
  -BaseUrl <url>         取得元URL (デフォルト: 診療報酬システムダウンロードページ)
  -ForceUpdate           更新チェックをスキップして強制ダウンロード
  -CheckOnly             更新チェックのみ実行（ダウンロードしない）
  -CleanupDuplicates     重複ファイルを整理（古い重複ファイルを削除）
  -FixFileFormats       誤った拡張子のファイルを修正（.csvなのにZIPなど）
  -ShowHelp              このヘルプを表示

例:
  .\shinryohoushuu_iyakuhin_get.ps1                    # 標準設定で実行（更新時のみダウンロード）
  .\shinryohoushuu_iyakuhin_get.ps1 -ForceUpdate       # 強制ダウンロード
  .\shinryohoushuu_iyakuhin_get.ps1 -CheckOnly         # 更新チェックのみ
  .\shinryohoushuu_iyakuhin_get.ps1 -CleanupDuplicates # 重複ファイル整理
  .\shinryohoushuu_iyakuhin_get.ps1 -FixFileFormats    # ファイル形式修正
  .\shinryohoushuu_iyakuhin_get.ps1 -OutputDir "csv"   # 出力先指定

取得されるファイル:
  - 医薬品マスター (CSV形式) → data/raw/
  - ダウンロードメタデータ (CSV形式) → data/raw/meta/
"@ -ForegroundColor White
    exit 0
}

# 出力ディレクトリの作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# メタデータ専用ディレクトリの作成
$metaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $metaDir)) {
    New-Item -ItemType Directory -Path $metaDir -Force | Out-Null
    Write-Host "メタデータディレクトリを作成しました: $metaDir" -ForegroundColor Green
}

# 統合メタデータファイルのパス（metaフォルダ内）
$masterMetadataPath = Join-Path $metaDir "iyakuhin_master_history.csv"

# 旧パスから新パスへの移行処理
$oldMasterMetadataPath = Join-Path $OutputDir "iyakuhin_master_history.csv"
if ((Test-Path $oldMasterMetadataPath) -and (-not (Test-Path $masterMetadataPath))) {
    Write-Host "統合メタデータファイルをmetaフォルダに移動中..." -ForegroundColor Yellow
    Move-Item -Path $oldMasterMetadataPath -Destination $masterMetadataPath
    Write-Host "移動完了: $masterMetadataPath" -ForegroundColor Green
}

# 既存の個別メタデータファイルもmetaフォルダに移動
$oldMetadataFiles = Get-ChildItem -Path $OutputDir -Filter "iyakuhin_master_metadata_*.csv"
if ($oldMetadataFiles.Count -gt 0) {
    Write-Host "既存の個別メタデータファイルをmetaフォルダに移動中..." -ForegroundColor Yellow
    foreach ($file in $oldMetadataFiles) {
        $newPath = Join-Path $metaDir $file.Name
        Move-Item -Path $file.FullName -Destination $newPath
        Write-Host "  移動: $($file.Name)" -ForegroundColor Green
    }
}

# 既存の統合メタデータを読み込み
$existingMetadata = @()
if (Test-Path $masterMetadataPath) {
    try {
        $existingMetadata = Import-Csv -Path $masterMetadataPath -Encoding UTF8
        Write-Host "既存のメタデータを読み込みました: $($existingMetadata.Count) 件" -ForegroundColor Green
    }
    catch {
        Write-Warning "メタデータファイルの読み込みに失敗しました: $_"
    }
}

# 重複ファイル整理機能
function Remove-DuplicateFiles {
    param(
        [string]$OutputDir,
        [string]$MetaDir
    )
    
    Write-Host "`n重複ファイルをチェック中..." -ForegroundColor Yellow
    
    # 医薬品マスターファイルを取得
    $csvFiles = Get-ChildItem -Path $OutputDir -Filter "iyakuhin_master_*.csv" | Sort-Object LastWriteTime -Descending
    
    if ($csvFiles.Count -le 1) {
        Write-Host "重複ファイルは見つかりませんでした" -ForegroundColor Green
        return
    }
    
    Write-Host "医薬品マスターファイルが $($csvFiles.Count) 個見つかりました:" -ForegroundColor Cyan
    $csvFiles | ForEach-Object {
        $sizeKB = [math]::Round($_.Length / 1024, 1)
        Write-Host "  $($_.Name) - ${sizeKB} KB - $($_.LastWriteTime)" -ForegroundColor White
    }
    
    # ファイルサイズで重複をグループ化
    $sizeGroups = $csvFiles | Group-Object Length
    $duplicatesFound = $false
    
    foreach ($group in $sizeGroups) {
        if ($group.Count -gt 1) {
            $duplicatesFound = $true
            $sizeKB = [math]::Round([long]$group.Name / 1024, 1)
            Write-Host "`n同じサイズ (${sizeKB} KB) のファイルが $($group.Count) 個見つかりました:" -ForegroundColor Yellow
            
            # 最新のファイルを保持、古いファイルを削除対象とする
            $filesToKeep = $group.Group | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            $filesToDelete = $group.Group | Sort-Object LastWriteTime -Descending | Select-Object -Skip 1
            
            Write-Host "保持: $($filesToKeep.Name)" -ForegroundColor Green
            
            foreach ($file in $filesToDelete) {
                Write-Host "削除対象: $($file.Name)" -ForegroundColor Red
                
                # 確認
                $response = Read-Host "削除しますか？ (y/N)"
                if ($response -eq 'y' -or $response -eq 'Y') {
                    Remove-Item -Path $file.FullName -Force
                    Write-Host "  削除完了: $($file.Name)" -ForegroundColor Green
                    
                    # 対応するメタデータファイルも削除
                    $timestamp = ($file.BaseName -replace 'iyakuhin_master_', '')
                    $metadataFile = Join-Path $MetaDir "iyakuhin_master_metadata_$timestamp.csv"
                    if (Test-Path $metadataFile) {
                        Remove-Item -Path $metadataFile -Force
                        Write-Host "  関連メタデータも削除: iyakuhin_master_metadata_$timestamp.csv" -ForegroundColor Green
                    }
                }
                else {
                    Write-Host "  スキップ: $($file.Name)" -ForegroundColor Yellow
                }
            }
        }
    }
    
    if (-not $duplicatesFound) {
        Write-Host "重複ファイルは見つかりませんでした" -ForegroundColor Green
    }
}

# ファイル形式修正機能
function Fix-FileFormats {
    param(
        [string]$OutputDir,
        [string]$MetaDir
    )
    
    Write-Host "`nファイル形式をチェック中..." -ForegroundColor Yellow
    
    # 医薬品マスターファイルを取得
    $csvFiles = Get-ChildItem -Path $OutputDir -Filter "iyakuhin_master_*.csv"
    
    if ($csvFiles.Count -eq 0) {
        Write-Host "医薬品マスターファイルが見つかりませんでした" -ForegroundColor Green
        return
    }
    
    $fixedCount = 0
    
    foreach ($file in $csvFiles) {
        Write-Host "チェック中: $($file.Name)" -ForegroundColor Cyan
        
        # ファイルの先頭4バイトを読み取り
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)[0..3]
        $magicNumber = ($bytes | ForEach-Object { [System.Convert]::ToString($_, 16).PadLeft(2, '0') }) -join ' '
        
        # ZIPファイルのマジックナンバーをチェック
        if ($magicNumber -eq '50 4b 03 04') {
            Write-Host "  → ZIPファイルを検出: $($file.Name)" -ForegroundColor Yellow
            
            # タイムスタンプを取得
            $timestamp = ($file.BaseName -replace 'iyakuhin_master_', '')
            
            # 新しいファイル名を生成
            $zipFileName = "iyakuhin_master_$timestamp.zip"
            $zipPath = Join-Path $OutputDir $zipFileName
            
            Write-Host "  → .csvを.zipに変更中..." -ForegroundColor Yellow
            Move-Item -Path $file.FullName -Destination $zipPath
            
            # ZIPファイルを解凍
            Write-Host "  → ZIPファイルを解凍中..." -ForegroundColor Yellow
            $extractPath = Join-Path $OutputDir "extracted_$timestamp"
            Expand-Archive -Path $zipPath -DestinationPath $extractPath -Force
            
            # 解凍されたCSVファイルを確認
            $extractedCsvFiles = Get-ChildItem -Path $extractPath -Filter "*.csv"
            if ($extractedCsvFiles.Count -gt 0) {
                $extractedCsv = $extractedCsvFiles[0]
                $finalCsvPath = Join-Path $OutputDir "iyakuhin_master_$timestamp.csv"
                Copy-Item -Path $extractedCsv.FullName -Destination $finalCsvPath
                
                Write-Host "  → CSVファイルを抽出: $($extractedCsv.Name)" -ForegroundColor Green
                
                # 実際のレコード数を確認
                $csvContent = Get-Content -Path $finalCsvPath
                $actualRecordCount = $csvContent.Count - 1  # ヘッダー行を除く
                
                Write-Host "  → レコード数: $actualRecordCount 件" -ForegroundColor Green
                
                # 一時解凍フォルダを削除
                Remove-Item -Path $extractPath -Recurse -Force
                
                # 元のZIPファイルを削除
                Remove-Item -Path $zipPath -Force
                
                $fixedCount++
            }
            else {
                Write-Warning "  → ZIPファイル内にCSVファイルが見つかりませんでした"
                # 元の名前に戻す
                Move-Item -Path $zipPath -Destination $file.FullName
            }
        }
        else {
            Write-Host "  → 正常なCSVファイル" -ForegroundColor Green
        }
    }
    
    if ($fixedCount -gt 0) {
        Write-Host "`n$fixedCount 個のファイルを修正しました" -ForegroundColor Green
    }
    else {
        Write-Host "`n修正が必要なファイルは見つかりませんでした" -ForegroundColor Green
    }
}

# CleanupDuplicatesモードの場合は重複整理のみ実行
if ($CleanupDuplicates) {
    Remove-DuplicateFiles -OutputDir $OutputDir -MetaDir $metaDir
    Write-Host "`n重複ファイル整理が完了しました" -ForegroundColor Green
    exit 0
}

# FixFileFormatsモードの場合はファイル形式修正のみ実行
if ($FixFileFormats) {
    Fix-FileFormats -OutputDir $OutputDir -MetaDir $metaDir
    Write-Host "`nファイル形式修正が完了しました" -ForegroundColor Green
    exit 0
}

# ダウンロードページの内容を取得
Write-Host "診療報酬システムのダウンロードページを取得中..." -ForegroundColor Yellow
try {
    $webContent = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing
    Write-Host "ページの取得が完了しました" -ForegroundColor Green
}
catch {
    Write-Error "ページの取得に失敗しました: $_"
    exit 1
}

# 医薬品マスターのダウンロードリンクを検索
Write-Host "医薬品マスターのダウンロードリンクを検索中..." -ForegroundColor Yellow

# HTMLからテーブルデータを解析して医薬品マスターの情報を抽出
$htmlContent = $webContent.Content

# 医薬品マスター（yFile）のリンクを検索
$yFileLink = $webContent.Links | Where-Object { $_.href -match "/downloadMenu/yFile" } | Select-Object -First 1

if ($yFileLink) {
    $downloadPath = $yFileLink.href
    $downloadUrl = if ($downloadPath.StartsWith("http")) { 
        $downloadPath 
    }
    else { 
        "https://shinryohoshu.mhlw.go.jp" + $downloadPath 
    }
    Write-Host "ダウンロードURL発見: $downloadUrl" -ForegroundColor Green
}
else {
    Write-Error "医薬品マスターのダウンロードリンクが見つかりませんでした"
    exit 1
}

# 件数と更新日の情報を抽出（HTMLから）
$recordCount = "不明"
$updateDate = "不明"

# 医薬品マスターの行から件数と更新日を抽出
if ($htmlContent -match '医薬品マスター.*?(\d+,?\d*)\s*件') {
    $recordCount = $matches[1] -replace ',', ''
    Write-Host "件数発見: $recordCount 件" -ForegroundColor Green
}

# 医薬品マスターの行から更新日を抽出
if ($htmlContent -match '医薬品マスター.*?令和\s*(\d+)\s*年\s*(\d+)\s*月\s*(\d+)\s*日') {
    $year = [int]$matches[1] + 2018  # 令和年を西暦に変換
    $month = $matches[2]
    $day = $matches[3]
    $updateDate = "$year-$month-$day"
    Write-Host "更新日発見: $year年${month}月${day}日" -ForegroundColor Green
}

Write-Host "医薬品マスター情報:" -ForegroundColor Cyan
Write-Host "  URL: $downloadUrl" -ForegroundColor White
Write-Host "  件数: $recordCount 件" -ForegroundColor White
Write-Host "  最終更新: $updateDate" -ForegroundColor White

# 更新チェック機能
$needsUpdate = $true
$lastMetadata = $null

if (-not $ForceUpdate -and $existingMetadata.Count -gt 0) {
    Write-Host "`n更新チェック中..." -ForegroundColor Yellow
    
    # 最新のメタデータエントリを取得
    $lastMetadata = $existingMetadata | Sort-Object DownloadDate -Descending | Select-Object -First 1
    
    if ($lastMetadata) {
        Write-Host "前回ダウンロード情報:" -ForegroundColor Cyan
        Write-Host "  日時: $($lastMetadata.DownloadDate)" -ForegroundColor White
        Write-Host "  件数: $($lastMetadata.RecordCount) 件" -ForegroundColor White
        Write-Host "  更新日: $($lastMetadata.LastUpdated)" -ForegroundColor White
        
        # 更新の必要性をチェック
        $sameRecordCount = ($recordCount -eq "不明" -or $lastMetadata.RecordCount -eq $recordCount)
        $sameUpdateDate = ($updateDate -eq "不明" -or $lastMetadata.LastUpdated -eq $updateDate)
        
        if ($sameRecordCount -and $sameUpdateDate) {
            $needsUpdate = $false
            Write-Host "  → データに変更がありません" -ForegroundColor Green
        }
        else {
            Write-Host "  → データの更新が検出されました" -ForegroundColor Yellow
            if (-not $sameRecordCount) {
                Write-Host "    件数変更: $($lastMetadata.RecordCount) → $recordCount" -ForegroundColor Yellow
            }
            if (-not $sameUpdateDate) {
                Write-Host "    更新日変更: $($lastMetadata.LastUpdated) → $updateDate" -ForegroundColor Yellow
            }
        }
    }
}

# CheckOnlyモードの場合はここで終了
if ($CheckOnly) {
    if ($needsUpdate) {
        Write-Host "`n結果: 更新が利用可能です" -ForegroundColor Yellow
        exit 1  # 更新ありの場合は exit code 1
    }
    else {
        Write-Host "`n結果: 最新データです" -ForegroundColor Green
        exit 0  # 更新なしの場合は exit code 0
    }
}

# 更新が不要な場合はダウンロードをスキップ
if (-not $needsUpdate) {
    Write-Host "`n最新のデータが既に存在するため、ダウンロードをスキップします" -ForegroundColor Green
    Write-Host "強制ダウンロードを行う場合は -ForceUpdate オプションを使用してください" -ForegroundColor Yellow
    exit 0
}

# ファイルをダウンロード
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$tempFileName = "iyakuhin_master_$timestamp.tmp"  # 一時ファイル名
$tempPath = Join-Path $OutputDir $tempFileName

Write-Host "`n医薬品マスターをダウンロード中..." -ForegroundColor Yellow
try {
    Write-Host "  ダウンロード中..." -ForegroundColor Cyan
    Invoke-WebRequest -Uri $downloadUrl -OutFile $tempPath -UseBasicParsing
    
    $fileInfo = Get-Item $tempPath
    Write-Host "  完了: $([math]::Round($fileInfo.Length / 1024, 1)) KB" -ForegroundColor Green
    
    # ファイルの実際の形式を判定
    Write-Host "  ファイル形式を判定中..." -ForegroundColor Cyan
    $fileBytes = [System.IO.File]::ReadAllBytes($tempPath) | Select-Object -First 4
    
    $actualExtension = ".csv"  # デフォルト
    $fileFormat = "CSV"
    
    if ($fileBytes.Count -ge 4 -and 
        $fileBytes[0] -eq 0x50 -and $fileBytes[1] -eq 0x4B -and 
        $fileBytes[2] -eq 0x03 -and $fileBytes[3] -eq 0x04) {
        $actualExtension = ".zip"
        $fileFormat = "ZIP"
    }
    
    # 正しい拡張子でファイル名を設定
    $outputFileName = "iyakuhin_master_$timestamp$actualExtension"
    $outputPath = Join-Path $OutputDir $outputFileName
    
    # 一時ファイルを正しい名前にリネーム
    Move-Item -Path $tempPath -Destination $outputPath
    
    Write-Host "  検出されたファイル形式: $fileFormat" -ForegroundColor Green
    Write-Host "  保存ファイル名: $outputFileName" -ForegroundColor Green
    
    # ファイル形式に応じた処理
    if ($actualExtension -eq ".zip") {
        # ZIPファイルを解凍
        Write-Host "  ZIPファイルを解凍中..." -ForegroundColor Cyan
        $extractPath = Join-Path $OutputDir "extracted_$timestamp"
        Expand-Archive -Path $outputPath -DestinationPath $extractPath -Force
        
        # 解凍されたCSVファイルを確認
        $csvFiles = Get-ChildItem -Path $extractPath -Filter "*.csv"
        if ($csvFiles.Count -gt 0) {
            $csvFile = $csvFiles[0]
            $finalCsvPath = Join-Path $OutputDir "iyakuhin_master_$timestamp.csv"
            Copy-Item -Path $csvFile.FullName -Destination $finalCsvPath
            
            Write-Host "  CSVファイルを抽出: $($csvFile.Name)" -ForegroundColor Green
            
            # 抽出したCSVの行数を確認
            $csvContent = Get-Content -Path $finalCsvPath
            $actualRecordCount = $csvContent.Count - 1  # ヘッダー行を除く
            
            Write-Host "  実際のレコード数: $actualRecordCount 件" -ForegroundColor Green
            
            # 一時解凍フォルダを削除
            Remove-Item -Path $extractPath -Recurse -Force
            
            # 元のZIPファイルを削除（CSVだけ残す）
            Remove-Item -Path $outputPath -Force
            
            $finalFileInfo = Get-Item $finalCsvPath
            $outputFileName = "iyakuhin_master_$timestamp.csv"
            $outputPath = $finalCsvPath
            $fileInfo = $finalFileInfo
            $recordCount = $actualRecordCount
        }
        else {
            Write-Warning "ZIPファイル内にCSVファイルが見つかりませんでした"
        }
    }
    else {
        # CSVファイルとして直接処理
        Write-Host "  CSVファイルとして処理中..." -ForegroundColor Cyan
        
        # CSVの行数を確認
        $csvContent = Get-Content -Path $outputPath
        $actualRecordCount = $csvContent.Count - 1  # ヘッダー行を除く
        
        Write-Host "  実際のレコード数: $actualRecordCount 件" -ForegroundColor Green
        
        $recordCount = $actualRecordCount
    }
    
    # ダウンロード情報をメタデータとして保存
    $metadata = [PSCustomObject]@{
        FileName        = $outputFileName
        FilePath        = $outputPath
        DownloadUrl     = $downloadUrl
        RecordCount     = $recordCount
        LastUpdated     = $updateDate
        DownloadDate    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        FileSize        = $fileInfo.Length
        FileSizeKB      = [math]::Round($fileInfo.Length / 1024, 1)
        OriginalCsvName = $csvFile.Name
        UpdateType      = if ($ForceUpdate) { "強制更新" } elseif ($lastMetadata) { "自動更新" } else { "初回取得" }
    }
    
    # 個別メタデータファイルを保存（metaフォルダ内）
    $metadataPath = Join-Path $metaDir "iyakuhin_master_metadata_$timestamp.csv"
    $metadata | Export-Csv -Path $metadataPath -NoTypeInformation -Encoding UTF8
    
    # 統合メタデータファイルに追加
    $allMetadata = @()
    if ($existingMetadata.Count -gt 0) {
        $allMetadata += $existingMetadata
    }
    $allMetadata += $metadata
    
    # 統合メタデータファイルを更新
    $allMetadata | Export-Csv -Path $masterMetadataPath -NoTypeInformation -Encoding UTF8
    Write-Host "統合メタデータを更新しました: $masterMetadataPath" -ForegroundColor Green
    
    Write-Host "`n=== ダウンロード完了 ===" -ForegroundColor Green
    Write-Host "ファイル名: $outputFileName" -ForegroundColor White
    Write-Host "保存先: $outputPath" -ForegroundColor White
    Write-Host "ファイルサイズ: $([math]::Round($fileInfo.Length / 1024, 1)) KB" -ForegroundColor White
    Write-Host "レコード数: $recordCount 件" -ForegroundColor White
    Write-Host "最終更新日: $updateDate" -ForegroundColor White
    Write-Host "更新種別: $($metadata.UpdateType)" -ForegroundColor White
    Write-Host "個別メタデータ: $metadataPath" -ForegroundColor White
    Write-Host "統合メタデータ: $masterMetadataPath" -ForegroundColor White
    
    # 変更の詳細表示
    if ($lastMetadata -and -not $ForceUpdate) {
        Write-Host "`n変更詳細:" -ForegroundColor Cyan
        if ($lastMetadata.RecordCount -ne $recordCount) {
            Write-Host "  レコード数: $($lastMetadata.RecordCount) → $recordCount" -ForegroundColor Yellow
        }
        if ($lastMetadata.LastUpdated -ne $updateDate) {
            Write-Host "  更新日: $($lastMetadata.LastUpdated) → $updateDate" -ForegroundColor Yellow
        }
        $sizeDiff = [math]::Round(($fileInfo.Length - [long]$lastMetadata.FileSize) / 1024, 1)
        if ($sizeDiff -ne 0) {
            $sign = if ($sizeDiff -gt 0) { "+" } else { "" }
            Write-Host "  ファイルサイズ: ${sign}$sizeDiff KB" -ForegroundColor Yellow
        }
    }
    
}
catch {
    Write-Error "ダウンロードに失敗しました: $_"
    exit 1
}

Write-Host "`nスクリプトが完了しました" -ForegroundColor Green
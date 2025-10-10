# HOTコードマスター CSV整理スクリプト
# MEDIS****.csvから必要なフィールドのみを抽出して新しいCSVを作成

param(
    [string]$InputDir = "raw",
    [string]$OutputDir = "csv"
)

Write-Host "=== HOTコードマスター CSV整理 ===" -ForegroundColor Cyan
Write-Host ""

# 入力ディレクトリの確認
$InputDir = Join-Path $PSScriptRoot $InputDir
if (-not (Test-Path $InputDir)) {
    Write-Error "入力ディレクトリが見つかりません: $InputDir"
    exit 1
}

# MEDIS****.csvファイルを検索（8桁数字のみ）
$inputFiles = Get-ChildItem -Path $InputDir -Filter "MEDIS*.csv" | Where-Object { $_.Name -match '^MEDIS\d{8}\.csv$' }

if ($inputFiles.Count -eq 0) {
    Write-Error "MEDIS****.csvファイルが見つかりません"
    Write-Host "先に hotcode_get.ps1 を実行してください。" -ForegroundColor Yellow
    exit 1
}

# 出力ディレクトリの作成
$OutputDir = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# 各ファイルを処理
foreach ($inputFile in $inputFiles) {
    Write-Host "処理中: $($inputFile.Name)" -ForegroundColor Yellow
    
    try {
        # UTF-8でCSVを読み込み
        $data = Import-Csv -Path $inputFile.FullName -Encoding UTF8
        
        Write-Host "  読み込み完了: $($data.Count) 行" -ForegroundColor Green
        
        # 必要なフィールドのみを抽出
        Write-Host "  必要なフィールドを抽出中..." -ForegroundColor Yellow
        
        $processedData = $data | Select-Object `
            @{Name='基準番号（HOTコード）'; Expression={$_.'基準番号（ＨＯＴコード）'}},
            @{Name='処方用番号（HOT7）'; Expression={$_.'処方用番号（ＨＯＴ７）'}},
            @{Name='薬価基準収載医薬品コード'; Expression={$_.'薬価基準収載医薬品コード'}},
            @{Name='個別医薬品コード'; Expression={$_.'個別医薬品コード'}},
            @{Name='販売名'; Expression={$_.'販売名'}}
        
        Write-Host "  抽出完了: 5 フィールド" -ForegroundColor Green
        
        # 出力ファイル名は入力ファイル名と同じ
        $outputPath = Join-Path $OutputDir $inputFile.Name
        
        # CSVとして保存
        Write-Host "  CSVファイルを保存中..." -ForegroundColor Yellow
        $processedData | Export-Csv -Path $outputPath -Encoding UTF8 -NoTypeInformation
        
        $fileInfo = Get-Item $outputPath
        $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        
        Write-Host "  ✓ 保存完了: $fileSizeMB MB" -ForegroundColor Green
        Write-Host ""
        
        # サンプルデータを表示
        Write-Host "=== サンプルデータ（最初の3行）===" -ForegroundColor Cyan
        $processedData | Select-Object -First 3 | Format-Table -AutoSize -Wrap
        
        Write-Host ""
        Write-Host "=== 処理完了 ===" -ForegroundColor Green
        Write-Host "  入力: $($inputFile.FullName)" -ForegroundColor White
        Write-Host "  出力: $outputPath" -ForegroundColor White
        Write-Host "  処理行数: $($processedData.Count) 行" -ForegroundColor White
    }
    catch {
        Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
        Write-Error $_.ScriptStackTrace
    }
}

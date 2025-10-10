# MEDHOT CSV整理スクリプト
# medhot.csvから必要なフィールドのみを抽出して新しいCSVを作成

param(
    [string]$InputFile = "raw\medhot.csv",
    [string]$OutputDir = "csv"
)

Write-Host "=== MEDHOT CSV整理 ===" -ForegroundColor Cyan
Write-Host ""

# 入力ファイルの確認
$InputFile = Join-Path $PSScriptRoot $InputFile
if (-not (Test-Path $InputFile)) {
    Write-Error "入力ファイルが見つかりません: $InputFile"
    Write-Host "先に medhot_get.ps1 を実行してください。" -ForegroundColor Yellow
    exit 1
}

# 出力ディレクトリの作成
$OutputDir = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# 出力ファイルパス
$OutputFile = Join-Path $OutputDir "medhot.csv"

try {
    Write-Host "CSVファイルを読み込み中..." -ForegroundColor Yellow
    
    # UTF-8でCSVを読み込み
    $data = Import-Csv -Path $InputFile -Encoding UTF8
    
    Write-Host "  読み込み完了: $($data.Count) 行" -ForegroundColor Green
    Write-Host ""
    
    # 必要なフィールドのみを抽出
    Write-Host "必要なフィールドを抽出中..." -ForegroundColor Yellow
    
    $processedData = $data | Select-Object `
    @{Name = '販売名'; Expression = { $_.販売名 } },
    @{Name = '薬価基準収載医薬品コード'; Expression = { $_.薬価コード } },
    @{Name = '包装形態'; Expression = { $_.包装形態 } },
    @{Name = '包装単位数'; Expression = { $_.包装単位数 } },
    @{Name = '包装単位数単位'; Expression = { $_.包装単位数単位 } },
    @{Name = '総数量数'; Expression = { $_.総数量数 } },
    @{Name = '総数量数単位'; Expression = { $_.総数量数単位 } },
    @{Name = '調剤包装単位コード'; Expression = { $_.調剤包装単位コード } },
    @{Name = '販売包装単位コード'; Expression = { $_.販売包装単位コード } },
    @{Name = '元梱包装単位コード'; Expression = { $_.元梱包装単位コード } }
    
    Write-Host "  抽出完了: 10 フィールド" -ForegroundColor Green
    Write-Host ""
    
    # CSVとして保存
    Write-Host "CSVファイルを保存中..." -ForegroundColor Yellow
    $processedData | Export-Csv -Path $OutputFile -Encoding UTF8 -NoTypeInformation
    
    $fileInfo = Get-Item $OutputFile
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    
    Write-Host "  ✓ 保存完了: $fileSizeMB MB" -ForegroundColor Green
    Write-Host ""
    
    # サンプルデータを表示
    Write-Host "=== サンプルデータ（最初の3行）===" -ForegroundColor Cyan
    $processedData | Select-Object -First 3 | Format-Table -AutoSize -Wrap
    
    Write-Host ""
    Write-Host "=== 処理完了 ===" -ForegroundColor Green
    Write-Host "  入力: $InputFile" -ForegroundColor White
    Write-Host "  出力: $OutputFile" -ForegroundColor White
    Write-Host "  処理行数: $($processedData.Count) 行" -ForegroundColor White
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

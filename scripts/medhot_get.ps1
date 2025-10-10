# MEDHOT 医薬品コードデータ ダウンロードスクリプト
# https://medhot.medd.jp/view_download から調剤・販売・元梱包装単位コードファイルをダウンロード

param(
    [string]$OutputDir = "raw"
)

# スクリプトの親ディレクトリ（プロジェクトルート）を取得
$ProjectRoot = Split-Path -Parent $PSScriptRoot

# 出力ディレクトリの作成
$OutputDir = Join-Path $ProjectRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

Write-Host "=== MEDHOT 医薬品コードデータ ダウンロード ===" -ForegroundColor Cyan
Write-Host ""

# ダウンロードページのURL
$downloadPageUrl = "https://medhot.medd.jp/view_download"

try {
    # ダウンロードページを取得
    Write-Host "ダウンロードページを取得中..." -ForegroundColor Yellow
    $response = Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing
    
    # CSVファイルのリンクを抽出（調剤・販売・元梱包装単位コードファイルのみ: A_YYYYMMDD_2.txt）
    $csvLink = $response.Links | Where-Object { 
        $_.href -match 'csv/A_\d{8}_2\.txt'
    } | Select-Object -First 1
    
    if (-not $csvLink) {
        Write-Error "調剤・販売・元梱包装単位コードファイルが見つかりませんでした"
        exit 1
    }
    
    # URLを構築
    $csvUrl = $csvLink.href
    if ($csvUrl -match '^https?://') {
        $csvUrl = $csvUrl -replace ';jsessionid=[^/]*', ''
    }
    else {
        $csvUrl = "https://medhot.medd.jp" + ($csvUrl -replace ';jsessionid=[^/]*', '')
    }
    
    # 日付を抽出
    $dateString = ""
    if ($csvUrl -match 'A_(\d{8})_2') {
        $dateString = $matches[1]
        Write-Host "データ日付: $dateString" -ForegroundColor Green
    }
    
    # 出力ファイルパス
    $outputPath = Join-Path $OutputDir "medhot.csv"
    $tempPath = Join-Path $OutputDir "medhot_temp.txt"
    
    Write-Host "ダウンロード中: 調剤・販売・元梱包装単位コード" -ForegroundColor Yellow
    Write-Host "  URL: $csvUrl" -ForegroundColor Gray
    Write-Host "  保存先: medhot.csv" -ForegroundColor Gray
    
    try {
        # 一時ファイルとしてダウンロード
        Invoke-WebRequest -Uri $csvUrl -OutFile $tempPath -UseBasicParsing
        
        # Shift-JISからUTF-8に変換
        Write-Host "  文字コード変換中 (Shift-JIS → UTF-8)..." -ForegroundColor Yellow
        $sjisEncoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
        $utf8Encoding = [System.Text.Encoding]::UTF8
        
        $content = [System.IO.File]::ReadAllText($tempPath, $sjisEncoding)
        [System.IO.File]::WriteAllText($outputPath, $content, $utf8Encoding)
        
        # 一時ファイルを削除
        Remove-Item $tempPath -Force
        
        $fileInfo = Get-Item $outputPath
        $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        
        Write-Host "  ✓ 完了: $fileSizeMB MB (UTF-8)" -ForegroundColor Green
    }
    catch {
        Write-Error "ファイルのダウンロードまたは変換に失敗しました: $($_.Exception.Message)"
        if (Test-Path $tempPath) {
            Remove-Item $tempPath -Force
        }
        exit 1
    }
    
    Write-Host ""
    Write-Host "=== ダウンロード完了 ===" -ForegroundColor Green
    Write-Host "  データ日付: $dateString" -ForegroundColor White
    Write-Host "  保存先: $outputPath" -ForegroundColor White
    Write-Host ""
    
    # 整理スクリプトを自動実行
    $processScript = Join-Path $PSScriptRoot "medhot_process.ps1"
    if (Test-Path $processScript) {
        Write-Host "CSV整理スクリプトを実行します..." -ForegroundColor Cyan
        & $processScript
    }
    else {
        Write-Warning "medhot_process.ps1 が見つかりませんでした。手動で実行してください。"
    }
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

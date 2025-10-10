# 医薬品HOTコードマスター ダウンロードスクリプト
# https://www2.medis.or.jp/hcode/ から全件マスターをダウンロード

param(
    [string]$OutputDir = "raw",
    [string]$TempDir = "temp"
)

Write-Host "=== 医薬品HOTコードマスター ダウンロード ===" -ForegroundColor Cyan
Write-Host ""

# 出力ディレクトリの作成
$OutputDir = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# 一時ディレクトリの作成
$TempDir = Join-Path $PSScriptRoot $TempDir
if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Path $TempDir -Force | Out-Null
}

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
    $zipPath = Join-Path $TempDir "hotcode_$dateString.zip"
    
    Write-Host "ダウンロード中: $zipUrl" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    
    $zipInfo = Get-Item $zipPath
    Write-Host "  ダウンロード完了: $([math]::Round($zipInfo.Length / 1MB, 2)) MB" -ForegroundColor Green
    Write-Host ""
    
    # ZIPファイルを展開
    Write-Host "ZIPファイルを展開中..." -ForegroundColor Yellow
    $extractDir = Join-Path $TempDir "hotcode_extract_$dateString"
    
    if (Test-Path $extractDir) {
        Remove-Item $extractDir -Recurse -Force
    }
    
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force
    
    # MEDIS****.TXTファイルを検索（MEDIS+8桁数字.TXTのみ、アンダースコア付きは除外）
    Write-Host "MEDIS****.TXTファイルを検索中..." -ForegroundColor Yellow
    $allMedisFiles = Get-ChildItem -Path $extractDir -Filter "MEDIS*.TXT" -Recurse -File
    $medisFiles = $allMedisFiles | Where-Object { $_.Name -match '^MEDIS\d{8}\.TXT$' }
    
    if ($medisFiles.Count -eq 0) {
        Write-Error "MEDIS****.TXTファイルが見つかりませんでした"
        exit 1
    }
    
    Write-Host "  見つかったファイル: $($medisFiles.Count) 件（除外: $($allMedisFiles.Count - $medisFiles.Count) 件）" -ForegroundColor Green
    Write-Host ""
    
    # 各TXTファイルをCSVとして保存（UTF-8変換）
    foreach ($file in $medisFiles) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $outputFileName = "$baseName.csv"
        $outputPath = Join-Path $OutputDir $outputFileName
        
        Write-Host "変換中: $($file.Name)" -ForegroundColor Yellow
        Write-Host "  保存先: $outputFileName" -ForegroundColor Gray
        
        try {
            # Shift-JISからUTF-8に変換
            $sjisEncoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
            $utf8Encoding = [System.Text.Encoding]::UTF8
            
            $content = [System.IO.File]::ReadAllText($file.FullName, $sjisEncoding)
            [System.IO.File]::WriteAllText($outputPath, $content, $utf8Encoding)
            
            $outputInfo = Get-Item $outputPath
            $outputSizeMB = [math]::Round($outputInfo.Length / 1MB, 2)
            
            Write-Host "  ✓ 完了: $outputSizeMB MB (UTF-8)" -ForegroundColor Green
        }
        catch {
            Write-Error "ファイル変換に失敗しました: $($file.Name)"
            Write-Error $_.Exception.Message
        }
    }
    
    Write-Host ""
    
    # 一時ファイルをクリーンアップ
    Write-Host "一時ファイルをクリーンアップ中..." -ForegroundColor Yellow
    Remove-Item $zipPath -Force
    Remove-Item $extractDir -Recurse -Force
    
    # 一時ディレクトリが空なら削除
    if ((Get-ChildItem $TempDir).Count -eq 0) {
        Remove-Item $TempDir -Force
    }
    
    Write-Host "  ✓ クリーンアップ完了" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "=== ダウンロード完了 ===" -ForegroundColor Green
    Write-Host "  データ日付: $dateString" -ForegroundColor White
    Write-Host "  変換ファイル数: $($medisFiles.Count) ファイル" -ForegroundColor White
    Write-Host "  保存先: $OutputDir" -ForegroundColor White
    Write-Host ""
    
    # 整理スクリプトを自動実行
    $processScript = Join-Path $PSScriptRoot "hotcode_process.ps1"
    if (Test-Path $processScript) {
        Write-Host "CSV整理スクリプトを実行します..." -ForegroundColor Cyan
        & $processScript
    }
    else {
        Write-Warning "hotcode_process.ps1 が見つかりませんでした。手動で実行してください。"
    }
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

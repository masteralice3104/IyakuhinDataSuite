# PMDA CSV整形スクリプト
# raw/pmda.csvから不要な文字を削除し、数値を正規化してcsvフォルダに保存

param(
    [string]$InputFile = "raw\pmda.csv",
    [string]$OutputDir = "csv"
)

# スクリプトの親ディレクトリ（プロジェクトルート）を取得
$ProjectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "=== PMDA CSV整形 ===" -ForegroundColor Cyan
Write-Host ""

# 入力ファイルの確認
$InputFile = Join-Path $ProjectRoot $InputFile
if (-not (Test-Path $InputFile)) {
    Write-Error "入力ファイルが見つかりません: $InputFile"
    Write-Host "先に pmda_get.ps1 を実行してください。" -ForegroundColor Yellow
    exit 1
}

# 出力ディレクトリの作成
$OutputDir = Join-Path $ProjectRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# 出力ファイルパス
$OutputFile = Join-Path $OutputDir "pmda.csv"

# 数値を正規化する関数
function Normalize-Size {
    param([string]$value)
    
    if ([string]::IsNullOrWhiteSpace($value)) {
        return ""
    }
    
    # 元の値を保持（デバッグ用）
    $original = $value
    
    # 全角スペースを半角に
    $value = $value -replace '　', ' '
    
    # 「約」を削除
    $value = $value -replace '約\s*', ''
    
    # 「mm」を削除
    $value = $value -replace 'mm', ''
    
    # 全角数字を半角に
    $value = $value -replace '[０-９]', { [char]([int]$_.Value - 0xFF10 + 0x0030) }
    
    # 「×」や「x」を含む場合は最初の数値のみ抽出（例: "7.9×4.1" → "7.9"）
    if ($value -match '([0-9]+\.?[0-9]*)\s*[×xX]\s*([0-9]+\.?[0-9]*)') {
        $value = $matches[1]
    }
    
    # 前後の空白を削除
    $value = $value.Trim()
    
    # 数値のみを抽出（小数点を含む）
    if ($value -match '^([0-9]+\.?[0-9]*)') {
        $value = $matches[1]
    }
    else {
        # 数値が抽出できない場合は空文字列
        return ""
    }
    
    return $value
}

try {
    Write-Host "CSVファイルを読み込み中..." -ForegroundColor Yellow
    
    # UTF-8でCSVを読み込み
    $data = Import-Csv -Path $InputFile -Encoding UTF8
    
    Write-Host "  読み込み完了: $($data.Count) 行" -ForegroundColor Green
    Write-Host ""
    
    # データを整形
    Write-Host "データを整形中..." -ForegroundColor Yellow
    
    $processedData = $data | ForEach-Object {
        [PSCustomObject]@{
            名称       = $_.名称
            個別医薬品コード = $_.個別医薬品コード
            長径       = Normalize-Size $_.'長径'
            短径       = Normalize-Size $_.'短径'
            直径       = Normalize-Size $_.'直径'
            厚さ       = Normalize-Size $_.'厚さ'
        }
    }
    
    Write-Host "  整形完了" -ForegroundColor Green
    Write-Host ""
    
    # 空のサイズ情報を持つ行を除外（すべてのサイズフィールドが空の場合）
    Write-Host "無効な行を除外中..." -ForegroundColor Yellow
    $validData = $processedData | Where-Object {
        -not ([string]::IsNullOrWhiteSpace($_.'長径') -and 
            [string]::IsNullOrWhiteSpace($_.'短径') -and 
            [string]::IsNullOrWhiteSpace($_.'直径') -and 
            [string]::IsNullOrWhiteSpace($_.'厚さ'))
    }
    
    $removedCount = $processedData.Count - $validData.Count
    Write-Host "  除外: $removedCount 行" -ForegroundColor Green
    Write-Host "  有効行数: $($validData.Count) 行" -ForegroundColor Green
    Write-Host ""
    
    # CSVとして保存
    Write-Host "CSVファイルを保存中..." -ForegroundColor Yellow
    $validData | Export-Csv -Path $OutputFile -Encoding UTF8 -NoTypeInformation
    
    $fileInfo = Get-Item $OutputFile
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    
    Write-Host "  ✓ 保存完了: $fileSizeMB MB" -ForegroundColor Green
    Write-Host ""
    
    # サンプルデータを表示（整形前後の比較）
    Write-Host "=== 整形後のサンプルデータ（最初の10件）===" -ForegroundColor Cyan
    $validData | Select-Object -First 10 | Format-Table -AutoSize
    
    Write-Host ""
    Write-Host "=== 処理完了 ===" -ForegroundColor Green
    Write-Host "  入力: $InputFile" -ForegroundColor White
    Write-Host "  出力: $OutputFile" -ForegroundColor White
    Write-Host "  処理行数: $($validData.Count) 行" -ForegroundColor White
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

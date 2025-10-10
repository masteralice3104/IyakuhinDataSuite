<#
.SYNOPSIS
個別医薬品コードをキーとして、GS1コードとPMDAサイズ情報をまとめたJSONを生成

.DESCRIPTION
MEDIS、MEDHOT、PMDAの各データを統合し、個別医薬品コード（YJコード）をキーとした
包括的な医薬品情報JSONを生成します。

各個別医薬品コードに対して以下の情報を含みます:
- 販売名
- 薬価基準収載医薬品コード
- GS1包装単位コード（調剤/販売/元梱）
- PMDAサイズ情報（長径/短径/直径/厚さ）

.EXAMPLE
.\generate_drug_json.ps1
# output/drug_data.json にJSON形式で出力

.EXAMPLE
.\generate_drug_json.ps1 -OutputPath "custom_output.json"
# カスタムパスに出力
#>

param(
    [string]$OutputPath = "output/drug_data.json"
)

# スクリプトの親ディレクトリ（プロジェクトルート）を取得
$ProjectRoot = Split-Path -Parent $PSScriptRoot

# エラー時に停止
$ErrorActionPreference = "Stop"

Write-Host "=== 個別医薬品コード統合JSONの生成 ===" -ForegroundColor Cyan
Write-Host ""

# 入力ファイルパス
$medisPath = Join-Path $ProjectRoot "csv/MEDIS20250930.csv"
$medhotPath = Join-Path $ProjectRoot "csv/medhot.csv"
$pmdaPath = Join-Path $ProjectRoot "csv/pmda.csv"

# ファイル存在チェック
$requiredFiles = @($medisPath, $medhotPath, $pmdaPath)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path $file)) {
        Write-Error "必要なファイルが見つかりません: $file"
        exit 1
    }
}

# 出力パスをプロジェクトルート基準に変換
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $ProjectRoot $OutputPath
}

# 出力ディレクトリ作成
$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# データ読み込み
Write-Host "データを読み込み中..." -ForegroundColor Yellow
Write-Host "  MEDIS データ: $medisPath" -ForegroundColor Gray
$medisData = Import-Csv -Path $medisPath -Encoding UTF8

Write-Host "  MEDHOT データ: $medhotPath" -ForegroundColor Gray
$medhotData = Import-Csv -Path $medhotPath -Encoding UTF8

Write-Host "  PMDA データ: $pmdaPath" -ForegroundColor Gray
$pmdaData = Import-Csv -Path $pmdaPath -Encoding UTF8

Write-Host "  読み込み完了" -ForegroundColor Green
Write-Host ""

# 個別医薬品コードベースのデータ構造を構築
Write-Host "データを統合中..." -ForegroundColor Yellow
$drugDataMap = @{}

# 1. MEDISデータから基本情報を取得
Write-Host "  [1/3] MEDIS データを処理中..." -ForegroundColor Gray
$processedCount = 0
foreach ($row in $medisData) {
    $yjcode = $row.'個別医薬品コード'
    $yakkakijun = $row.'薬価基準収載医薬品コード'
    $salesName = $row.'販売名'
    
    if (-not $yjcode) { continue }
    
    if (-not $drugDataMap.ContainsKey($yjcode)) {
        $drugDataMap[$yjcode] = @{
            個別医薬品コード     = $yjcode
            販売名          = $salesName
            薬価基準収載医薬品コード = @()
            GS1コード       = @{
                調剤包装単位 = @()
                販売包装単位 = @()
                元梱包装単位 = @()
            }
            PMDAサイズ情報    = $null
        }
    }
    
    # 薬価基準収載医薬品コードを追加（重複なし）
    if ($yakkakijun -and $drugDataMap[$yjcode].薬価基準収載医薬品コード -notcontains $yakkakijun) {
        $drugDataMap[$yjcode].薬価基準収載医薬品コード += $yakkakijun
    }
    
    $processedCount++
    if ($processedCount % 10000 -eq 0) {
        Write-Host "    処理中: $processedCount 件..." -ForegroundColor DarkGray
    }
}
Write-Host "    完了: $processedCount 件処理" -ForegroundColor Green

# 2. MEDHOTデータからGS1コードを追加
Write-Host "  [2/3] MEDHOT データを処理中..." -ForegroundColor Gray

# 高速化のため、薬価基準収載医薬品コード→個別医薬品コードのインデックスを作成
Write-Host "    インデックスを作成中..." -ForegroundColor DarkGray
$yakkakijunToYjcodeMap = @{}
foreach ($yjcode in $drugDataMap.Keys) {
    $salesName = $drugDataMap[$yjcode].販売名
    foreach ($yakkakijun in $drugDataMap[$yjcode].薬価基準収載医薬品コード) {
        $key = "$yakkakijun|$salesName"
        if (-not $yakkakijunToYjcodeMap.ContainsKey($key)) {
            $yakkakijunToYjcodeMap[$key] = @()
        }
        if ($yakkakijunToYjcodeMap[$key] -notcontains $yjcode) {
            $yakkakijunToYjcodeMap[$key] += $yjcode
        }
    }
}
Write-Host "    インデックス作成完了（$($yakkakijunToYjcodeMap.Count) キー）" -ForegroundColor Green

# GS1コードを収集
Write-Host "    GS1コードを収集中..." -ForegroundColor DarkGray
$gs1Updates = @{}
$processedCount = 0

foreach ($row in $medhotData) {
    $yakkakijun = $row.'薬価コード'
    $salesName = $row.'販売名'
    
    if (-not $yakkakijun) { continue }
    
    $key = "$yakkakijun|$salesName"
    $matchingYjcodes = $yakkakijunToYjcodeMap[$key]
    
    if ($matchingYjcodes) {
        foreach ($yjcode in $matchingYjcodes) {
            if (-not $gs1Updates.ContainsKey($yjcode)) {
                $gs1Updates[$yjcode] = @{
                    調剤 = [System.Collections.Generic.HashSet[string]]::new()
                    販売 = [System.Collections.Generic.HashSet[string]]::new()
                    元梱 = [System.Collections.Generic.HashSet[string]]::new()
                }
            }
            
            $chozai = $row.'調剤包装単位コード'
            if ($chozai) { [void]$gs1Updates[$yjcode].調剤.Add($chozai) }
            
            $hanbai = $row.'販売包装単位コード'
            if ($hanbai) { [void]$gs1Updates[$yjcode].販売.Add($hanbai) }
            
            $motokon = $row.'元梱包装単位コード'
            if ($motokon) { [void]$gs1Updates[$yjcode].元梱.Add($motokon) }
        }
    }
    
    $processedCount++
    if ($processedCount % 10000 -eq 0) {
        Write-Host "    処理中: $processedCount / $($medhotData.Count) 件..." -ForegroundColor DarkGray
    }
}

# drugDataMapに反映
Write-Host "    drugDataMapに反映中..." -ForegroundColor DarkGray
foreach ($yjcode in $gs1Updates.Keys) {
    $drugDataMap[$yjcode].GS1コード.調剤包装単位 = @($gs1Updates[$yjcode].調剤)
    $drugDataMap[$yjcode].GS1コード.販売包装単位 = @($gs1Updates[$yjcode].販売)
    $drugDataMap[$yjcode].GS1コード.元梱包装単位 = @($gs1Updates[$yjcode].元梱)
}

Write-Host "    完了: $processedCount 件処理" -ForegroundColor Green

# 3. PMDAデータからサイズ情報を追加
Write-Host "  [3/3] PMDA データを処理中..." -ForegroundColor Gray
$processedCount = 0
$matchedCount = 0
foreach ($row in $pmdaData) {
    $yjcode = $row.'個別医薬品コード'
    
    if (-not $yjcode) { continue }
    
    if ($drugDataMap.ContainsKey($yjcode)) {
        $sizeInfo = @{}
        
        if ($row.'長径') { $sizeInfo['長径'] = $row.'長径' }
        if ($row.'短径') { $sizeInfo['短径'] = $row.'短径' }
        if ($row.'直径') { $sizeInfo['直径'] = $row.'直径' }
        if ($row.'厚さ') { $sizeInfo['厚さ'] = $row.'厚さ' }
        
        if ($sizeInfo.Count -gt 0) {
            $drugDataMap[$yjcode].PMDAサイズ情報 = $sizeInfo
            $matchedCount++
        }
    }
    
    $processedCount++
    if ($processedCount % 1000 -eq 0) {
        Write-Host "    処理中: $processedCount 件..." -ForegroundColor DarkGray
    }
}
Write-Host "    完了: $processedCount 件処理（マッチ: $matchedCount 件）" -ForegroundColor Green

Write-Host ""
Write-Host "データ統合完了" -ForegroundColor Green
Write-Host ""

# 統計情報
$totalDrugs = $drugDataMap.Count
$withGS1 = ($drugDataMap.Values | Where-Object { 
        $_.GS1コード.調剤包装単位.Count -gt 0 -or 
        $_.GS1コード.販売包装単位.Count -gt 0 -or 
        $_.GS1コード.元梱包装単位.Count -gt 0 
    }).Count
$withPMDA = ($drugDataMap.Values | Where-Object { $null -ne $_.PMDAサイズ情報 }).Count

Write-Host "=== 統計情報 ===" -ForegroundColor Cyan
Write-Host "  総個別医薬品コード数: $totalDrugs" -ForegroundColor White
Write-Host "  GS1コード保有数: $withGS1" -ForegroundColor White
Write-Host "  PMDAサイズ情報保有数: $withPMDA" -ForegroundColor White
Write-Host ""

# JSON出力
Write-Host "JSON出力中: $OutputPath" -ForegroundColor Yellow

# PowerShell標準のJSONコンバーターを使用（Depthを十分に設定）
$jsonContent = $drugDataMap | ConvertTo-Json -Depth 10 -Compress:$false

# UTF8（BOMなし）で保存
[System.IO.File]::WriteAllText($OutputPath, $jsonContent, [System.Text.UTF8Encoding]::new($false))

Write-Host "  出力完了" -ForegroundColor Green
Write-Host ""

# サンプル表示（最初の3件）
Write-Host "=== サンプルデータ（最初の3件） ===" -ForegroundColor Cyan
$sampleKeys = $drugDataMap.Keys | Select-Object -First 3
foreach ($key in $sampleKeys) {
    $drug = $drugDataMap[$key]
    Write-Host ""
    Write-Host "個別医薬品コード: $key" -ForegroundColor Yellow
    Write-Host "  販売名: $($drug.販売名)" -ForegroundColor Gray
    Write-Host "  薬価基準収載医薬品コード: $($drug.薬価基準収載医薬品コード -join ', ')" -ForegroundColor Gray
    Write-Host "  GS1コード:" -ForegroundColor Gray
    Write-Host "    調剤包装単位: $($drug.GS1コード.調剤包装単位.Count) 件" -ForegroundColor DarkGray
    Write-Host "    販売包装単位: $($drug.GS1コード.販売包装単位.Count) 件" -ForegroundColor DarkGray
    Write-Host "    元梱包装単位: $($drug.GS1コード.元梱包装単位.Count) 件" -ForegroundColor DarkGray
    if ($drug.PMDAサイズ情報) {
        Write-Host "  PMDAサイズ情報: あり" -ForegroundColor Gray
        foreach ($key in $drug.PMDAサイズ情報.Keys) {
            Write-Host "    $key : $($drug.PMDAサイズ情報[$key])" -ForegroundColor DarkGray
        }
    }
    else {
        Write-Host "  PMDAサイズ情報: なし" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "=== 完了 ===" -ForegroundColor Green
Write-Host "  出力ファイル: $OutputPath"
Write-Host "  総レコード数: $totalDrugs"
Write-Host ""

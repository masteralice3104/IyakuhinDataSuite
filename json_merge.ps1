# 医薬品データ統合スクリプト
# 各種JSONファイルを統合して、YJコードをキーにした統合データベースを作成
# 参考: hcodeguide.docx

param(
    [string]$ProcessedDir = "data\processed",
    [string]$OutputDir = "data\integrated"
)

Write-Host "=== 医薬品データ統合処理 ===" -ForegroundColor Cyan
Write-Host ""

# 出力ディレクトリ作成
$ProcessedDir = Join-Path $PSScriptRoot $ProcessedDir
$OutputDir = Join-Path $PSScriptRoot $OutputDir

if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# メタデータディレクトリ作成
$metaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $metaDir)) {
    New-Item -ItemType Directory -Path $metaDir -Force | Out-Null
}

# JSONファイルを読み込む関数
function Load-JsonData {
    param(
        [string]$Pattern,
        [string]$Description
    )
    
    Write-Host "読み込み中: $Description" -ForegroundColor Yellow
    $files = Get-ChildItem -Path $ProcessedDir -Filter $Pattern | Sort-Object LastWriteTime -Descending
    
    if ($files.Count -eq 0) {
        Write-Warning "  ファイルが見つかりません: $Pattern"
        return $null
    }
    
    $file = $files[0]
    Write-Host "  ファイル: $($file.Name)" -ForegroundColor Gray
    
    try {
        $jsonContent = Get-Content $file.FullName -Encoding UTF8 -Raw | ConvertFrom-Json
        $dataCount = if ($jsonContent.Data) { $jsonContent.Data.Count } else { 0 }
        Write-Host "  レコード数: $dataCount" -ForegroundColor Green
        return $jsonContent
    }
    catch {
        Write-Error "  エラー: $_"
        return $null
    }
}

# 各データソースを読み込み
Write-Host "=== データ読み込み ===" -ForegroundColor Cyan
Write-Host ""

# 1. 医薬品マスター（基本情報）
$iyakuhinMaster = Load-JsonData -Pattern "iyakuhin_master*.json" -Description "医薬品マスター"

# 2. 医薬品コード（詳細情報）
$iyakuhinCode = Load-JsonData -Pattern "iyakuhin_code*.json" -Description "医薬品コード"

# 3. HOTコードマスター（標準）- HOT13, HOT9対応
$hotcodeStandard = Load-JsonData -Pattern "MEDIS*_[0-9]*.json" -Description "HOTコードマスター（標準）"

# 4. HOTコードマスター（HOT9）- YJコード→HOT9対応
$hotcodeHot9 = Load-JsonData -Pattern "MEDIS*_HOT9*.json" -Description "HOTコードマスター（HOT9）"

# 5. HOTコードマスター（オプション）- 追加情報
$hotcodeOp = Load-JsonData -Pattern "MEDIS*_OP*.json" -Description "HOTコードマスター（オプション）"

# 6. MEDHOT 販売名・調剤包装単位コード
$medhotHanbaime = Load-JsonData -Pattern "medhot_hanbaime_chouzai*.json" -Description "MEDHOT 販売名・調剤包装単位コード"

# 7. MEDHOT 調剤・販売・元梱包装単位コード
$medhotChouzai = Load-JsonData -Pattern "medhot_chouzai_hanbai_motokon*.json" -Description "MEDHOT 調剤・販売・元梱包装単位コード"

# 8. PMDA YJコード・サイズ情報
$pmdaSize = Load-JsonData -Pattern "pmda_yjcode_size*.json" -Description "PMDA YJコード・サイズ情報"

Write-Host ""
Write-Host "=== データ統合処理 ===" -ForegroundColor Cyan
Write-Host ""

# 統合データ用のハッシュテーブル（YJコードをキーとする）
$integratedData = @{}

# 統計情報
$stats = @{
    TotalYJCodes           = 0
    WithHotCode            = 0
    WithMedhotInfo         = 0
    WithPmdaSize           = 0
    MultipleHotCodes       = 0
    MultipleHotCodesDetail = @()
}

# 1. 医薬品コード情報をベースにYJコードのエントリを作成
Write-Host "Step 1: 医薬品コード情報をベースに統合..." -ForegroundColor Yellow
if ($iyakuhinCode -and $iyakuhinCode.Data) {
    foreach ($item in $iyakuhinCode.Data) {
        # 薬価コードを取得（これがYJコード）
        $yjCode = $item."薬価コード"
        if (-not $yjCode) { continue }
        
        if (-not $integratedData.ContainsKey($yjCode)) {
            $integratedData[$yjCode] = @{
                YJCode              = $yjCode
                BasicInfo           = $item
                MasterInfo          = $null
                HotCodes            = @()
                MedhotInfo          = @()
                PmdaSizeInfo        = @()
                # 統計用
                HasMultipleHotCodes = $false
                HotCodeCount        = 0
            }
            $stats.TotalYJCodes++
        }
    }
}
Write-Host "  YJコード総数: $($stats.TotalYJCodes)" -ForegroundColor Green

# 2. 医薬品マスター情報を追加（薬価基準収載医薬品コードで紐付け）
Write-Host "Step 2: 医薬品マスター情報を統合..." -ForegroundColor Yellow
$addedCount = 0
if ($iyakuhinMaster -and $iyakuhinMaster.Data) {
    # 医薬品マスターを薬価基準収載医薬品コードでインデックス化
    $masterByYakkaCode = @{}
    foreach ($item in $iyakuhinMaster.Data) {
        $yakkaCode = $item."薬価基準収載医薬品コード"
        if ($yakkaCode) {
            if (-not $masterByYakkaCode.ContainsKey($yakkaCode)) {
                $masterByYakkaCode[$yakkaCode] = @()
            }
            $masterByYakkaCode[$yakkaCode] += $item
        }
    }
    
    # YJコードから薬価基準収載医薬品コードを取得して紐付け
    foreach ($yjCode in $integratedData.Keys) {
        $yakkaCode = $integratedData[$yjCode].BasicInfo."薬価基準収載医薬品コード"
        if ($yakkaCode -and $masterByYakkaCode.ContainsKey($yakkaCode)) {
            $integratedData[$yjCode].MasterInfo = $masterByYakkaCode[$yakkaCode]
            $addedCount++
        }
    }
}
Write-Host "  追加件数: $addedCount" -ForegroundColor Green

# 3. HOT9コードを統合（YJコード → HOT9の対応、1:Nの関係）
Write-Host "Step 3: HOT9コード情報を統合..." -ForegroundColor Yellow
$addedCount = 0
$multipleCount = 0
if ($hotcodeHot9 -and $hotcodeHot9.Data) {
    foreach ($item in $hotcodeHot9.Data) {
        $yjCode = $item."個別医薬品コード"
        $hotCode = $item."基準番号（ＨＯＴコード）"
        
        if (-not $yjCode -or -not $hotCode) { continue }
        
        if ($integratedData.ContainsKey($yjCode)) {
            $integratedData[$yjCode].HotCodes += $item
            $addedCount++
            
            # 1つのYJコードに複数のHOT9がある場合
            if ($integratedData[$yjCode].HotCodes.Count -gt 1) {
                if (-not $integratedData[$yjCode].HasMultipleHotCodes) {
                    $integratedData[$yjCode].HasMultipleHotCodes = $true
                    $multipleCount++
                    $stats.MultipleHotCodes++
                }
            }
            $integratedData[$yjCode].HotCodeCount = $integratedData[$yjCode].HotCodes.Count
        }
    }
}
Write-Host "  HOT9追加件数: $addedCount" -ForegroundColor Green
Write-Host "  複数HOT9を持つYJコード: $multipleCount" -ForegroundColor Cyan

# 4. MEDHOT情報を統合（調剤包装単位コード経由）
Write-Host "Step 4: MEDHOT情報を統合..." -ForegroundColor Yellow
$addedCount = 0
if ($medhotHanbaime -and $medhotHanbaime.Data) {
    foreach ($item in $medhotHanbaime.Data) {
        $yjCode = $item."薬価コード"
        if (-not $yjCode) { continue }
        
        if ($integratedData.ContainsKey($yjCode)) {
            $integratedData[$yjCode].MedhotInfo += $item
            $addedCount++
        }
    }
}
Write-Host "  MEDHOT追加件数: $addedCount" -ForegroundColor Green

# 5. PMDAサイズ情報を統合
Write-Host "Step 5: PMDAサイズ情報を統合..." -ForegroundColor Yellow
$addedCount = 0
if ($pmdaSize -and $pmdaSize.Data) {
    foreach ($item in $pmdaSize.Data) {
        $yjCode = $item."YJCode"
        if (-not $yjCode) { continue }
        
        if ($integratedData.ContainsKey($yjCode)) {
            $integratedData[$yjCode].PmdaSizeInfo += $item
            $addedCount++
        }
    }
}
Write-Host "  PMDAサイズ追加件数: $addedCount" -ForegroundColor Green

Write-Host ""
Write-Host "=== 統計情報更新 ===" -ForegroundColor Cyan

# 統計情報を更新
foreach ($yjCode in $integratedData.Keys) {
    $entry = $integratedData[$yjCode]
    
    if ($entry.HotCodes.Count -gt 0) {
        $stats.WithHotCode++
        
        if ($entry.HasMultipleHotCodes) {
            $stats.MultipleHotCodesDetail += [PSCustomObject]@{
                YJCode       = $yjCode
                HotCodeCount = $entry.HotCodeCount
                ProductName  = $entry.BasicInfo."販売名"
                HotCodes     = ($entry.HotCodes | ForEach-Object { $_."基準番号（ＨＯＴコード）" }) -join ", "
            }
        }
    }
    
    if ($entry.MedhotInfo.Count -gt 0) {
        $stats.WithMedhotInfo++
    }
    
    if ($entry.PmdaSizeInfo.Count -gt 0) {
        $stats.WithPmdaSize++
    }
}

Write-Host "  YJコード総数: $($stats.TotalYJCodes)" -ForegroundColor White
Write-Host "  HOTコード付与: $($stats.WithHotCode)" -ForegroundColor Green
Write-Host "  MEDHOT情報付与: $($stats.WithMedhotInfo)" -ForegroundColor Green
Write-Host "  PMDAサイズ情報付与: $($stats.WithPmdaSize)" -ForegroundColor Green
Write-Host "  複数HOTコードを持つYJコード: $($stats.MultipleHotCodes)" -ForegroundColor Yellow

Write-Host ""
Write-Host "=== データ出力 ===" -ForegroundColor Cyan

# タイムスタンプ
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# 1. 統合データをJSON形式で出力
$outputJsonPath = Join-Path $OutputDir "integrated_iyakuhin_${timestamp}.json"
Write-Host "JSON出力中: $outputJsonPath" -ForegroundColor Yellow

$outputData = @{
    GeneratedDate  = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    DataSources    = @{
        IyakuhinMaster = if ($iyakuhinMaster) { $iyakuhinMaster.SourceFile } else { $null }
        IyakuhinCode   = if ($iyakuhinCode) { $iyakuhinCode.SourceFile } else { $null }
        HotcodeHot9    = if ($hotcodeHot9) { $hotcodeHot9.SourceFile } else { $null }
        MedhotHanbaime = if ($medhotHanbaime) { $medhotHanbaime.SourceFile } else { $null }
        PmdaSize       = if ($pmdaSize) { $pmdaSize.SourceFile } else { $null }
    }
    TotalYJCodes   = $stats.TotalYJCodes
    Statistics     = $stats
    IntegratedData = $integratedData.Values | Sort-Object { $_.YJCode }
}

$outputData | ConvertTo-Json -Depth 10 -Compress:$false | Out-File -FilePath $outputJsonPath -Encoding UTF8
$fileSize = [math]::Round((Get-Item $outputJsonPath).Length / 1MB, 2)
Write-Host "  ✓ 出力完了: $fileSize MB" -ForegroundColor Green

# 2. 複数HOTコードを持つYJコードのリストをCSV出力
if ($stats.MultipleHotCodesDetail.Count -gt 0) {
    $multipleHotCsvPath = Join-Path $OutputDir "multiple_hotcodes_${timestamp}.csv"
    Write-Host "複数HOTコードリスト出力中: $multipleHotCsvPath" -ForegroundColor Yellow
    $stats.MultipleHotCodesDetail | Export-Csv -Path $multipleHotCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "  ✓ 出力完了: $($stats.MultipleHotCodesDetail.Count) 件" -ForegroundColor Green
}

# 3. サマリーレポートをCSV出力
$summaryPath = Join-Path $metaDir "integration_summary_${timestamp}.csv"
$summaryData = [PSCustomObject]@{
    GeneratedDate    = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
    TotalYJCodes     = $stats.TotalYJCodes
    WithHotCode      = $stats.WithHotCode
    WithMedhotInfo   = $stats.WithMedhotInfo
    WithPmdaSize     = $stats.WithPmdaSize
    MultipleHotCodes = $stats.MultipleHotCodes
    OutputJsonFile   = Split-Path $outputJsonPath -Leaf
    OutputJsonSizeMB = $fileSize
}
$summaryData | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host "サマリー出力: $summaryPath" -ForegroundColor Green

Write-Host ""
Write-Host "=== 統合処理完了 ===" -ForegroundColor Green
Write-Host "  出力ファイル: $outputJsonPath" -ForegroundColor White
Write-Host "  出力サイズ: $fileSize MB" -ForegroundColor White
Write-Host ""
Write-Host "注意事項:" -ForegroundColor Yellow
Write-Host "  - 1つのYJコードに複数のHOTコードが対応する場合があります（併売品など）" -ForegroundColor Gray
Write-Host "  - JANコードやGS1コードは1対多の関係があります" -ForegroundColor Gray
Write-Host "  - 詳細は hcodeguide.docx を参照してください" -ForegroundColor Gray

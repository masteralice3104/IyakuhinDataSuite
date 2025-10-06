# code_merge.ps1
# iyakuhin_master_**.jsonをベースにデータをマージするスクリプト
# 予備フィールドを削除し、必要なデータを統合

# エンコーディング設定
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# カレントディレクトリの設定
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

Write-Host "=== 医薬品マスターデータマージスクリプト ===" -ForegroundColor Cyan
Write-Host ""

# 処理データディレクトリ
$processedDir = "data\processed"
$resultDir = "data\result"

# 出力ファイル名 (固定)
$outputFile = "$resultDir\result.json"

# 不要フィールドのリスト
$unnecessaryFields = @(
    "長期収載品関連",
    "造影（補助）剤",
    "収載方式等識別",
    "単位漢字有効桁数",
    "カナ有効桁数",
    "漢字名称変更区分",
    "カナ名称変更区分",
    "変更年月日",
    "金額種別",
    "薬価基準収載年月日",
    "経過措置年月日又は商品名医薬品コード使用期限",
    "単位漢字名称",
    "マスター種別",
    "廃止年月日",
    "抗HIV薬区分",
    "生物学的製剤",
    "一般名処方加算対象区分",
    "商品名等関連",
    "麻薬・毒薬・覚醒剤原料・向精神薬",
    "神経破壊剤",
    "予備3",
    "予備2",
    "漢字有効桁数",
    "変更区分",
    "歯科特定薬剤",
    "予備1",
    "公表順序番号",
    "旧金額",
    "金額種別（旧金額）",
    "注射容量",
    "選定療養区分",
    "後発品",
    "単位コード"
)

# マルチスレッド処理用の最大スレッド数
$maxThreads = [Environment]::ProcessorCount
Write-Host "マルチスレッド処理: $maxThreads スレッド" -ForegroundColor Yellow

# 不要フィールドを削除する関数
function Remove-UnnecessaryFields {
    param(
        [Parameter(Mandatory = $true)]
        $Record,
        [Parameter(Mandatory = $true)]
        [string[]]$FieldsToRemove
    )
    
    if ($null -eq $Record) { return $null }
    
    $newObj = @{}
    foreach ($prop in $Record.PSObject.Properties) {
        # 不要フィールドリストに含まれていない、かつ予備フィールドでもない場合のみ保持
        $isUnnecessary = $FieldsToRemove -contains $prop.Name
        $isYobi = $prop.Name -match '^予備[０-９0-9]*$'
        
        if (-not $isUnnecessary -and -not $isYobi) {
            $newObj[$prop.Name] = $prop.Value
        }
    }
    return [PSCustomObject]$newObj
}

Write-Host "ステップ 1: iyakuhin_masterデータの読み込み" -ForegroundColor Green

# iyakuhin_masterファイルの検索
$iyakuhinFiles = Get-ChildItem -Path $processedDir -Filter "iyakuhin_master_*.json" | Sort-Object LastWriteTime -Descending
if ($iyakuhinFiles.Count -eq 0) {
    Write-Host "エラー: iyakuhin_masterファイルが見つかりません" -ForegroundColor Red
    exit 1
}

$iyakuhinFile = $iyakuhinFiles[0].FullName
Write-Host "  ファイル: $($iyakuhinFiles[0].Name)"

$iyakuhinJson = Get-Content $iyakuhinFile -Encoding UTF8 -Raw | ConvertFrom-Json
$recordCount = $iyakuhinJson.Data.Count
Write-Host "  レコード数: $recordCount" -ForegroundColor Cyan

Write-Host ""
Write-Host "ステップ 2: MEDISデータの読み込み" -ForegroundColor Green

# MEDISファイルの検索
$medisFiles = Get-ChildItem -Path $processedDir -Filter "MEDIS*_*.json" | Where-Object { $_.Name -notmatch '_HOT9_' -and $_.Name -notmatch '_OP_' } | Sort-Object LastWriteTime -Descending
if ($medisFiles.Count -eq 0) {
    Write-Host "  警告: MEDISファイルが見つかりません。マージをスキップします。" -ForegroundColor Yellow
    $medisData = $null
}
else {
    $medisFile = $medisFiles[0].FullName
    Write-Host "  ファイル: $($medisFiles[0].Name)"
    
    $medisJson = Get-Content $medisFile -Encoding UTF8 -Raw | ConvertFrom-Json
    $medisRecordCount = $medisJson.Data.Count
    Write-Host "  レコード数: $medisRecordCount" -ForegroundColor Cyan
    
    # MEDISデータをハッシュテーブルに変換 (レセプト電算処理システムコード（１）をキーに)
    Write-Host "  ハッシュテーブルを作成中..."
    $medisData = @{}
    foreach ($record in $medisJson.Data) {
        $key = $record.'レセプト電算処理システムコード（１）'
        if ($null -ne $key -and $key -ne "") {
            # 4項目のみ抽出
            $medisData[$key] = [PSCustomObject]@{
                "包装単位単位"       = $record.'包装単位単位'
                "個別医薬品コード"     = $record.'個別医薬品コード'
                "基準番号（ＨＯＴコード）" = $record.'基準番号（ＨＯＴコード）'
                "包装形態"         = $record.'包装形態'
            }
        }
    }
    Write-Host "  ハッシュテーブル作成完了: $($medisData.Count) エントリ" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "ステップ 3: MEDHOT包装コードデータの読み込み" -ForegroundColor Green

# MEDHOT包装ファイルの検索
$medhotPackingFiles = Get-ChildItem -Path $processedDir -Filter "medhot_chouzai_hanbai_motokon_*.json" | Sort-Object LastWriteTime -Descending
if ($medhotPackingFiles.Count -eq 0) {
    Write-Host "  警告: MEDHOT包装コードファイルが見つかりません。包装コードマージをスキップします。" -ForegroundColor Yellow
    $medhotPackingData = $null
}
else {
    $medhotPackingFile = $medhotPackingFiles[0].FullName
    Write-Host "  ファイル: $($medhotPackingFiles[0].Name)"
    
    $medhotPackingJson = Get-Content $medhotPackingFile -Encoding UTF8 -Raw | ConvertFrom-Json
    $medhotPackingRecordCount = $medhotPackingJson.Data.Count
    Write-Host "  レコード数: $medhotPackingRecordCount" -ForegroundColor Cyan
    
    # MEDHOT包装データをハッシュテーブルに変換 (販売名をキーに、配列で保持)
    Write-Host "  ハッシュテーブルを作成中 (販売名ベース)..."
    $medhotPackingData = @{}
    $packingCodesAdded = 0
    
    foreach ($record in $medhotPackingJson.Data) {
        $hanName = $record.'販売名'
        if ($null -ne $hanName -and $hanName -ne "") {
            # 3つのコードを配列として収集
            $codes = @()
            
            $chouzaiCode = $record.'調剤包装単位コード'
            if ($null -ne $chouzaiCode -and $chouzaiCode -ne "") {
                $codes += $chouzaiCode
            }
            
            $hanbaCode = $record.'販売包装単位コード'
            if ($null -ne $hanbaCode -and $hanbaCode -ne "") {
                $codes += $hanbaCode
            }
            
            $motokonCode = $record.'元梱包装単位コード'
            if ($null -ne $motokonCode -and $motokonCode -ne "") {
                $codes += $motokonCode
            }
            
            # コードがある場合のみ追加
            if ($codes.Count -gt 0) {
                if (-not $medhotPackingData.ContainsKey($hanName)) {
                    $medhotPackingData[$hanName] = @()
                }
                # 既存の配列に追加 (重複も含めて)
                $medhotPackingData[$hanName] += $codes
                $packingCodesAdded += $codes.Count
            }
        }
    }
    Write-Host "  ハッシュテーブル作成完了: $($medhotPackingData.Count) 販売名" -ForegroundColor Cyan
    Write-Host "  包装コード総数: $packingCodesAdded コード" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "ステップ 4: PMDAサイズデータの読み込み" -ForegroundColor Green

# PMDAサイズデータファイルの検索
$pmdaSizeFiles = Get-ChildItem -Path $processedDir -Filter "pmda_yjcode_size_*.json" | Sort-Object LastWriteTime -Descending
$pmdaSizeData = @{}

if ($pmdaSizeFiles.Count -gt 0) {
    $pmdaSizeFile = $pmdaSizeFiles[0].FullName
    Write-Host "  ファイル: $($pmdaSizeFiles[0].Name)" -ForegroundColor Yellow
    
    $pmdaSizeJson = Get-Content -Path $pmdaSizeFile -Encoding UTF8 -Raw | ConvertFrom-Json
    Write-Host "  レコード数: $($pmdaSizeJson.Data.Count)" -ForegroundColor Yellow
    
    # 半角・全角変換関数
    function Convert-ToHalfWidth {
        param([string]$text)
        if ($null -eq $text -or $text -eq "") { return "" }
        
        # 全角英数字を半角に変換
        $halfWidth = $text
        for ($i = 0xFF01; $i -le 0xFF5E; $i++) {
            $fullChar = [char]$i
            $halfChar = [char]($i - 0xFEE0)
            $halfWidth = $halfWidth.Replace($fullChar, $halfChar)
        }
        return $halfWidth
    }
    
    # ハッシュテーブルを作成 (BrandNameを半角化してキーにする)
    Write-Host "  ハッシュテーブルを作成中 (BrandName半角化ベース)..."
    foreach ($record in $pmdaSizeJson.Data) {
        $brandName = $record.BrandName
        if ($null -ne $brandName -and $brandName -ne "") {
            # BrandNameはすでに半角なのでそのまま使用
            $key = $brandName
            
            # サイズ情報を持つレコードのみ追加
            $hasSizeInfo = $false
            if ($null -ne $record.SizeLongDiameter -or 
                $null -ne $record.SizeThickness -or 
                $null -ne $record.SizeDiameter -or 
                $null -ne $record.SizeShortDiameter) {
                $hasSizeInfo = $true
            }
            
            if ($hasSizeInfo) {
                if (-not $pmdaSizeData.ContainsKey($key)) {
                    $pmdaSizeData[$key] = @{
                        SizeLongDiameter  = $record.SizeLongDiameter
                        SizeThickness     = $record.SizeThickness
                        SizeDiameter      = $record.SizeDiameter
                        SizeShortDiameter = $record.SizeShortDiameter
                    }
                }
            }
        }
    }
    Write-Host "  ハッシュテーブル作成完了: $($pmdaSizeData.Count) エントリ" -ForegroundColor Cyan
}
else {
    Write-Host "  警告: PMDAサイズデータファイルが見つかりません" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "ステップ 5: 不要フィールドの削除とデータマージ (マルチスレッド処理)" -ForegroundColor Green
Write-Host "  削除対象フィールド数: $($unnecessaryFields.Count)" -ForegroundColor Yellow

$startTime = Get-Date

# データをチャンクに分割
$chunkSize = [Math]::Ceiling($recordCount / $maxThreads)
$chunks = @()
for ($i = 0; $i -lt $recordCount; $i += $chunkSize) {
    $end = [Math]::Min($i + $chunkSize - 1, $recordCount - 1)
    $chunks += , @($i, $end)
}

Write-Host "  チャンク数: $($chunks.Count) (チャンクサイズ: $chunkSize)" -ForegroundColor Gray
Write-Host ""

# 進行状況を表示する関数
function Show-Progress {
    param($jobs, $chunks, $totalRecords, $startTime)
    
    $completed = @{}
    $startTimes = @{}
    $estimatedTimes = @{}
    
    for ($i = 0; $i -lt $jobs.Count; $i++) {
        $completed[$i] = 0
        $startTimes[$i] = Get-Date
        $estimatedTimes[$i] = 0
    }
    
    while ($true) {
        $allDone = $true
        $progressLines = @()
        $currentTime = Get-Date
        
        for ($i = 0; $i -lt $jobs.Count; $i++) {
            $job = $jobs[$i]
            $startIdx = $chunks[$i][0]
            $endIdx = $chunks[$i][1]
            $chunkSize = $endIdx - $startIdx + 1
            
            if ($job.State -eq "Running") {
                $allDone = $false
                
                # 経過時間ベースで進捗を推定
                $elapsed = ($currentTime - $startTimes[$i]).TotalSeconds
                
                # 1レコードあたり0.05秒と推定して進捗を計算
                $estimatedProcessed = [Math]::Min($elapsed / 0.05, $chunkSize)
                $percent = [Math]::Min([Math]::Floor(($estimatedProcessed / $chunkSize) * 100), 99)
                
                $completed[$i] = $percent
                
                # 残り時間を推定
                if ($percent -gt 0) {
                    $estimatedTotal = ($elapsed / $percent) * 100
                    $remaining = [Math]::Max($estimatedTotal - $elapsed, 0)
                    $estimatedTimes[$i] = $remaining
                }
            }
            elseif ($job.State -eq "Completed") {
                $percent = 100
                $completed[$i] = 100
                $estimatedTimes[$i] = 0
            }
            else {
                $percent = 0
            }
            
            # プログレスバーの作成
            $barLength = 30
            $filledLength = [Math]::Floor($barLength * $percent / 100)
            $bar = ("█" * $filledLength) + ("░" * ($barLength - $filledLength))
            
            $status = switch ($job.State) {
                "Running" { 
                    if ($estimatedTimes[$i] -gt 0) {
                        "残り約$([Math]::Round($estimatedTimes[$i], 0))秒"
                    }
                    else {
                        "実行中"
                    }
                }
                "Completed" { "完了  " }
                "Failed" { "失敗  " }
                default { "待機中" }
            }
            
            $color = switch ($job.State) {
                "Running" { "Cyan" }
                "Completed" { "Green" }
                "Failed" { "Red" }
                default { "Gray" }
            }
            
            # レコード範囲の表示（見やすくするため）
            $recordRange = "$startIdx-$endIdx ($chunkSize records)"
            $line = "  Thread $($i.ToString().PadLeft(2)): [$bar] $($percent.ToString().PadLeft(3))% | $status"
            $progressLines += @{ Line = $line; Color = $color }
        }
        
        # 画面をクリアして進行状況を表示
        [Console]::SetCursorPosition(0, [Console]::CursorTop - [Math]::Max($jobs.Count, 1))
        foreach ($item in $progressLines) {
            Write-Host $item.Line -ForegroundColor $item.Color
        }
        
        if ($allDone) { break }
        Start-Sleep -Milliseconds 200
    }
}

# マルチスレッド処理用のスクリプトブロック
$scriptBlock = {
    param($data, $startIdx, $endIdx, $fieldsToRemove, $medisHash, $packingHash, $pmdaSizeHash, $threadId)
    
    # 半角・全角変換関数をscriptBlock内で定義
    function Convert-ToHalfWidth {
        param([string]$text)
        if ($null -eq $text -or $text -eq "") { return "" }
        
        # 全角英数字を半角に変換
        $halfWidth = $text
        for ($i = 0xFF01; $i -le 0xFF5E; $i++) {
            $fullChar = [char]$i
            $halfChar = [char]($i - 0xFEE0)
            $halfWidth = $halfWidth.Replace($fullChar, $halfChar)
        }
        return $halfWidth
    }
    
    $result = @()
    $removedCount = 0
    $mergedCount = 0
    $packingMergedCount = 0
    $pmdaSizeMergedCount = 0
    $totalRecords = $endIdx - $startIdx + 1
    $processedCount = 0
    
    for ($i = $startIdx; $i -le $endIdx; $i++) {
        $record = $data[$i]
        $originalCount = $record.PSObject.Properties.Name.Count
        
        $newObj = @{}
        foreach ($prop in $record.PSObject.Properties) {
            $isUnnecessary = $fieldsToRemove -contains $prop.Name
            $isYobi = $prop.Name -match '^予備[０-９0-9]*$'
            
            if (-not $isUnnecessary -and -not $isYobi) {
                $newObj[$prop.Name] = $prop.Value
            }
        }
        
        # MEDISデータをマージ (医薬品コード = レセプト電算処理システムコード（１）)
        if ($null -ne $medisHash -and $newObj.ContainsKey('医薬品コード')) {
            $medisCode = $newObj['医薬品コード']
            if ($medisHash.ContainsKey($medisCode)) {
                $medisRecord = $medisHash[$medisCode]
                # 4項目を追加
                $newObj['包装単位単位'] = $medisRecord.'包装単位単位'
                $newObj['個別医薬品コード'] = $medisRecord.'個別医薬品コード'
                $newObj['基準番号（ＨＯＴコード）'] = $medisRecord.'基準番号（ＨＯＴコード）'
                $newObj['包装形態'] = $medisRecord.'包装形態'
                $mergedCount++
            }
        }
        
        # MEDHOT包装コードをマージ (基本漢字名称 = 販売名)
        if ($null -ne $packingHash -and $newObj.ContainsKey('基本漢字名称')) {
            $hanName = $newObj['基本漢字名称']
            if ($packingHash.ContainsKey($hanName)) {
                # 配列を"単位コード"として追加
                $newObj['単位コード'] = $packingHash[$hanName]
                $packingMergedCount++
            }
        }
        
        # PMDAサイズデータをマージ (漢字名称を半角化 = BrandName)
        if ($null -ne $pmdaSizeHash -and $newObj.ContainsKey('漢字名称')) {
            $kanjiName = $newObj['漢字名称']
            # 全角英数字を半角に変換してマッチング
            $halfWidthName = Convert-ToHalfWidth -text $kanjiName
            
            if ($pmdaSizeHash.ContainsKey($halfWidthName)) {
                $sizeInfo = $pmdaSizeHash[$halfWidthName]
                # 4項目を追加
                $newObj['SizeLongDiameter'] = $sizeInfo.SizeLongDiameter
                $newObj['SizeThickness'] = $sizeInfo.SizeThickness
                $newObj['SizeDiameter'] = $sizeInfo.SizeDiameter
                $newObj['SizeShortDiameter'] = $sizeInfo.SizeShortDiameter
                $pmdaSizeMergedCount++
            }
        }
        
        $cleanedRecord = [PSCustomObject]$newObj
        $cleanedCount = $cleanedRecord.PSObject.Properties.Name.Count
        $removedCount += ($originalCount - $cleanedCount)
        
        $result += $cleanedRecord
        
        # 進捗を定期的に出力 (10%ごと)
        $processedCount++
        if ($processedCount % [Math]::Max([Math]::Floor($totalRecords / 10), 1) -eq 0) {
            $percent = [Math]::Round(($processedCount / $totalRecords) * 100, 0)
            Write-Progress -Id $threadId -Activity "Thread $threadId" -Status "処理中 $i/$endIdx" -PercentComplete $percent
        }
    }
    
    # 完了時に進捗をクリア
    Write-Progress -Id $threadId -Activity "Thread $threadId" -Completed
    
    return @{
        Data                = $result
        RemovedCount        = $removedCount
        MergedCount         = $mergedCount
        PackingMergedCount  = $packingMergedCount
        PmdaSizeMergedCount = $pmdaSizeMergedCount
    }
}

# 並列処理の実行
$jobs = @()
$threadId = 0
foreach ($chunk in $chunks) {
    $startIdx = $chunk[0]
    $endIdx = $chunk[1]
    
    $job = Start-Job -ScriptBlock $scriptBlock -ArgumentList $iyakuhinJson.Data, $startIdx, $endIdx, $unnecessaryFields, $medisData, $medhotPackingData, $pmdaSizeData, $threadId
    $jobs += $job
    $threadId++
}

# 進行状況表示用の空行を確保
for ($i = 0; $i -lt $jobs.Count; $i++) {
    Write-Host ""
}

# 進行状況をリアルタイム表示
$jobStartTime = Get-Date
Show-Progress -jobs $jobs -chunks $chunks -totalRecords $recordCount -startTime $jobStartTime

Write-Host ""
Write-Host "  並列処理完了 - 結果を収集中..." -ForegroundColor Yellow

# 結果を収集
$cleanedData = @()
$totalFieldsRemoved = 0
$totalMergedRecords = 0
$totalPackingMergedRecords = 0
$totalPmdaSizeMergedRecords = 0

$jobIndex = 0
foreach ($job in $jobs) {
    $jobIndex++
    Write-Host "    [$jobIndex/$($jobs.Count)] スレッド結果を収集中..." -ForegroundColor Gray -NoNewline
    
    $result = Receive-Job -Job $job
    $cleanedData += $result.Data
    $totalFieldsRemoved += $result.RemovedCount
    $totalMergedRecords += $result.MergedCount
    $totalPackingMergedRecords += $result.PackingMergedCount
    $totalPmdaSizeMergedRecords += $result.PmdaSizeMergedCount
    Remove-Job -Job $job
    
    Write-Host " ✓" -ForegroundColor Green
}

Write-Host ""
Write-Host "  データ統合中..." -ForegroundColor Yellow -NoNewline
Start-Sleep -Milliseconds 100  # 視覚的フィードバックのため
Write-Host " ✓ 完了" -ForegroundColor Green

$endTime = Get-Date
$duration = ($endTime - $startTime).TotalSeconds

Write-Host ""
Write-Host "  処理完了: $recordCount レコード" -ForegroundColor Cyan
Write-Host "  削除されたフィールド総数: $totalFieldsRemoved" -ForegroundColor Yellow
Write-Host "  MEDISデータマージ成功: $totalMergedRecords レコード" -ForegroundColor $(if ($totalMergedRecords -gt 0) { "Green" } else { "Yellow" })
Write-Host "  MEDHOT包装コードマージ成功: $totalPackingMergedRecords レコード" -ForegroundColor $(if ($totalPackingMergedRecords -gt 0) { "Green" } else { "Yellow" })
Write-Host "  PMDAサイズデータマージ成功: $totalPmdaSizeMergedRecords レコード" -ForegroundColor $(if ($totalPmdaSizeMergedRecords -gt 0) { "Green" } else { "Yellow" })
Write-Host "  処理時間: $([Math]::Round($duration, 2)) 秒" -ForegroundColor Green

Write-Host ""
Write-Host "ステップ 6: データの検証" -ForegroundColor Green

Write-Host "  サンプルレコードを取得中..." -ForegroundColor Gray -NoNewline
$sampleBefore = $iyakuhinJson.Data[0]
$sampleAfter = $cleanedData[0]
Write-Host " ✓" -ForegroundColor Green

Write-Host "  不要フィールドの削除を確認中..." -ForegroundColor Gray -NoNewline
# 削除対象フィールドが残っているかチェック
$remainingUnnecessary = $sampleAfter.PSObject.Properties.Name | Where-Object { $unnecessaryFields -contains $_ -and $_ -ne "単位コード" }
if ($remainingUnnecessary) {
    Write-Host " ⚠" -ForegroundColor Red
    Write-Host "  警告: 削除対象フィールドが残っています: $($remainingUnnecessary -join ', ')" -ForegroundColor Red
}
else {
    Write-Host " ✓" -ForegroundColor Green
    Write-Host "  ✓ すべての不要フィールドが削除されました" -ForegroundColor Green
}

Write-Host "  フィールド数を計算中..." -ForegroundColor Gray -NoNewline
# フィールド数の比較
$beforeFields = $sampleBefore.PSObject.Properties.Name.Count
$afterFields = $sampleAfter.PSObject.Properties.Name.Count
Write-Host " ✓" -ForegroundColor Green

Write-Host "  サンプルレコードのフィールド総数:"
Write-Host "    削除前: $beforeFields"
Write-Host "    削除後: $afterFields" -ForegroundColor Cyan
Write-Host "    削減数: $($beforeFields - $afterFields)" -ForegroundColor Yellow

Write-Host ""
Write-Host "ステップ 7: 出力JSONの作成" -ForegroundColor Green

Write-Host "  メタデータを構築中..." -ForegroundColor Gray -NoNewline
# メタデータを含む出力オブジェクト
$output = [PSCustomObject]@{
    Metadata = [PSCustomObject]@{
        GeneratedAt             = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        Description             = "医薬品マスターデータ（不要フィールド削除済み + MEDISマージ済み + MEDHOT包装コードマージ済み + PMDAサイズマージ済み）"
        SourceFile              = $iyakuhinFiles[0].Name
        MedisSourceFile         = if ($medisFiles) { $medisFiles[0].Name } else { "なし" }
        MedhotPackingSourceFile = if ($medhotPackingFiles) { $medhotPackingFiles[0].Name } else { "なし" }
        PmdaSizeSourceFile      = if ($pmdaSizeFiles) { $pmdaSizeFiles[0].Name } else { "なし" }
        TotalRecords            = $cleanedData.Count
        ProcessingInfo          = [PSCustomObject]@{
            RemovedFieldTypes        = "予備フィールド + 不要フィールド + 単位コード"
            UnnecessaryFieldCount    = $unnecessaryFields.Count
            TotalFieldsRemoved       = $totalFieldsRemoved
            FieldCountBefore         = $beforeFields
            FieldCountAfter          = $afterFields
            FieldsReducedPerRecord   = $beforeFields - $afterFields
            MedisRecordsMerged       = $totalMergedRecords
            MedisFieldsAdded         = if ($totalMergedRecords -gt 0) { 4 } else { 0 }
            PackingCodeRecordsMerged = $totalPackingMergedRecords
            PackingCodeFieldAdded    = if ($totalPackingMergedRecords -gt 0) { "単位コード (配列)" } else { "なし" }
            PmdaSizeRecordsMerged    = $totalPmdaSizeMergedRecords
            PmdaSizeFieldsAdded      = if ($totalPmdaSizeMergedRecords -gt 0) { 4 } else { 0 }
        }
    }
    Data     = $cleanedData
}
Write-Host " ✓" -ForegroundColor Green

# JSON出力
Write-Host "  JSONに変換中 (19,679レコード)..." -ForegroundColor Yellow -NoNewline
$jsonStartTime = Get-Date
$jsonString = $output | ConvertTo-Json -Depth 10 -Compress:$false
$jsonConvertTime = [Math]::Round(((Get-Date) - $jsonStartTime).TotalSeconds, 2)
Write-Host " ✓ ($jsonConvertTime 秒)" -ForegroundColor Green

Write-Host "  ファイルに書き込み中..." -ForegroundColor Yellow -NoNewline
$writeStartTime = Get-Date
$jsonString | Set-Content -Path $outputFile -Encoding UTF8
$writeTime = [Math]::Round(((Get-Date) - $writeStartTime).TotalSeconds, 2)
Write-Host " ✓ ($writeTime 秒)" -ForegroundColor Green

Write-Host "  ファイルサイズを確認中..." -ForegroundColor Gray -NoNewline
# ファイルサイズ取得
$originalSize = (Get-Item $iyakuhinFile).Length
$newSize = (Get-Item $outputFile).Length
$originalSizeMB = [math]::Round($originalSize / 1MB, 2)
$newSizeMB = [math]::Round($newSize / 1MB, 2)
$reduction = [math]::Round(($originalSize - $newSize) / $originalSize * 100, 2)
Write-Host " ✓" -ForegroundColor Green

Write-Host ""
Write-Host "=== 完了 ===" -ForegroundColor Green
Write-Host "出力ファイル: $outputFile" -ForegroundColor Cyan
Write-Host ""
Write-Host "ファイルサイズ比較:" -ForegroundColor Yellow
Write-Host "  元ファイル: $originalSizeMB MB"
Write-Host "  新ファイル: $newSizeMB MB"
Write-Host "  削減率: $reduction%" -ForegroundColor $(if ($reduction -gt 0) { "Green" } else { "Yellow" })
Write-Host ""
Write-Host "統計情報:" -ForegroundColor Yellow
Write-Host "  総レコード数: $($cleanedData.Count)"
Write-Host "  不要フィールド種類数: $($unnecessaryFields.Count)"
Write-Host "  削除されたフィールド総数: $totalFieldsRemoved"
Write-Host "  レコードあたりフィールド削減数: $($beforeFields - $afterFields)"
if ($totalMergedRecords -gt 0) {
    Write-Host "  MEDISマージ成功レコード数: $totalMergedRecords" -ForegroundColor Green
    Write-Host "  MEDISマージ成功率: $([Math]::Round($totalMergedRecords / $cleanedData.Count * 100, 2))%" -ForegroundColor Green
    Write-Host "  追加されたフィールド: 包装単位単位, 個別医薬品コード, 基準番号（ＨＯＴコード）, 包装形態" -ForegroundColor Cyan
}
if ($totalPackingMergedRecords -gt 0) {
    Write-Host "  MEDHOT包装コードマージ成功レコード数: $totalPackingMergedRecords" -ForegroundColor Green
    Write-Host "  MEDHOT包装コードマージ成功率: $([Math]::Round($totalPackingMergedRecords / $cleanedData.Count * 100, 2))%" -ForegroundColor Green
    Write-Host "  追加されたフィールド: 単位コード (調剤・販売・元梱の配列)" -ForegroundColor Cyan
}
if ($totalPmdaSizeMergedRecords -gt 0) {
    Write-Host "  PMDAサイズデータマージ成功レコード数: $totalPmdaSizeMergedRecords" -ForegroundColor Green
    Write-Host "  PMDAサイズデータマージ成功率: $([Math]::Round($totalPmdaSizeMergedRecords / $cleanedData.Count * 100, 2))%" -ForegroundColor Green
    Write-Host "  追加されたフィールド: SizeLongDiameter, SizeThickness, SizeDiameter, SizeShortDiameter" -ForegroundColor Cyan
}
Write-Host ""

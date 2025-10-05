param(
    [string]$InputDir = "data\raw",
    [string]$OutputDir = "data\processed",
    [string]$SpecificFile = "",
    [switch]$AnalyzeStructure,
    [switch]$ProcessCsv,
    [switch]$ProcessExcel,
    [switch]$Help
)

if ($Help) {
    Write-Host @"

薬価基準データ処理スクリプト
===========================

使用方法:
  .\data_process.ps1                              # 全ファイルをJSON変換
  .\data_process.ps1 -ProcessExcel                # Excelファイルのみ処理
  .\data_process.ps1 -ProcessCsv                  # CSVファイルのみ処理
  .\data_process.ps1 -AnalyzeStructure            # ファイル構造を解析
  .\data_process.ps1 -SpecificFile "filename"     # 特定ファイルのみ処理
  .\data_process.ps1 -InputDir "path"             # 入力ディレクトリ指定
  .\data_process.ps1 -OutputDir "path"            # 出力ディレクトリ指定

パラメータ:
  -InputDir <path>      入力ディレクトリ (デフォルト: data\raw)
  -OutputDir <path>     出力ディレクトリ (デフォルト: data\processed)
  -SpecificFile <name>  処理対象ファイル名
  -ProcessExcel         Excelファイルのみ処理
  -ProcessCsv           CSVファイルのみ処理
  -AnalyzeStructure     ファイル構造解析モード
  -Help                このヘルプを表示

変換されるファイル:
  - tp****-**_**.xlsx (薬価基準各種ファイル)
  - tp****_kiso.xlsx (基礎的医薬品リスト)
  - iyakuhin_master_*.csv (医薬品マスター Shift_JIS CSV)

"@ -ForegroundColor White
    exit 0
}

# ImportExcel モジュールの確認とインストール
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel モジュールをインストール中..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Force -AllowClobber -Scope CurrentUser
        Write-Host "ImportExcel モジュールのインストールが完了しました" -ForegroundColor Green
    }
    catch {
        Write-Error "ImportExcel モジュールのインストールに失敗しました: $($_.Exception.Message)"
        Write-Host "手動でインストールしてください: Install-Module -Name ImportExcel" -ForegroundColor Yellow
        exit 1
    }
}

Import-Module ImportExcel -Force

# 出力ディレクトリの作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# JSON出力用のメタディレクトリ作成
$jsonMetaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $jsonMetaDir)) {
    New-Item -ItemType Directory -Path $jsonMetaDir -Force | Out-Null
    Write-Host "JSONメタディレクトリを作成しました: $jsonMetaDir" -ForegroundColor Green
}

# ファイル種別の定義
$fileTypes = @{
    "01"                            = @{ Name = "内用薬"; Category = "naiyou" }
    "02"                            = @{ Name = "注射薬"; Category = "chuusha" }
    "03"                            = @{ Name = "外用薬"; Category = "gaiyou" }
    "04"                            = @{ Name = "歯科用薬剤"; Category = "shika" }
    "05"                            = @{ Name = "その他（後発医薬品情報）"; Category = "kouhatsu" }
    "06"                            = @{ Name = "その他（加算対象外品目）"; Category = "kasan_gai" }
    "07"                            = @{ Name = "その他（カットオフ値算出対象品目）"; Category = "cutoff" }
    "kiso"                          = @{ Name = "基礎的医薬品リスト"; Category = "kiso" }
    "iyakuhin_master"               = @{ Name = "医薬品マスター"; Category = "iyakuhin_master" }
    "iyakuhin_code"                 = @{ Name = "医薬品コード"; Category = "iyakuhin_code" }
    "hotcode"                       = @{ Name = "HOTコードマスター（標準）"; Category = "hotcode" }
    "hotcode_hot9"                  = @{ Name = "HOTコードマスター（HOT9）"; Category = "hotcode_hot9" }
    "hotcode_op"                    = @{ Name = "HOTコードマスター（オプション）"; Category = "hotcode_op" }
    "pmda_yjcode_size"              = @{ Name = "PMDA YJコード・サイズ情報"; Category = "pmda_yjcode_size" }
    "medhot_hanbaime_chouzai"       = @{ Name = "MEDHOT 販売名・調剤包装単位コード"; Category = "medhot_hanbaime_chouzai" }
    "medhot_chouzai_hanbai_motokon" = @{ Name = "MEDHOT 調剤・販売・元梱包装単位コード"; Category = "medhot_chouzai_hanbai_motokon" }
}

# Excelファイルの構造解析関数
function Analyze-ExcelStructure {
    param([string]$FilePath)
    
    Write-Host "`n=== ファイル構造解析: $(Split-Path $FilePath -Leaf) ===" -ForegroundColor Cyan
    
    try {
        # Excelパッケージを開いてワークシート一覧を取得
        $workbook = Open-ExcelPackage -Path $FilePath
        $worksheets = $workbook.Workbook.Worksheets
        Write-Host "ワークシート数: $($worksheets.Count)" -ForegroundColor Green
        
        foreach ($sheet in $worksheets) {
            Write-Host "`nワークシート: $($sheet.Name)" -ForegroundColor Yellow
            Write-Host "  行数: $($sheet.Dimension.Rows)" -ForegroundColor White
            Write-Host "  列数: $($sheet.Dimension.Columns)" -ForegroundColor White
            
            # 最初の数行を表示してヘッダー構造を確認
            if ($sheet.Dimension.Rows -gt 0) {
                Write-Host "  先頭5行のサンプル:" -ForegroundColor White
                try {
                    $sampleData = Import-Excel -Path $FilePath -WorksheetName $sheet.Name -StartRow 1 -EndRow 5 -NoHeader
                    
                    for ($i = 0; $i -lt [Math]::Min(5, $sampleData.Count); $i++) {
                        $row = $sampleData[$i]
                        $rowData = @()
                        $row.PSObject.Properties | ForEach-Object {
                            if ($_.Value) {
                                $cellValue = $_.Value.ToString()
                                $rowData += $cellValue.Substring(0, [Math]::Min(20, $cellValue.Length))
                            }
                            else {
                                $rowData += "(空)"
                            }
                        }
                        Write-Host "    行$($i+1): $($rowData -join ' | ')" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "    サンプルデータの取得に失敗: $($_.Exception.Message)" -ForegroundColor Red
                }
            }
        }
        
        # パッケージを閉じる
        Close-ExcelPackage $workbook
    }
    catch {
        Write-Error "ファイル解析エラー: $($_.Exception.Message)"
        if ($workbook) { Close-ExcelPackage $workbook }
    }
}

# CSVファイルの構造解析関数
function Analyze-CsvStructure {
    param([string]$FilePath)
    
    Write-Host "`n=== ファイル構造解析: $(Split-Path $FilePath -Leaf) ===" -ForegroundColor Cyan
    
    try {
        # Shift_JISエンコーディングでCSVファイルを読み取り
        $csvContent = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetString([System.IO.File]::ReadAllBytes($FilePath))
        $lines = $csvContent.Split("`n")
        
        Write-Host "総行数: $($lines.Count)" -ForegroundColor Green
        
        # 最初の数行をパースしてヘッダー構造を確認
        Write-Host "先頭5行のサンプル:" -ForegroundColor White
        
        for ($i = 0; $i -lt [Math]::Min(5, $lines.Count); $i++) {
            if ($lines[$i].Trim()) {
                # CSVパース（簡易版）
                $fields = $lines[$i] -split '","' | ForEach-Object { $_.Trim('"') }
                Write-Host "  行$($i+1): $($fields.Count)列 - $($fields[0..4] -join ' | ')..." -ForegroundColor Gray
            }
        }
        
        # 最初の行から列数を推定
        if ($lines[0]) {
            $firstRowFields = $lines[0] -split '","' | ForEach-Object { $_.Trim('"') }
            Write-Host "推定列数: $($firstRowFields.Count)" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "CSVファイル解析エラー: $($_.Exception.Message)"
    }
}

# CSVをJSONに変換する関数
function Convert-CsvToJson {
    param(
        [string]$FilePath,
        [string]$OutputPath,
        [hashtable]$FileTypeInfo
    )
    
    Write-Host "`n=== JSON変換: $(Split-Path $FilePath -Leaf) ===" -ForegroundColor Cyan
    
    try {
        $fileName = Split-Path $FilePath -Leaf
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        
        # エンコーディングを判定
        if ($FileTypeInfo.Category -match "pmda") {
            $encoding = "UTF8"
        }
        else {
            $encoding = "Shift_JIS"
        }
        
        # CSVファイルを読み取り
        Write-Host "  ${encoding}エンコーディングでCSVを読み取り中..." -ForegroundColor Yellow
        if ($encoding -eq "UTF8") {
            $csvContent = [System.IO.File]::ReadAllText($FilePath, [System.Text.Encoding]::UTF8)
        }
        else {
            $csvContent = [System.Text.Encoding]::GetEncoding("Shift_JIS").GetString([System.IO.File]::ReadAllBytes($FilePath))
        }
        
        # 行に分割
        $lines = $csvContent.Split("`n") | Where-Object { $_.Trim() }
        
        # CSVデータをパース
        $csvData = @()
        $hasHeader = $false
        $startRow = 0
        
        # ファイル種別に応じたヘッダー設定
        if ($FileTypeInfo.Category -eq "iyakuhin_code") {
            # 医薬品コード: ヘッダー行あり
            $hasHeader = $true
            $startRow = 1
            $headerLine = $lines[0]
            $headers = $headerLine -split ',' | ForEach-Object { $_.Trim('"') }
            Write-Host "  ヘッダー行を検出: $($headers.Count)列" -ForegroundColor Green
        }
        elseif ($FileTypeInfo.Category -eq "iyakuhin_master") {
            # 医薬品マスター: ヘッダー行なし（正式なヘッダー定義を使用）
            $hasHeader = $false
            $startRow = 0
            # 医薬品マスターの正式なヘッダー定義（R06診療報酬対応）
            $headers = @(
                "変更区分",                              # 1
                "マスター種別",                          # 2
                "医薬品コード",                          # 3
                "漢字有効桁数",                          # 4
                "漢字名称",                              # 5
                "カナ有効桁数",                          # 6
                "カナ名称",                              # 7
                "単位コード",                            # 8
                "単位漢字有効桁数",                      # 9
                "単位漢字名称",                          # 10
                "金額種別",                              # 11
                "新又は現金額",                          # 12
                "予備1",                                 # 13
                "麻薬・毒薬・覚醒剤原料・向精神薬",      # 14
                "神経破壊剤",                            # 15
                "生物学的製剤",                          # 16
                "後発品",                                # 17
                "予備2",                                 # 18
                "歯科特定薬剤",                          # 19
                "造影（補助）剤",                        # 20
                "注射容量",                              # 21
                "収載方式等識別",                        # 22
                "商品名等関連",                          # 23
                "金額種別（旧金額）",                    # 24
                "旧金額",                                # 25
                "漢字名称変更区分",                      # 26
                "カナ名称変更区分",                      # 27
                "剤形",                                  # 28
                "予備3",                                 # 29
                "変更年月日",                            # 30
                "廃止年月日",                            # 31
                "薬価基準収載医薬品コード",              # 32
                "公表順序番号",                          # 33
                "経過措置年月日又は商品名医薬品コード使用期限", # 34
                "基本漢字名称",                          # 35
                "薬価基準収載年月日",                    # 36
                "一般名コード",                          # 37
                "一般名処方の標準的な記載",              # 38
                "一般名処方加算対象区分",                # 39
                "抗HIV薬区分",                           # 40
                "長期収載品関連",                        # 41
                "選定療養区分"                           # 42
            )
        }
        else {
            # その他のCSVファイル: ヘッダー行を確認
            $hasHeader = $true
            $startRow = 1
            $headerLine = $lines[0]
            $headers = $headerLine -split ',' | ForEach-Object { $_.Trim().Trim('"') }
            Write-Host "  ヘッダー行を検出: $($headers.Count)列" -ForegroundColor Green
        }
        
        # ヘッダー数の確認
        if ($FileTypeInfo.Category -eq "iyakuhin_master" -and $lines.Count -gt 0) {
            # 医薬品マスター: 最初の行でカラム数を確認
            $firstRowFields = $lines[0] -split '","' | ForEach-Object { $_.Trim('"') }
            $columnCount = $firstRowFields.Count
            
            # 実際の列数とヘッダー数が異なる場合の警告
            if ($columnCount -ne $headers.Count) {
                Write-Warning "列数が期待値と異なります: 実際=$columnCount, 期待=$($headers.Count)"
                # 実際の列数に合わせてヘッダーを調整
                if ($columnCount -gt $headers.Count) {
                    for ($i = $headers.Count; $i -lt $columnCount; $i++) {
                        $headers += "追加項目$($i - $headers.Count + 1)"
                    }
                }
                else {
                    $headers = $headers[0..($columnCount - 1)]
                }
            }
        }
        else {
            # ヘッダー行から列数を取得
            $columnCount = $headers.Count
        }
        
        # データ行数を計算
        $dataRowCount = if ($hasHeader) { $lines.Count - 1 } else { $lines.Count }
        Write-Host "  データ行数: $dataRowCount" -ForegroundColor Green
        Write-Host "  列数: $columnCount" -ForegroundColor Green
        
        # 各行をオブジェクトに変換
        foreach ($line in $lines[$startRow..($lines.Count - 1)]) {
            if ($line.Trim()) {
                # ファイルタイプに応じて区切り文字を選択
                if ($FileTypeInfo.Category -eq "iyakuhin_master") {
                    # 医薬品マスター: ダブルクォート付きCSV
                    $fields = $line -split '","' | ForEach-Object { $_.Trim('"') }
                }
                else {
                    # その他: カンマ区切り、ダブルクォートを除去
                    $fields = $line -split ',' | ForEach-Object { $_.Trim().Trim('"') }
                }
                    
                # フィールド数を統一（不足分は空文字で補完）
                while ($fields.Count -lt $columnCount) {
                    $fields += ""
                }
                    
                $rowObject = @{}
                for ($i = 0; $i -lt $columnCount; $i++) {
                    $value = $fields[$i]
                    $headerName = $headers[$i]
                    
                    # コード類は常に文字列として保持（数値変換しない）
                    $isCodeField = $headerName -match 'コード|code|CODE|番号|ＪＡＮ|JAN'
                    
                    # 数値変換を試行（コードフィールド以外で、半角数字のみ）
                    if (-not $isCodeField -and $value -and $value -match '^-?[0-9]+\.?[0-9]*$' -and $value -notmatch '[^\x00-\x7F]') {
                        try {
                            # 小数点を含む場合のみdecimal変換、それ以外は文字列
                            if ($value -match '\.') {
                                $rowObject[$headerName] = [decimal]$value
                            }
                            else {
                                $rowObject[$headerName] = $value
                            }
                        }
                        catch {
                            # 変換失敗時は文字列として保持
                            $rowObject[$headerName] = $value
                        }
                    }
                    else {
                        $rowObject[$headerName] = if ($value) { $value } else { $null }
                    }
                }
                $csvData += $rowObject
            }
        }
        
        # JSON結果を作成
        $result = @{
            SourceFile    = $fileName
            ProcessedDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            FileType      = $FileTypeInfo
            Encoding      = "Shift_JIS"
            Headers       = $headers
            RowCount      = $csvData.Count
            ColumnCount   = $headers.Count
            Data          = $csvData
        }
        
        # JSONファイルとして出力
        $jsonPath = Join-Path $OutputPath "$baseName.json"
        $result | ConvertTo-Json -Depth 10 -Compress:$false | Out-File -FilePath $jsonPath -Encoding UTF8
        
        Write-Host "  → JSON出力完了: $jsonPath" -ForegroundColor Green
        
        # メタデータ情報を作成
        $metadata = @{
            SourceFile    = $fileName
            JsonFile      = "$baseName.json"
            ProcessedDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            FileType      = $FileTypeInfo.Name
            Category      = $FileTypeInfo.Category
            RowCount      = $csvData.Count
            ColumnCount   = $headers.Count
            Encoding      = "Shift_JIS"
            TotalDataRows = $csvData.Count
        }
        
        return $metadata
    }
    catch {
        Write-Error "CSV→JSON変換エラー ($FilePath): $($_.Exception.Message)"
        return $null
    }
}

# ExcelをJSONに変換する関数
function Convert-ExcelToJson {
    param(
        [string]$FilePath,
        [string]$OutputPath,
        [hashtable]$FileTypeInfo
    )
    
    Write-Host "`n=== JSON変換: $(Split-Path $FilePath -Leaf) ===" -ForegroundColor Cyan
    
    try {
        $fileName = Split-Path $FilePath -Leaf
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
        
        # ワークシート情報を取得
        $workbook = Open-ExcelPackage -Path $FilePath
        $result = @{
            SourceFile    = $fileName
            ProcessedDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            FileType      = $FileTypeInfo
            Worksheets    = @()
        }
        
        foreach ($sheet in $workbook.Workbook.Worksheets) {
            Write-Host "  ワークシート処理中: $($sheet.Name)" -ForegroundColor Yellow
            
            # データを読み込み
            $data = Import-Excel -Path $FilePath -WorksheetName $sheet.Name
            
            $worksheetInfo = @{
                Name        = $sheet.Name
                RowCount    = $sheet.Dimension.Rows
                ColumnCount = $sheet.Dimension.Columns
                Data        = @()
            }
            
            # データが存在する場合のみ処理
            if ($data -and $data.Count -gt 0) {
                # 最初の行からヘッダーを特定
                $headers = @()
                $firstRow = $data[0]
                $firstRow.PSObject.Properties | ForEach-Object {
                    $headers += $_.Name
                }
                
                $worksheetInfo.Headers = $headers
                $worksheetInfo.Data = $data
                
                Write-Host "    データ行数: $($data.Count)" -ForegroundColor Green
                Write-Host "    ヘッダー: $($headers -join ', ')" -ForegroundColor Gray
            }
            
            $result.Worksheets += $worksheetInfo
        }
        
        # JSONファイルとして出力
        $jsonPath = Join-Path $OutputPath "$baseName.json"
        $result | ConvertTo-Json -Depth 10 -Compress:$false | Out-File -FilePath $jsonPath -Encoding UTF8
        
        Write-Host "  → JSON出力完了: $jsonPath" -ForegroundColor Green
        
        # パッケージを閉じる
        Close-ExcelPackage $workbook
        
        # メタデータ情報を作成
        $metadata = @{
            SourceFile     = $fileName
            JsonFile       = "$baseName.json"
            ProcessedDate  = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            FileType       = $FileTypeInfo.Name
            Category       = $FileTypeInfo.Category
            WorksheetCount = $result.Worksheets.Count
            TotalDataRows  = ($result.Worksheets | ForEach-Object { if ($_.Data) { $_.Data.Count } else { 0 } } | Measure-Object -Sum).Sum
        }
        
        return $metadata
    }
    catch {
        Write-Error "JSON変換エラー ($FilePath): $($_.Exception.Message)"
        if ($workbook) { Close-ExcelPackage $workbook }
        return $null
    }
}

# メイン処理
Write-Host "薬価基準データ処理を開始します..." -ForegroundColor Green

# 処理対象ファイルを取得
$excelFiles = @()
$csvFiles = @()

if ($SpecificFile) {
    # 特定ファイルが指定された場合
    $targetFile = Get-ChildItem -Path $InputDir -Filter $SpecificFile -ErrorAction SilentlyContinue
    if ($targetFile) {
        if ($targetFile.Extension -eq ".xlsx") {
            $excelFiles = @($targetFile)
        }
        elseif ($targetFile.Extension -eq ".csv" -or $targetFile.Extension -eq ".txt") {
            $csvFiles = @($targetFile)
        }
    }
}
else {
    # ファイル種別に応じて処理対象を決定
    if ($ProcessExcel -or (-not $ProcessCsv -and -not $ProcessExcel)) {
        $excelFiles = Get-ChildItem -Path $InputDir -Filter "tp*.xlsx"
    }
    if ($ProcessCsv -or (-not $ProcessCsv -and -not $ProcessExcel)) {
        $csvFiles = @()
        $csvFiles += Get-ChildItem -Path $InputDir -Filter "iyakuhin_master*.csv"
        $csvFiles += Get-ChildItem -Path $InputDir -Filter "iyakuhin_code*.txt"
        $csvFiles += Get-ChildItem -Path $InputDir -Filter "MEDIS*.txt"
        $csvFiles += Get-ChildItem -Path $InputDir -Filter "pmda_yjcode_size*.csv"
        $csvFiles += Get-ChildItem -Path $InputDir -Filter "medhot_*.txt"
    }
}

$totalFiles = $excelFiles.Count + $csvFiles.Count
if ($totalFiles -eq 0) {
    Write-Warning "処理対象のファイルが見つかりませんでした"
    exit 1
}

Write-Host "処理対象ファイル数: $totalFiles (Excel: $($excelFiles.Count), CSV: $($csvFiles.Count))" -ForegroundColor Cyan

# 構造解析モードの場合
if ($AnalyzeStructure) {
    foreach ($file in $excelFiles) {
        Analyze-ExcelStructure -FilePath $file.FullName
    }
    foreach ($file in $csvFiles) {
        Analyze-CsvStructure -FilePath $file.FullName
    }
    exit 0
}

# JSON変換処理
$processedMetadata = @()

# Excelファイルの処理
foreach ($file in $excelFiles) {
    # ファイル種別を特定
    $fileTypeInfo = $null
    $fileName = $file.Name
    
    if ($fileName -match "tp\d{8}-\d{2}_(\d{2})\.xlsx") {
        $typeCode = $matches[1]
        $fileTypeInfo = $fileTypes[$typeCode]
    }
    elseif ($fileName -match "tp\d{4}_kiso\.xlsx") {
        $fileTypeInfo = $fileTypes["kiso"]
    }
    
    if (-not $fileTypeInfo) {
        Write-Warning "Excelファイル種別を特定できませんでした: $fileName"
        $fileTypeInfo = @{ Name = "不明"; Category = "unknown" }
    }
    
    # JSON変換実行
    $metadata = Convert-ExcelToJson -FilePath $file.FullName -OutputPath $OutputDir -FileTypeInfo $fileTypeInfo
    if ($metadata) {
        $processedMetadata += $metadata
    }
}

# CSVファイルの処理
foreach ($file in $csvFiles) {
    # ファイル種別を特定
    $fileName = $file.Name
    if ($fileName -match "iyakuhin_code") {
        $fileTypeInfo = $fileTypes["iyakuhin_code"]
    }
    elseif ($fileName -match "iyakuhin_master") {
        $fileTypeInfo = $fileTypes["iyakuhin_master"]
    }
    elseif ($fileName -match "pmda_yjcode_size") {
        $fileTypeInfo = $fileTypes["pmda_yjcode_size"]
    }
    elseif ($fileName -match "MEDIS\d+_HOT9") {
        $fileTypeInfo = $fileTypes["hotcode_hot9"]
    }
    elseif ($fileName -match "MEDIS\d+_OP") {
        $fileTypeInfo = $fileTypes["hotcode_op"]
    }
    elseif ($fileName -match "MEDIS\d+_\d+\.txt") {
        $fileTypeInfo = $fileTypes["hotcode"]
    }
    elseif ($fileName -match "medhot_hanbaime_chouzai") {
        $fileTypeInfo = $fileTypes["medhot_hanbaime_chouzai"]
    }
    elseif ($fileName -match "medhot_chouzai_hanbai_motokon") {
        $fileTypeInfo = $fileTypes["medhot_chouzai_hanbai_motokon"]
    }
    else {
        $fileTypeInfo = @{ Name = "不明なCSVファイル"; Category = "unknown" }
    }
    
    # JSON変換実行
    $metadata = Convert-CsvToJson -FilePath $file.FullName -OutputPath $OutputDir -FileTypeInfo $fileTypeInfo
    if ($metadata) {
        $processedMetadata += $metadata
    }
}

# 処理結果のメタデータを保存
if ($processedMetadata.Count -gt 0) {
    $metadataPath = Join-Path $jsonMetaDir "data_to_json_metadata.csv"
    $processedMetadata | Export-Csv -Path $metadataPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nメタデータを保存しました: $metadataPath" -ForegroundColor Green
    
    $summaryPath = Join-Path $jsonMetaDir "processing_summary.json"
    $summary = @{
        ProcessedDate = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
        TotalFiles    = $processedMetadata.Count
        TotalDataRows = ($processedMetadata | ForEach-Object { $_.TotalDataRows } | Measure-Object -Sum).Sum
        FileTypes     = $processedMetadata | Group-Object FileType | ForEach-Object { @{ Type = $_.Name; Count = $_.Count } }
        ExcelFiles    = ($processedMetadata | Where-Object { $_.SourceFile -like "*.xlsx" }).Count
        CsvFiles      = ($processedMetadata | Where-Object { $_.SourceFile -like "*.csv" }).Count
        Files         = $processedMetadata
    }
    $summary | ConvertTo-Json -Depth 5 | Out-File -FilePath $summaryPath -Encoding UTF8
    Write-Host "処理サマリを保存しました: $summaryPath" -ForegroundColor Green
}

Write-Host "`n=== 処理完了 ===" -ForegroundColor Green
Write-Host "処理ファイル数: $($processedMetadata.Count)" -ForegroundColor White
Write-Host "出力ディレクトリ: $OutputDir" -ForegroundColor White
Write-Host "総データ行数: $(($processedMetadata | ForEach-Object { $_.TotalDataRows } | Measure-Object -Sum).Sum)" -ForegroundColor White

# 薬価基準収載品目リストExcelファイル取得スクリプト
# 厚生労働省の薬価基準ページから最新のExcelファイルを取得する

param(
    [string]$OutputDir = "data\raw",
    [string]$BaseUrl = $null,  # 指定されない場合は自動検出
    [int]$MaxYearRange = 3,    # 現在年度から何年前まで検索するか
    [switch]$ShowHelp          # ヘルプ表示
)

# ヘルプ表示
if ($ShowHelp) {
    Write-Host @"
薬価基準収載品目リスト取得スクリプト

使用方法:
  .\yakkakijun_get.ps1 [オプション]

オプション:
  -OutputDir <path>      出力ディレクトリ (デフォルト: data\raw)
  -BaseUrl <url>         取得元URL (省略時は自動検出)
  -MaxYearRange <num>    URL自動検出時の検索年数範囲 (デフォルト: 3年)
  -ShowHelp              このヘルプを表示

例:
  .\yakkakijun_get.ps1                                    # 自動検出で最新ページから取得
  .\yakkakijun_get.ps1 -OutputDir "custom\output"         # 出力先指定
  .\yakkakijun_get.ps1 -BaseUrl "https://..."             # URL手動指定
  .\yakkakijun_get.ps1 -MaxYearRange 5                    # 5年前まで検索

取得されるファイル:
  - 内用薬 (01)
  - 注射薬 (02)  
  - 外用薬 (03)
  - 歯科用薬剤 (04)
  - その他（後発医薬品情報） (05)
  - その他（加算対象外品目） (06)
  - その他（カットオフ値算出対象品目） (07)
  - 基礎的医薬品リスト (08)
"@ -ForegroundColor White
    exit 0
}

# 最新の薬価基準ページURLを自動検出する関数
function Find-LatestYakkakijunPage {
    param(
        [int]$MaxYearRange = 3
    )
    
    $currentYear = (Get-Date).Year
    Write-Host "最新の薬価基準ページを検索中..." -ForegroundColor Yellow
    
    # 現在年度から過去に向かって検索
    for ($yearOffset = 0; $yearOffset -le $MaxYearRange; $yearOffset++) {
        $targetYear = $currentYear - $yearOffset
        
        # 各年度で可能性のある日付パターンを試行（4月1日、10月1日など）
        $candidateDates = @(
            "${targetYear}0401",  # 4月1日
            "${targetYear}1001",  # 10月1日
            "${targetYear}0701",  # 7月1日
            "${targetYear}0101"   # 1月1日
        )
        
        foreach ($dateStr in $candidateDates) {
            $candidateUrl = "https://www.mhlw.go.jp/topics/$targetYear/04/tp$dateStr-01.html"
            
            try {
                Write-Host "  検証中: $candidateUrl" -ForegroundColor Cyan
                $response = Invoke-WebRequest -Uri $candidateUrl -Method Head -UseBasicParsing -TimeoutSec 10
                
                if ($response.StatusCode -eq 200) {
                    Write-Host "  ✓ 有効なページを発見: $candidateUrl" -ForegroundColor Green
                    return $candidateUrl
                }
            }
            catch {
                Write-Host "  ✗ 無効: $candidateUrl" -ForegroundColor DarkGray
                continue
            }
        }
    }
    
    Write-Warning "最新のページが見つかりませんでした。デフォルトURLを使用します。"
    return "https://www.mhlw.go.jp/topics/2025/04/tp20250401-01.html"
}

# BaseUrlが指定されていない場合は自動検出
if (-not $BaseUrl) {
    $BaseUrl = Find-LatestYakkakijunPage -MaxYearRange $MaxYearRange
    Write-Host "使用するURL: $BaseUrl" -ForegroundColor Green
}
else {
    Write-Host "指定されたURL: $BaseUrl" -ForegroundColor Green
}

# 出力ディレクトリの作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "出力ディレクトリを作成しました: $OutputDir" -ForegroundColor Green
}

# メタデータディレクトリの作成
$metaDir = Join-Path $OutputDir "meta"
if (-not (Test-Path $metaDir)) {
    New-Item -ItemType Directory -Path $metaDir -Force | Out-Null
    Write-Host "メタデータディレクトリを作成しました: $metaDir" -ForegroundColor Green
}

# 既存のメタデータファイルをmetaディレクトリに移動
$oldMetadataPath = Join-Path $OutputDir "download_metadata.csv"
$newMetadataPath = Join-Path $metaDir "yakkakijun_download_metadata.csv"
if ((Test-Path $oldMetadataPath) -and -not (Test-Path $newMetadataPath)) {
    Move-Item -Path $oldMetadataPath -Destination $newMetadataPath -Force
    Write-Host "既存のメタデータファイルを移動しました: $newMetadataPath" -ForegroundColor Green
}

# ウェブページの内容を取得
Write-Host "ウェブページを取得中..." -ForegroundColor Yellow
try {
    $webContent = Invoke-WebRequest -Uri $BaseUrl -UseBasicParsing
    Write-Host "ウェブページの取得が完了しました" -ForegroundColor Green
}
catch {
    Write-Error "ウェブページの取得に失敗しました: $_"
    exit 1
}

# Excelファイルのリンクを抽出
Write-Host "Excelファイルのリンクを抽出中..." -ForegroundColor Yellow

# tp[日付]-01_[番号].xlsxパターンにマッチするリンクを抽出
$excelLinks = $webContent.Links | Where-Object { 
    $_.href -match '/xls/tp(\d{8})-01_(\d{2})\.xlsx$' 
} | ForEach-Object {
    [PSCustomObject]@{
        FullUrl  = if ($_.href.StartsWith("http")) { $_.href } else { "https://www.mhlw.go.jp" + $_.href }
        FileName = [System.IO.Path]::GetFileName($_.href)
        Date     = $matches[1]
        Type     = $matches[2]
        TypeName = switch ($matches[2]) {
            "01" { "内用薬" }
            "02" { "注射薬" }
            "03" { "外用薬" }
            "04" { "歯科用薬剤" }
            "05" { "その他（後発医薬品情報）" }
            "06" { "その他（加算対象外品目）" }
            "07" { "その他（カットオフ値算出対象品目）" }
            default { "不明（種類：$($matches[2])）" }
        }
    }
}

# 基礎的リストExcel版のリンクを抽出
$kisoLinks = $webContent.Links | Where-Object { 
    $_.href -match '/xls/tp\d+_kiso\.xlsx$' 
} | ForEach-Object {
    [PSCustomObject]@{
        FullUrl  = if ($_.href.StartsWith("http")) { $_.href } else { "https://www.mhlw.go.jp" + $_.href }
        FileName = [System.IO.Path]::GetFileName($_.href)
        Date     = "99999999"  # 基礎的リストは日付ベースの更新ではないため、常に最新として扱う
        Type     = "08"
        TypeName = "基礎的医薬品リスト"
    }
}

if (-not $excelLinks -and -not $kisoLinks) {
    Write-Error "Excelファイルのリンクが見つかりませんでした"
    exit 1
}

# 全てのファイルリンクを結合
$allLinks = @()
if ($excelLinks) { $allLinks += $excelLinks }
if ($kisoLinks) { $allLinks += $kisoLinks }

Write-Host "抽出されたファイル数: $($allLinks.Count) (通常: $($excelLinks.Count), 基礎的リスト: $($kisoLinks.Count))" -ForegroundColor Green

# 各タイプごとに最新のファイルを特定
Write-Host "各種類の最新ファイルを特定中..." -ForegroundColor Yellow

$latestFiles = $allLinks | 
Group-Object Type | 
ForEach-Object {
    $_.Group | Sort-Object Date -Descending | Select-Object -First 1
}

Write-Host "最新ファイル一覧:" -ForegroundColor Cyan
$latestFiles | ForEach-Object {
    Write-Host "  [$($_.Type)] $($_.TypeName): $($_.FileName) (更新日: $($_.Date))" -ForegroundColor White
}

# ファイルをダウンロード
Write-Host "`nファイルをダウンロード中..." -ForegroundColor Yellow

$downloadedFiles = @()

foreach ($file in $latestFiles) {
    $outputPath = Join-Path $OutputDir $file.FileName
    
    try {
        Write-Host "  ダウンロード中: $($file.FileName)" -ForegroundColor Cyan
        Invoke-WebRequest -Uri $file.FullUrl -OutFile $outputPath -UseBasicParsing
        
        $fileInfo = Get-Item $outputPath
        Write-Host "    完了: $($fileInfo.Length) bytes" -ForegroundColor Green
        
        $downloadedFiles += [PSCustomObject]@{
            Type     = $file.Type
            TypeName = $file.TypeName
            FileName = $file.FileName
            FilePath = $outputPath
            Size     = $fileInfo.Length
            Date     = $file.Date
        }
    }
    catch {
        Write-Error "    エラー: $($file.FileName) のダウンロードに失敗しました: $_"
    }
}

# 結果の表示
Write-Host "`n=== ダウンロード結果 ===" -ForegroundColor Green
Write-Host "ダウンロード完了: $($downloadedFiles.Count) / $($latestFiles.Count) ファイル" -ForegroundColor Green

if ($downloadedFiles.Count -gt 0) {
    Write-Host "`nダウンロードされたファイル:" -ForegroundColor White
    $downloadedFiles | ForEach-Object {
        $sizeKB = [math]::Round($_.Size / 1024, 1)
        Write-Host "  [$($_.Type)] $($_.TypeName)" -ForegroundColor Cyan
        Write-Host "    ファイル名: $($_.FileName)" -ForegroundColor White
        Write-Host "    パス: $($_.FilePath)" -ForegroundColor White
        Write-Host "    サイズ: $sizeKB KB" -ForegroundColor White
        Write-Host "    更新日: $($_.Date)" -ForegroundColor White
        Write-Host ""
    }
    
    # CSVファイルとしてメタデータを保存
    $metadataPath = Join-Path $metaDir "yakkakijun_download_metadata.csv"
    $downloadedFiles | Export-Csv -Path $metadataPath -NoTypeInformation -Encoding UTF8
    Write-Host "メタデータを保存しました: $metadataPath" -ForegroundColor Green
}
else {
    Write-Warning "ダウンロードに成功したファイルはありませんでした"
}

Write-Host "`nスクリプトが完了しました" -ForegroundColor Green
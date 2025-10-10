<#
.SYNOPSIS
drug_data.jsonをHTMLに埋め込んだスタンドアロン版ビューアーを生成

.DESCRIPTION
Webサーバーを立てられない環境でも使用できるよう、JSONデータを
HTMLファイル内に直接埋め込んだスタンドアロン版を生成します。
生成されたHTMLファイルはダブルクリックで直接開けます。

.EXAMPLE
.\generate_standalone_viewer.ps1
# output/drug_viewer_standalone.html が生成されます

.EXAMPLE
.\generate_standalone_viewer.ps1 -OutputPath "standalone.html"
# 指定したパスに生成されます
#>

param(
    [string]$ViewerPath = "output/drug_viewer.html",
    [string]$DataPath = "output/drug_data.json",
    [string]$OutputPath = "output/drug_viewer_standalone.html"
)

# スクリプトの親ディレクトリ（プロジェクトルート）を取得
$ProjectRoot = Split-Path -Parent $PSScriptRoot

$ErrorActionPreference = "Stop"

Write-Host "=== スタンドアロン版ビューアー生成 ===" -ForegroundColor Cyan
Write-Host ""

# パスをプロジェクトルート基準に変換
if (-not [System.IO.Path]::IsPathRooted($ViewerPath)) {
    $ViewerPath = Join-Path $ProjectRoot $ViewerPath
}
if (-not [System.IO.Path]::IsPathRooted($DataPath)) {
    $DataPath = Join-Path $ProjectRoot $DataPath
}
if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path $ProjectRoot $OutputPath
}

# ファイル存在チェック
if (-not (Test-Path $ViewerPath)) {
    Write-Error "ビューアーHTMLが見つかりません: $ViewerPath"
    exit 1
}

if (-not (Test-Path $DataPath)) {
    Write-Error "データJSONが見つかりません: $DataPath"
    Write-Host "先に generate_drug_json.ps1 を実行してください" -ForegroundColor Yellow
    exit 1
}

Write-Host "ファイルを読み込み中..." -ForegroundColor Yellow
Write-Host "  ビューアーHTML: $ViewerPath" -ForegroundColor Gray
$htmlContent = Get-Content -Path $ViewerPath -Encoding UTF8 -Raw

Write-Host "  データJSON: $DataPath" -ForegroundColor Gray
$jsonContent = Get-Content -Path $DataPath -Encoding UTF8 -Raw

Write-Host "  読み込み完了" -ForegroundColor Green
Write-Host ""

# JSONデータのサイズ確認
$jsonSizeMB = [math]::Round((Get-Item $DataPath).Length / 1MB, 2)
Write-Host "JSONデータサイズ: $jsonSizeMB MB" -ForegroundColor Yellow

if ($jsonSizeMB -gt 50) {
    Write-Warning "JSONデータが大きいため、ブラウザの動作が重くなる可能性があります"
}

Write-Host ""
Write-Host "スタンドアロン版を生成中..." -ForegroundColor Yellow

# XMLHttpRequestのfetch部分を埋め込みデータに置き換え
$embeddedScript = @"
  <script>
    // 埋め込みデータ
    var embeddedDrugData = $jsonContent;
    
    var drugData = {};
    var filteredData = {};
    var currentFilter = 'all';
    var currentSort = 'name';
"@

# 元のスクリプト開始部分を置換
$htmlContent = $htmlContent -replace '<script>\s+var drugData = \{\};', $embeddedScript

# XMLHttpRequestの部分を削除して直接データを使用
$xhrPattern = '(?s)// データ読み込み.*?xhr\.send\(\);.*?function showLoadError\(\) \{.*?\}'
$replacement = @'
// データ読み込み（埋め込みデータを使用）
    try {
      drugData = embeddedDrugData;
      filteredData = drugData;
      updateStats();
      displayResults();
      document.getElementById('resultText').textContent = Object.keys(filteredData).length + ' 件の医薬品';
    } catch (error) {
      console.error('データ読み込みエラー:', error);
      document.getElementById('resultsContainer').innerHTML = 
        '<div class="no-results">' +
        '<div class="no-results-icon">⚠️</div>' +
        '<h2>データの読み込みに失敗しました</h2>' +
        '<p>埋め込みデータを確認してください</p>' +
        '</div>';
    }
'@

$htmlContent = $htmlContent -replace $xhrPattern, $replacement

# 出力ディレクトリ作成
$outputDir = Split-Path -Parent $OutputPath
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

# UTF8（BOMなし）で保存
Write-Host "  ファイル出力中: $OutputPath" -ForegroundColor Gray
[System.IO.File]::WriteAllText($OutputPath, $htmlContent, [System.Text.UTF8Encoding]::new($false))

$outputSizeMB = [math]::Round((Get-Item $OutputPath).Length / 1MB, 2)

Write-Host ""
Write-Host "=== 生成完了 ===" -ForegroundColor Green
Write-Host ""
Write-Host "出力ファイル: $OutputPath" -ForegroundColor White
Write-Host "ファイルサイズ: $outputSizeMB MB" -ForegroundColor White
Write-Host ""
Write-Host "使い方:" -ForegroundColor Cyan
Write-Host "  1. $OutputPath をダブルクリックしてブラウザで開く" -ForegroundColor Gray
Write-Host "  2. Webサーバー不要で直接動作します" -ForegroundColor Gray
Write-Host "  3. ファイルを配布することも可能です" -ForegroundColor Gray
Write-Host ""

# オプション：ブラウザで開くか確認
$openBrowser = Read-Host "ブラウザで開きますか？ (Y/N)"
if ($openBrowser -eq 'Y' -or $openBrowser -eq 'y') {
    Write-Host "ブラウザを開いています..." -ForegroundColor Yellow
    Start-Process (Resolve-Path $OutputPath).Path
}

Write-Host ""
Write-Host "✓ 完了" -ForegroundColor Green

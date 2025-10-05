# PMDA XMLファイルから薬剤サイズ情報を抽出するスクリプト

param(
    [string]$ZipPath = "data\pmda_zip\pmda_all_sgml_xml_20251005.zip",
    [string]$OutputDir = "data\raw"
)

Write-Host "=== PMDA XMLファイルからサイズ情報を抽出 ===" -ForegroundColor Cyan

# 出力ディレクトリの作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# ZIPファイルを開く
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)

Write-Host "ZIPファイルを開きました: $ZipPath" -ForegroundColor Yellow
Write-Host "XMLファイル数: $($zip.Entries.Count)" -ForegroundColor Green

$results = @()
$xmlEntries = $zip.Entries | Where-Object { $_.Name -like "*.xml" }
$totalFiles = $xmlEntries.Count
$processedFiles = 0

Write-Host "`nXMLファイルを処理中..." -ForegroundColor Yellow

foreach ($entry in $xmlEntries) {
    $processedFiles++
    
    if ($processedFiles % 100 -eq 0) {
        Write-Host "  処理中: $processedFiles / $totalFiles ファイル" -ForegroundColor Gray
    }
    
    try {
        # XMLファイルを読み取り
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream)
        $content = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()
        
        # XMLとしてパース
        [xml]$xmlDoc = $content
        
        # YJコードを取得
        $yjcodes = @()
        if ($xmlDoc.PackIns.ApprovalEtc.DetailBrandName) {
            foreach ($brand in $xmlDoc.PackIns.ApprovalEtc.DetailBrandName) {
                if ($brand.BrandCode.YJCode) {
                    $yjcodes += $brand.BrandCode.YJCode
                }
            }
        }
        
        # 包装情報を取得
        $package = ""
        if ($xmlDoc.PackIns.Package.Detail.Lang.'#text') {
            $package = $xmlDoc.PackIns.Package.Detail.Lang.'#text'
        }
        
        # 錠剤またはカプセルの場合のみ処理
        if ($package -match '錠|カプセル') {
            # すべてのセクションを検索してサイズ情報を探す
            $sizeInfo = $null
            $sectionName = ""
            
            foreach ($section in $xmlDoc.PackIns.ChildNodes) {
                if ($section.NodeType -eq "Element" -and $section.Detail) {
                    foreach ($detail in $section.Detail) {
                        if ($detail.Lang) {
                            $langNode = $detail.Lang
                            # Langノードが配列の場合も単一の場合も対応
                            $textNodes = @($langNode)
                            foreach ($langItem in $textNodes) {
                                $text = $langItem.'#text'
                                if ($text) {
                                    # より厳密なサイズ情報のパターン（形状や寸法を示すもの）
                                    if ($text -match '直径|長径|短径|厚[さみ]' -or 
                                        $text -match '([0-9]+\.?[0-9]*)\s*mm\s*[、×x]\s*([0-9]+\.?[0-9]*)\s*mm' -or
                                        $text -match '([0-9]+\.?[0-9]*)\s*ｍｍ\s*[、×x]\s*([0-9]+\.?[0-9]*)\s*ｍｍ') {
                                        # 誤検出を除外（血液検査などの単位）
                                        if (-not ($text -match '\d+/mm|mm3|mm2|mm Hg|mmHg')) {
                                            $sizeInfo = $text
                                            $sectionName = $section.LocalName
                                            break
                                        }
                                    }
                                }
                            }
                            if ($sizeInfo) { break }
                        }
                    }
                    if ($sizeInfo) { break }
                }
            }
            
            # サイズ情報が見つかった場合のみ結果に追加
            if ($sizeInfo) {
                # 文字列として扱う
                $sizeText = if ($sizeInfo -is [Array]) { $sizeInfo[0] } else { $sizeInfo }
                $sizeText = $sizeText.ToString().Trim()
                
                foreach ($yjcode in $yjcodes) {
                    # サイズ情報から数値を抽出（より柔軟なパターン）
                    $diameter = ""
                    $majorAxis = ""
                    $minorAxis = ""
                    $thickness = ""
                    
                    # 直径パターン
                    if ($sizeText -match '直径[：:\s]*([0-9]+\.?[0-9]*)\s*(?:mm|ｍｍ|ミリ)') {
                        $diameter = $matches[1]
                    }
                    # 長径パターン
                    if ($sizeText -match '長径[：:\s]*([0-9]+\.?[0-9]*)\s*(?:mm|ｍｍ|ミリ)') {
                        $majorAxis = $matches[1]
                    }
                    # 短径パターン
                    if ($sizeText -match '短径[：:\s]*([0-9]+\.?[0-9]*)\s*(?:mm|ｍｍ|ミリ)') {
                        $minorAxis = $matches[1]
                    }
                    # 厚さパターン
                    if ($sizeText -match '厚[さみ][：:\s]*([0-9]+\.?[0-9]*)\s*(?:mm|ｍｍ|ミリ)') {
                        $thickness = $matches[1]
                    }
                    
                    $results += [PSCustomObject]@{
                        FileName  = $entry.Name
                        YJCode    = $yjcode
                        Package   = $package
                        Section   = $sectionName
                        SizeInfo  = $sizeText
                        Diameter  = $diameter
                        MajorAxis = $majorAxis
                        MinorAxis = $minorAxis
                        Thickness = $thickness
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "ファイル処理エラー ($($entry.Name)): $($_.Exception.Message)"
    }
}

$zip.Dispose()

Write-Host "`n処理完了:" -ForegroundColor Green
Write-Host "  総XMLファイル数: $totalFiles" -ForegroundColor White
Write-Host "  サイズ情報あり: $($results.Count)" -ForegroundColor White

if ($results.Count -gt 0) {
    # CSV出力
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $OutputDir "pmda_medicine_size_$timestamp.csv"
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    
    Write-Host "`nCSVファイルを出力しました: $csvPath" -ForegroundColor Cyan
    
    # サンプルデータを表示
    Write-Host "`n=== サンプルデータ（最初の5件） ===" -ForegroundColor Cyan
    $results | Select-Object -First 5 | Format-Table YJCode, Diameter, MajorAxis, MinorAxis, Thickness -AutoSize
}
else {
    Write-Warning "サイズ情報が見つかりませんでした"
}

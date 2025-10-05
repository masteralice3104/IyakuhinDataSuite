# PMDA XML から YJコードとサイズ情報を抽出するスクリプト
# サイズ情報は SizeLongDiameter, SizeShortDiameter, SizeDiameter, SizeThickness を取得

param(
    [string]$ZipPath = "data\pmda_zip\pmda_all_sgml_xml_20251005.zip",
    [string]$OutputDir = "data\raw"
)

# 絶対パスに変換
$ZipPath = Join-Path $PSScriptRoot $ZipPath | Resolve-Path
$OutputDir = Join-Path $PSScriptRoot $OutputDir

# 出力ディレクトリが存在しない場合は作成
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# タイムスタンプ付きの出力ファイル名
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputCsv = Join-Path $OutputDir "pmda_yjcode_size_${timestamp}.csv"

Write-Host "PMDA XMLファイルからYJコードとサイズ情報を抽出します..."
Write-Host "ZIP: $ZipPath"
Write-Host "出力: $outputCsv"
Write-Host ""

# System.IO.Compression.FileSystemアセンブリをロード
Add-Type -AssemblyName System.IO.Compression.FileSystem

# 結果を格納する配列
$results = @()

# ZIPファイルを開く
$zip = [System.IO.Compression.ZipFile]::OpenRead($ZipPath)

$totalFiles = $zip.Entries.Count
$processedFiles = 0
$extractedCount = 0

try {
    foreach ($entry in $zip.Entries) {
        # XMLファイルのみ処理
        if ($entry.Name -notmatch '\.xml$') {
            continue
        }

        $processedFiles++

        # 進捗表示（500ファイルごと）
        if ($processedFiles % 500 -eq 0) {
            Write-Host "処理中: $processedFiles / $totalFiles ファイル (抽出: $extractedCount 件)"
        }

        try {
            # ZIPエントリからストリームを開く
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $xmlContent = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            # XML解析
            [xml]$xmlDoc = $xmlContent

            # 名前空間マネージャーを作成
            $nsManager = New-Object System.Xml.XmlNamespaceManager($xmlDoc.NameTable)
            $nsManager.AddNamespace("pi", "http://info.pmda.go.jp/namespace/prescription_drugs/package_insert/1.0")

            # 各ブランド（DetailBrandName）を処理
            $brandNodes = $xmlDoc.SelectNodes("//pi:DetailBrandName", $nsManager)
            
            foreach ($brandNode in $brandNodes) {
                # YJコードを取得
                $yjCodeNode = $brandNode.SelectSingleNode(".//pi:YJCode", $nsManager)
                if (-not $yjCodeNode) {
                    continue
                }
                $yjCode = $yjCodeNode.InnerText.Trim()

                # ブランド名を取得
                $brandNameNode = $brandNode.SelectSingleNode(".//pi:ApprovalBrandName/pi:Lang[@xml:lang='ja']", $nsManager)
                $brandName = if ($brandNameNode) { $brandNameNode.InnerText.Trim() } else { "" }

                # 親のPropertyForBrandを探す（同じrefを持つもの）
                $brandId = $brandNode.GetAttribute("id")
                if ($brandId) {
                    $propertyNode = $xmlDoc.SelectSingleNode("//pi:PropertyForBrand[@ref='$brandId']", $nsManager)
                    
                    if ($propertyNode) {
                        # Sizeノードを取得
                        $sizeNode = $propertyNode.SelectSingleNode(".//pi:Size", $nsManager)
                        
                        if ($sizeNode) {
                            # 各サイズ情報を取得
                            $sizeLongDiameter = ""
                            $sizeShortDiameter = ""
                            $sizeDiameter = ""
                            $sizeThickness = ""

                            $longDiameterNode = $sizeNode.SelectSingleNode(".//pi:SizeLongDiameter/pi:Lang[@xml:lang='ja']", $nsManager)
                            if ($longDiameterNode) {
                                $sizeLongDiameter = $longDiameterNode.InnerText.Trim()
                            }

                            $shortDiameterNode = $sizeNode.SelectSingleNode(".//pi:SizeShortDiameter/pi:Lang[@xml:lang='ja']", $nsManager)
                            if ($shortDiameterNode) {
                                $sizeShortDiameter = $shortDiameterNode.InnerText.Trim()
                            }

                            $diameterNode = $sizeNode.SelectSingleNode(".//pi:SizeDiameter/pi:Lang[@xml:lang='ja']", $nsManager)
                            if ($diameterNode) {
                                $sizeDiameter = $diameterNode.InnerText.Trim()
                            }

                            $thicknessNode = $sizeNode.SelectSingleNode(".//pi:SizeThickness/pi:Lang[@xml:lang='ja']", $nsManager)
                            if ($thicknessNode) {
                                $sizeThickness = $thicknessNode.InnerText.Trim()
                            }

                            # いずれかのサイズ情報があれば記録
                            if ($sizeLongDiameter -or $sizeShortDiameter -or $sizeDiameter -or $sizeThickness) {
                                $results += [PSCustomObject]@{
                                    FileName = $entry.Name
                                    YJCode = $yjCode
                                    BrandName = $brandName
                                    SizeLongDiameter = $sizeLongDiameter
                                    SizeShortDiameter = $sizeShortDiameter
                                    SizeDiameter = $sizeDiameter
                                    SizeThickness = $sizeThickness
                                }
                                $extractedCount++
                            }
                        }
                    }
                }
            }
        }
        catch {
            Write-Host "警告: ファイル $($entry.Name) の処理中にエラーが発生しました: $_"
        }
    }
}
finally {
    $zip.Dispose()
}

Write-Host ""
Write-Host "=== 処理完了 ==="
Write-Host "総XMLファイル数: $totalFiles"
Write-Host "抽出レコード数: $extractedCount"
Write-Host ""

# CSVに出力
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $outputCsv -Encoding UTF8 -NoTypeInformation
    Write-Host "CSVファイルを出力しました: $outputCsv"
    
    # サンプルを表示
    Write-Host ""
    Write-Host "=== 抽出サンプル（最初の5件）==="
    $results | Select-Object -First 5 | Format-Table -AutoSize
}
else {
    Write-Host "サイズ情報を含むレコードが見つかりませんでした。"
}

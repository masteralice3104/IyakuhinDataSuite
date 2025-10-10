# PMDA XML YJコード・サイズ情報抽出スクリプト
# pmda_zip内のZIPファイルからXMLを読み込み、医薬品名・YJコード・サイズ情報を抽出

param(
    [string]$ZipDir = "pmda_zip",
    [string]$OutputDir = "raw"
)

Write-Host "=== PMDA XML YJコード・サイズ情報抽出 ===" -ForegroundColor Cyan
Write-Host ""

# ZIPディレクトリの確認
$ZipDir = Join-Path $PSScriptRoot $ZipDir
if (-not (Test-Path $ZipDir)) {
    Write-Error "pmda_zipディレクトリが見つかりません: $ZipDir"
    exit 1
}

# ZIPファイルを検索
$zipFiles = Get-ChildItem -Path $ZipDir -Filter "*.zip"

if ($zipFiles.Count -eq 0) {
    Write-Error "ZIPファイルが見つかりませんでした"
    exit 1
}

# 最初のZIPファイルを使用
$zipFile = $zipFiles[0]
Write-Host "使用するZIPファイル: $($zipFile.Name)" -ForegroundColor Green
Write-Host "ファイルサイズ: $([math]::Round($zipFile.Length / 1MB, 2)) MB" -ForegroundColor Green
Write-Host ""

# 出力ディレクトリの作成
$OutputDir = Join-Path $PSScriptRoot $OutputDir
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

# 出力ファイルパス
$outputCsv = Join-Path $OutputDir "pmda.csv"

try {
    # System.IO.Compression.FileSystemアセンブリをロード
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    
    # 結果を格納する配列
    $results = @()
    
    # ZIPファイルを開く
    Write-Host "ZIPファイルを開いています..." -ForegroundColor Yellow
    $zip = [System.IO.Compression.ZipFile]::OpenRead($zipFile.FullName)
    
    $totalFiles = ($zip.Entries | Where-Object { $_.Name -match '\.xml$' }).Count
    $processedFiles = 0
    $extractedCount = 0
    
    Write-Host "総XMLファイル数: $totalFiles" -ForegroundColor Green
    Write-Host ""
    
    foreach ($entry in $zip.Entries) {
        # XMLファイルのみ処理
        if ($entry.Name -notmatch '\.xml$') {
            continue
        }
        
        $processedFiles++
        
        # 進捗表示（500ファイルごと）
        if ($processedFiles % 500 -eq 0) {
            $percentComplete = [math]::Round(($processedFiles / $totalFiles) * 100, 1)
            Write-Host "処理中: $processedFiles / $totalFiles ファイル ($percentComplete%) - 抽出: $extractedCount 件" -ForegroundColor Cyan
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
                                    名称       = $brandName
                                    個別医薬品コード = $yjCode
                                    長径       = $sizeLongDiameter
                                    短径       = $sizeShortDiameter
                                    直径       = $sizeDiameter
                                    厚さ       = $sizeThickness
                                }
                                $extractedCount++
                            }
                        }
                    }
                }
            }
        }
        catch {
            # 個別ファイルのエラーは警告として処理し、処理を継続
            if ($processedFiles % 1000 -eq 0) {
                Write-Host "  警告: ファイル $($entry.Name) の処理中にエラー: $($_.Exception.Message)" -ForegroundColor DarkYellow
            }
        }
    }
    
    $zip.Dispose()
    
    Write-Host ""
    Write-Host "=== 処理完了 ===" -ForegroundColor Green
    Write-Host "  処理XMLファイル数: $processedFiles" -ForegroundColor White
    Write-Host "  抽出レコード数: $extractedCount" -ForegroundColor White
    Write-Host ""
    
    # CSVに出力（重複を削除）
    if ($results.Count -gt 0) {
        Write-Host "重複データを削除中..." -ForegroundColor Yellow
        
        # YJコードとサイズ情報の組み合わせでユニーク化
        $uniqueResults = $results | Sort-Object 個別医薬品コード, 名称 | 
        Group-Object 個別医薬品コード, 長径, 短径, 直径, 厚さ | 
        ForEach-Object { $_.Group[0] }
        
        $duplicateCount = $results.Count - $uniqueResults.Count
        Write-Host "  重複削除: $duplicateCount 件" -ForegroundColor Green
        Write-Host "  ユニークレコード数: $($uniqueResults.Count) 件" -ForegroundColor Green
        Write-Host ""
        
        Write-Host "CSVファイルを保存中..." -ForegroundColor Yellow
        $uniqueResults | Export-Csv -Path $outputCsv -Encoding UTF8 -NoTypeInformation
        
        $fileInfo = Get-Item $outputCsv
        $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        
        Write-Host "  ✓ 保存完了: $fileSizeMB MB" -ForegroundColor Green
        Write-Host ""
        
        # サンプルを表示
        Write-Host "=== 抽出サンプル（最初の5件）===" -ForegroundColor Cyan
        $uniqueResults | Select-Object -First 5 | Format-Table -AutoSize
        
        Write-Host ""
        Write-Host "=== ダウンロード完了 ===" -ForegroundColor Green
        Write-Host "  保存先: $outputCsv" -ForegroundColor White
        Write-Host ""
        
        # 整理スクリプトを自動実行
        $processScript = Join-Path $PSScriptRoot "pmda_process.ps1"
        if (Test-Path $processScript) {
            Write-Host "CSV整形スクリプトを実行します..." -ForegroundColor Cyan
            & $processScript
        }
        else {
            Write-Warning "pmda_process.ps1 が見つかりませんでした。手動で実行してください。"
        }
    }
    else {
        Write-Warning "サイズ情報を含むレコードが見つかりませんでした。"
    }
}
catch {
    Write-Error "処理中にエラーが発生しました: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
    exit 1
}

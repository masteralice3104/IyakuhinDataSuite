# 医薬品コード双方向変換スクリプト
# 包装単位コード ⇔ 個別医薬品コード（YJコード）の相互変換

param(
    [Parameter(Mandatory = $false)]
    [string]$Code,
    
    [Parameter(Mandatory = $false)]
    [string]$InputFile,
    
    [switch]$Test,
    [switch]$ShowHelp
)

# ヘルプ表示
if ($ShowHelp -or (-not $Code -and -not $InputFile -and -not $Test)) {
    Write-Host @"
医薬品コード双方向変換スクリプト

使用方法:
  .\code_bidirectional.ps1 -Code <コード>
  .\code_bidirectional.ps1 -InputFile <ファイルパス>
  .\code_bidirectional.ps1 -Test
  .\code_bidirectional.ps1 -ShowHelp

オプション:
  -Code <コード>       変換するコード（自動判定）
  -InputFile <ファイル> 複数コードを含むファイル（1行1コード）
  -Test                循環変換テストを実行
  -ShowHelp            このヘルプを表示

例:
  # 包装単位コード → 個別医薬品コード
  .\code_bidirectional.ps1 -Code 04987934648347
  
  # 個別医薬品コード → 包装単位コード
  .\code_bidirectional.ps1 -Code 1290401G1021
  
  # 循環変換テスト
  .\code_bidirectional.ps1 -Test

入力可能なコード:
  - 調剤包装単位コード (14桁、先頭0)
  - 販売包装単位コード (14桁、先頭1)
  - 元梱包装単位コード (14桁、先頭2)
  - 個別医薬品コード（YJコード、12桁）
"@ -ForegroundColor White
    exit 0
}

# データファイルパス
$medhot = Join-Path $PSScriptRoot "csv\medhot.csv"
$medis = Join-Path $PSScriptRoot "csv\MEDIS20250930.csv"

# データファイルの存在確認
if (-not (Test-Path $medhot)) {
    Write-Error "medhot.csvが見つかりません: $medhot"
    exit 1
}

if (-not (Test-Path $medis)) {
    Write-Error "MEDIS20250930.csvが見つかりません: $medis"
    exit 1
}

# データ読み込み
Write-Host "=== 医薬品コード双方向変換 ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "データを読み込み中..." -ForegroundColor Yellow
$medhotData = Import-Csv -Path $medhot -Encoding UTF8
$medisData = Import-Csv -Path $medis -Encoding UTF8

# 薬価基準収載医薬品コード → 個別医薬品コードのマッピング（MEDISから）
# 注意: 統一名収載品（末尾01*）でもMEDISには個別医薬品コードが記録されている
$yjcodeMap = @{}
foreach ($row in $medisData) {
    $yakkakijun = $row.'薬価基準収載医薬品コード'
    $yjcode = $row.'個別医薬品コード'
    if ($yakkakijun -and $yjcode) {
        if (-not $yjcodeMap.ContainsKey($yakkakijun)) {
            $yjcodeMap[$yakkakijun] = @()
        }
        if ($yjcodeMap[$yakkakijun] -notcontains $yjcode) {
            $yjcodeMap[$yakkakijun] += $yjcode
        }
    }
}

# 個別医薬品コード → 薬価基準収載医薬品コードのマッピング（逆引き）
$reverseMap = @{}
foreach ($key in $yjcodeMap.Keys) {
    foreach ($yjcode in $yjcodeMap[$key]) {
        if (-not $reverseMap.ContainsKey($yjcode)) {
            $reverseMap[$yjcode] = @()
        }
        if ($reverseMap[$yjcode] -notcontains $key) {
            $reverseMap[$yjcode] += $key
        }
    }
}

Write-Host "  MEDHOT: $($medhotData.Count) 行" -ForegroundColor Green
Write-Host "  MEDIS: $($medisData.Count) 行" -ForegroundColor Green
Write-Host "  個別医薬品コード: $(($reverseMap.Keys | Measure-Object).Count) 種類" -ForegroundColor Green
Write-Host ""

# コード種類を判定する関数
function Get-CodeType {
    param([string]$code)
    
    if ($code.Length -eq 14) {
        switch ($code[0]) {
            '0' { return "調剤包装単位コード" }
            '1' { return "販売包装単位コード" }
            '2' { return "元梱包装単位コード" }
            default { return "不明な包装単位コード" }
        }
    }
    elseif ($code.Length -eq 12) {
        return "個別医薬品コード"
    }
    else {
        return "不明"
    }
}

# 包装単位コード → 個別医薬品コード
function ConvertFrom-PackagingCode {
    param([string]$code)
    
    $codeType = Get-CodeType -code $code
    
    $matchedRows = $medhotData | Where-Object {
        $_.'調剤包装単位コード' -eq $code -or
        $_.'販売包装単位コード' -eq $code -or
        $_.'元梱包装単位コード' -eq $code
    }
    
    if ($matchedRows.Count -eq 0) {
        return [PSCustomObject]@{
            入力コード   = $code
            入力コード種類 = $codeType
            出力コード   = ""
            出力コード種類 = ""
            販売名     = ""
            変換方向    = "包装→個別"
            状態      = "見つかりません"
        }
    }
    
    $results = @()
    foreach ($row in $matchedRows) {
        $yakkakijun = $row.'薬価コード'
        $productName = $row.'販売名'
        
        if ($yakkakijun -and $productName) {
            # MEDISマッピングを使用して個別医薬品コードを取得
            $yjcodes = $yjcodeMap[$yakkakijun]
            if ($yjcodes) {
                # 販売名が一致する個別医薬品コードのみを返す
                foreach ($yjcode in $yjcodes) {
                    # MEDISで該当する個別医薬品コードの販売名を確認
                    $medisRow = $medisData | Where-Object { 
                        $_.'個別医薬品コード' -eq $yjcode -and 
                        $_.'薬価基準収載医薬品コード' -eq $yakkakijun 
                    } | Select-Object -First 1
                    
                    if ($medisRow -and $medisRow.'販売名' -eq $productName) {
                        $results += [PSCustomObject]@{
                            入力コード   = $code
                            入力コード種類 = $codeType
                            出力コード   = $yjcode
                            出力コード種類 = "個別医薬品コード"
                            販売名     = $productName
                            変換方向    = "包装→個別"
                            状態      = "成功"
                        }
                    }
                }
            }
        }
    }
    
    if ($results.Count -eq 0) {
        return [PSCustomObject]@{
            入力コード   = $code
            入力コード種類 = $codeType
            出力コード   = ""
            出力コード種類 = ""
            販売名     = ""
            変換方向    = "包装→個別"
            状態      = "見つかりません"
        }
    }
    
    return $results
}

# 個別医薬品コード → 包装単位コード
function ConvertTo-PackagingCode {
    param([string]$code)
    
    $codeType = Get-CodeType -code $code
    
    # MEDISから個別医薬品コードの販売名を取得
    $medisRow = $medisData | Where-Object { $_.'個別医薬品コード' -eq $code } | Select-Object -First 1
    
    if (-not $medisRow) {
        return [PSCustomObject]@{
            入力コード   = $code
            入力コード種類 = $codeType
            出力コード   = ""
            出力コード種類 = ""
            販売名     = ""
            変換方向    = "個別→包装"
            状態      = "見つかりません"
        }
    }
    
    $productName = $medisRow.'販売名'
    $yakkakijunCodes = $reverseMap[$code]
    
    if (-not $yakkakijunCodes) {
        return [PSCustomObject]@{
            入力コード   = $code
            入力コード種類 = $codeType
            出力コード   = ""
            出力コード種類 = ""
            販売名     = $productName
            変換方向    = "個別→包装"
            状態      = "見つかりません"
        }
    }
    
    $results = @()
    foreach ($yakkakijun in $yakkakijunCodes) {
        # 薬価基準収載医薬品コード → 包装単位コード（MEDHOTから）
        # 販売名が一致するもののみを返す
        $matchedRows = $medhotData | Where-Object { 
            $_.'薬価コード' -eq $yakkakijun -and 
            $_.'販売名' -eq $productName 
        }
        
        foreach ($row in $matchedRows) {
            # 調剤包装単位コード
            if ($row.'調剤包装単位コード') {
                $results += [PSCustomObject]@{
                    入力コード   = $code
                    入力コード種類 = $codeType
                    出力コード   = $row.'調剤包装単位コード'
                    出力コード種類 = "調剤包装単位コード"
                    販売名     = $productName
                    変換方向    = "個別→包装"
                    状態      = "成功"
                }
            }
            # 販売包装単位コード
            if ($row.'販売包装単位コード') {
                $results += [PSCustomObject]@{
                    入力コード   = $code
                    入力コード種類 = $codeType
                    出力コード   = $row.'販売包装単位コード'
                    出力コード種類 = "販売包装単位コード"
                    販売名     = $productName
                    変換方向    = "個別→包装"
                    状態      = "成功"
                }
            }
            # 元梱包装単位コード
            if ($row.'元梱包装単位コード') {
                $results += [PSCustomObject]@{
                    入力コード   = $code
                    入力コード種類 = $codeType
                    出力コード   = $row.'元梱包装単位コード'
                    出力コード種類 = "元梱包装単位コード"
                    販売名     = $productName
                    変換方向    = "個別→包装"
                    状態      = "成功"
                }
            }
        }
    }
    
    return $results
}

# 単一コード変換
if ($Code) {
    $codeType = Get-CodeType -code $Code
    Write-Host "変換中: $Code ($codeType)" -ForegroundColor Yellow
    Write-Host ""
    
    if ($codeType -eq "個別医薬品コード") {
        $results = ConvertTo-PackagingCode -code $Code
    }
    else {
        $results = ConvertFrom-PackagingCode -code $Code
    }
    
    if ($results) {
        $results | Format-Table -AutoSize
        $successCount = ($results | Where-Object { $_.状態 -eq "成功" }).Count
        Write-Host "✓ 変換成功: $successCount 件" -ForegroundColor Green
    }
}

# 循環変換テスト
if ($Test) {
    Write-Host "=== ランダム循環変換テスト ===" -ForegroundColor Cyan
    Write-Host ""
    
    # 薬価コードが存在する包装単位コードのリストを高速に作成
    Write-Host "テストデータを準備中..." -ForegroundColor Yellow
    $validPackagingCodes = [System.Collections.Generic.HashSet[string]]::new()
    
    foreach ($row in $medhotData) {
        if ($row.'薬価コード') {
            if ($row.'調剤包装単位コード') { 
                [void]$validPackagingCodes.Add($row.'調剤包装単位コード')
            }
            if ($row.'販売包装単位コード') { 
                [void]$validPackagingCodes.Add($row.'販売包装単位コード')
            }
            if ($row.'元梱包装単位コード') { 
                [void]$validPackagingCodes.Add($row.'元梱包装単位コード')
            }
        }
    }
    
    $validPackagingCodesArray = @($validPackagingCodes)
    Write-Host "  利用可能な包装単位コード: $($validPackagingCodesArray.Count) 件" -ForegroundColor Green
    Write-Host ""
    
    # 全体の統計
    $totalTests = 0
    $totalSuccess = 0
    $totalFailures = 0
    $allFailures = @()
    
    # テスト用スクリプトブロック（並列実行用、メモリ効率化）
    $testScriptBlock = {
        param($testCode, $batch, $medhotPath, $medisPath)
        
        # 必要なデータのみ読み込み（メモリ効率化）
        $medhotRow = Import-Csv -Path $medhotPath -Encoding UTF8 | Where-Object {
            $_.'調剤包装単位コード' -eq $testCode -or
            $_.'販売包装単位コード' -eq $testCode -or
            $_.'元梱包装単位コード' -eq $testCode
        } | Select-Object -First 1
        
        if (-not $medhotRow) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                失敗理由 = "包装→個別: コードが見つからない"
                成功   = $false
            }
        }
        
        $yakkakijun = $medhotRow.'薬価コード'
        $productName = $medhotRow.'販売名'
        
        if (-not $yakkakijun -or -not $productName) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                失敗理由 = "包装→個別: 薬価コードまたは販売名が空"
                成功   = $false
            }
        }
        
        # MEDISデータから該当する個別医薬品コードを検索
        $medisRows = Import-Csv -Path $medisPath -Encoding UTF8 | Where-Object {
            $_.'薬価基準収載医薬品コード' -eq $yakkakijun -and
            $_.'販売名' -eq $productName
        }
        
        if ($medisRows.Count -eq 0) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                失敗理由 = "包装→個別: 販売名が一致する個別医薬品コードなし"
                成功   = $false
            }
        }
        
        $yjcode = $medisRows[0].'個別医薬品コード'
        
        # 個別 → 包装（逆変換チェック）
        $reverseRows = Import-Csv -Path $medisPath -Encoding UTF8 | Where-Object {
            $_.'個別医薬品コード' -eq $yjcode
        }
        
        if ($reverseRows.Count -eq 0) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                失敗理由 = "個別→包装: 薬価基準収載医薬品コードが見つからない"
                成功   = $false
            }
        }
        
        $yakkakijunCodes = $reverseRows.'薬価基準収載医薬品コード' | Select-Object -Unique
        $foundOriginal = $false
        $foundNames = @()
        
        foreach ($ykCode in $yakkakijunCodes) {
            $rows = Import-Csv -Path $medhotPath -Encoding UTF8 | Where-Object { 
                $_.'薬価コード' -eq $ykCode -and 
                $_.'販売名' -eq $productName 
            }
            
            foreach ($r in $rows) {
                if ($r.'調剤包装単位コード' -eq $testCode -or 
                    $r.'販売包装単位コード' -eq $testCode -or 
                    $r.'元梱包装単位コード' -eq $testCode) {
                    $foundOriginal = $true
                }
                
                if ($r.'販売名' -and $foundNames -notcontains $r.'販売名') {
                    $foundNames += $r.'販売名'
                }
            }
        }
        
        if (-not $foundOriginal) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                販売名  = $productName
                失敗理由 = "個別→包装: 元のコードが見つからない"
                成功   = $false
            }
        }
        
        if ($foundNames.Count -gt 1) {
            return [PSCustomObject]@{
                バッチ  = $batch
                コード  = $testCode
                販売名  = $productName
                失敗理由 = "個別→包装: 異なる販売名が混在"
                成功   = $false
            }
        }
        
        return [PSCustomObject]@{
            バッチ = $batch
            コード = $testCode
            販売名 = $productName
            成功  = $true
        }
    }
    
    # 10回のバッチテスト（並列実行）
    for ($batch = 1; $batch -le 10; $batch++) {
        Write-Host "=== バッチ $batch / 10 ===" -ForegroundColor Cyan
        $batchStart = Get-Date
        
        # ランダムに100個選択
        $randomCodes = $validPackagingCodesArray | Get-Random -Count 100
        
        # 並列ジョブ実行（10個ずつ処理して進捗表示）
        $jobs = @()
        $completed = 0
        
        for ($i = 0; $i -lt $randomCodes.Count; $i += 10) {
            $chunk = $randomCodes[$i..[Math]::Min($i + 9, $randomCodes.Count - 1)]
            
            Write-Host "  進捗: $completed / 100 完了..." -ForegroundColor Yellow -NoNewline
            Write-Host "`r" -NoNewline
            
            # 10個のジョブを起動
            $chunkJobs = @()
            foreach ($testCode in $chunk) {
                $chunkJobs += Start-Job -ScriptBlock $testScriptBlock -ArgumentList $testCode, $batch, $medhot, $medis
            }
            
            # このチャンクのジョブ完了を待機
            $chunkJobs | Wait-Job | Out-Null
            $jobs += $chunkJobs
            
            $completed += $chunk.Count
        }
        
        Write-Host "  進捗: 100 / 100 完了      " -ForegroundColor Green
        
        # 結果収集
        $results = $jobs | Receive-Job
        $jobs | Remove-Job
        
        # 結果集計
        $batchSuccess = ($results | Where-Object { $_.成功 }).Count
        $batchFailures = $results | Where-Object { -not $_.成功 }
        
        $totalTests += $randomCodes.Count
        $totalSuccess += $batchSuccess
        $totalFailures += $batchFailures.Count
        $allFailures += $batchFailures
        
        $batchEnd = Get-Date
        $elapsed = ($batchEnd - $batchStart).TotalSeconds
        
        # バッチ結果表示
        if ($batchFailures.Count -eq 0) {
            Write-Host "  ✓ バッチ $batch 完了: 100/100 成功 ($([math]::Round($elapsed, 1))秒)" -ForegroundColor Green
        }
        else {
            Write-Host "  ⚠ バッチ $batch 完了: $batchSuccess/100 成功, $($batchFailures.Count) 失敗 ($([math]::Round($elapsed, 1))秒)" -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # 最終結果
    Write-Host "=== 最終結果 ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "総テスト数: $totalTests" -ForegroundColor White
    Write-Host "成功: $totalSuccess" -ForegroundColor Green
    Write-Host "失敗: $totalFailures" -ForegroundColor $(if ($totalFailures -eq 0) { "Green" } else { "Red" })
    Write-Host "成功率: $([math]::Round($totalSuccess / $totalTests * 100, 2))%" -ForegroundColor $(if ($totalFailures -eq 0) { "Green" } else { "Yellow" })
    Write-Host ""
    
    if ($allFailures.Count -gt 0) {
        Write-Host "=== 失敗の詳細 ===" -ForegroundColor Red
        $allFailures | Format-Table -AutoSize
        Write-Error "一部のテストが失敗しました"
        exit 1
    }
    else {
        Write-Host "=== すべてのテストが成功しました！ ===" -ForegroundColor Green
        Write-Host "  1000個の包装単位コードで循環変換テストが成功" -ForegroundColor White
        Write-Host "  - 包装単位コード → 個別医薬品コード → 包装単位コード" -ForegroundColor White
        Write-Host "  - すべてのケースで元のコードに戻ることを確認" -ForegroundColor White
        Write-Host "  - すべてのケースで販売名が一致することを確認" -ForegroundColor White
    }
}

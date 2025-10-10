<#
.SYNOPSIS
医薬品データ統合システム - メインメニュー

.DESCRIPTION
対話式のCUIメニューから各機能を実行できる統合スクリプトです。
データ収集、処理、変換、ビューアー起動などすべての機能にアクセスできます。

.EXAMPLE
.\main.ps1
# 対話式メニューが起動します
#>

$ErrorActionPreference = "Stop"

# カラースキーム
$colors = @{
    Title     = "Cyan"
    Menu      = "White"
    Highlight = "Yellow"
    Success   = "Green"
    Warning   = "Yellow"
    Error     = "Red"
    Info      = "Gray"
}

# バナー表示
function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ╔════════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
    Write-Host "  ║                                                                ║" -ForegroundColor Cyan
    Write-Host "  ║        " -NoNewline -ForegroundColor Cyan
    Write-Host "💊 医薬品データ統合システム" -NoNewline -ForegroundColor White
    Write-Host " (IyakuhinDataSuite)" -NoNewline -ForegroundColor Yellow
    Write-Host "     ║" -ForegroundColor Cyan
    Write-Host "  ║                                                                ║" -ForegroundColor Cyan
    Write-Host "  ╚════════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
    Write-Host ""
}

# メインメニュー表示
function Show-MainMenu {
    Show-Banner
    Write-Host "  ┌────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
    Write-Host "📋 メインメニュー" -NoNewline -ForegroundColor Yellow
    Write-Host "                                              │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🚀 " -NoNewline -ForegroundColor Green
    Write-Host "[0] 全自動セットアップ（推奨・初回実行）" -ForegroundColor Green
    Write-Host "         └─ データ収集→JSON生成→スタンドアロン版生成" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    📥 " -NoNewline -ForegroundColor Cyan
    Write-Host "[1] データ収集・更新" -ForegroundColor White
    Write-Host "    📄 " -NoNewline -ForegroundColor Cyan
    Write-Host "[2] JSON生成" -ForegroundColor White
    Write-Host "    🌐 " -NoNewline -ForegroundColor Cyan
    Write-Host "[3] ビューアー起動" -ForegroundColor White
    Write-Host "    🔄 " -NoNewline -ForegroundColor Cyan
    Write-Host "[4] コード変換" -ForegroundColor White
    Write-Host "    ℹ️  " -NoNewline -ForegroundColor Cyan
    Write-Host "[5] システム情報" -ForegroundColor White
    Write-Host ""
    Write-Host "    ❌ " -NoNewline -ForegroundColor Red
    Write-Host "[Q] 終了" -ForegroundColor White
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
}

# データ収集メニュー
function Show-DataCollectionMenu {
    Show-Banner
    Write-Host "  ┌────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
    Write-Host "📥 データ収集・更新" -NoNewline -ForegroundColor Yellow
    Write-Host "                                        │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🔹 " -NoNewline -ForegroundColor Cyan
    Write-Host "[1] MEDHOTデータ取得・処理" -ForegroundColor White
    Write-Host "    🔹 " -NoNewline -ForegroundColor Cyan
    Write-Host "[2] HOTコードマスター取得・処理" -ForegroundColor White
    Write-Host "    🔹 " -NoNewline -ForegroundColor Cyan
    Write-Host "[3] PMDAデータ取得・処理" -ForegroundColor White
    Write-Host ""
    Write-Host "    ⭐ " -NoNewline -ForegroundColor Green
    Write-Host "[4] すべて実行（推奨）" -ForegroundColor Green
    Write-Host ""
    Write-Host "    ⬅️  " -NoNewline -ForegroundColor Yellow
    Write-Host "[B] 戻る" -ForegroundColor White
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
}

# ビューアーメニュー
function Show-ViewerMenu {
    Show-Banner
    Write-Host "  ┌────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
    Write-Host "🌐 ビューアー起動" -NoNewline -ForegroundColor Yellow
    Write-Host "                                          │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🖥️  " -NoNewline -ForegroundColor Cyan
    Write-Host "[1] HTTPサーバー起動（通常版）" -ForegroundColor White
    Write-Host "         └─ http://localhost:8080 でビューアーを起動" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    📦 " -NoNewline -ForegroundColor Magenta
    Write-Host "[2] スタンドアロン版を生成" -ForegroundColor White
    Write-Host "         └─ サーバー不要の単一HTMLファイル作成" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🚀 " -NoNewline -ForegroundColor Green
    Write-Host "[3] スタンドアロン版を開く" -ForegroundColor White
    Write-Host "         └─ 既存のスタンドアロン版をブラウザで開く" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    ⬅️  " -NoNewline -ForegroundColor Yellow
    Write-Host "[B] 戻る" -ForegroundColor White
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
}

# コード変換メニュー
function Show-CodeConversionMenu {
    Show-Banner
    Write-Host "  ┌────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
    Write-Host "🔄 コード変換" -NoNewline -ForegroundColor Yellow
    Write-Host "                                              │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🔍 " -NoNewline -ForegroundColor Cyan
    Write-Host "[1] 単一コード変換" -ForegroundColor White
    Write-Host "         └─ GS1コード ⇔ YJコード の双方向変換" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    🧪 " -NoNewline -ForegroundColor Magenta
    Write-Host "[2] 循環変換テスト（1000件）" -ForegroundColor White
    Write-Host "         └─ 変換精度の検証（GS1→YJ→GS1）" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "    ⬅️  " -NoNewline -ForegroundColor Yellow
    Write-Host "[B] 戻る" -ForegroundColor White
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
}

# システム情報表示
function Show-SystemInfo {
    Show-Banner
    Write-Host "システム情報" -ForegroundColor $colors.Highlight
    Write-Host "─────────────────────────────────────────────────────────────" -ForegroundColor $colors.Info
    Write-Host ""
    
    # データファイルの状態確認
    $files = @{
        "MEDHOT CSV" = "csv\medhot.csv"
        "MEDIS CSV"  = "csv\MEDIS20250930.csv"
        "PMDA CSV"   = "csv\pmda.csv"
        "統合JSON"     = "output\drug_data.json"
        "ビューアーHTML"  = "output\drug_viewer.html"
        "スタンドアロン版"   = "output\drug_viewer_standalone.html"
    }
    
    Write-Host "データファイル:" -ForegroundColor $colors.Menu
    foreach ($name in $files.Keys) {
        $path = $files[$name]
        if (Test-Path $path) {
            $size = [math]::Round((Get-Item $path).Length / 1MB, 2)
            $modified = (Get-Item $path).LastWriteTime.ToString("yyyy/MM/dd HH:mm")
            Write-Host "  ✓ $name" -ForegroundColor $colors.Success -NoNewline
            Write-Host " ($size MB, $modified)" -ForegroundColor $colors.Info
        }
        else {
            Write-Host "  ✗ $name" -ForegroundColor $colors.Error -NoNewline
            Write-Host " (未作成)" -ForegroundColor $colors.Info
        }
    }
    
    Write-Host ""
    
    Write-Host "スクリプト:" -ForegroundColor $colors.Menu
    $scripts = @(
        "scripts\medhot_get.ps1", "scripts\medhot_process.ps1",
        "scripts\hotcode_get.ps1", "scripts\hotcode_process.ps1",
        "scripts\pmda_get.ps1", "scripts\pmda_process.ps1",
        "scripts\code_bidirectional.ps1",
        "scripts\generate_drug_json.ps1",
        "scripts\generate_standalone_viewer.ps1",
        "scripts\start_viewer.ps1"
    )
    
    foreach ($script in $scripts) {
        $scriptName = Split-Path $script -Leaf
        if (Test-Path $script) {
            Write-Host "  ✓ $scriptName" -ForegroundColor $colors.Success
        }
        else {
            Write-Host "  ✗ $scriptName" -ForegroundColor $colors.Error
        }
    }    Write-Host ""
    Write-Host "PowerShell バージョン: $($PSVersionTable.PSVersion)" -ForegroundColor $colors.Info
    Write-Host ""
    
    Pause
}

# スクリプト実行関数
function Invoke-ScriptWithConfirmation {
    param(
        [string]$ScriptPath,
        [string]$Description,
        [string[]]$Arguments = @()
    )
    
    Write-Host ""
    Write-Host "実行中: $Description" -ForegroundColor $colors.Highlight
    Write-Host "スクリプト: $ScriptPath" -ForegroundColor $colors.Info
    Write-Host ""
    
    if (-not (Test-Path $ScriptPath)) {
        Write-Host "エラー: スクリプトが見つかりません: $ScriptPath" -ForegroundColor $colors.Error
        Pause
        return $false
    }
    
    try {
        if ($Arguments.Count -gt 0) {
            & $ScriptPath @Arguments
        }
        else {
            & $ScriptPath
        }
        
        Write-Host ""
        Write-Host "✓ 完了: $Description" -ForegroundColor $colors.Success
        Write-Host ""
        Pause
        return $true
    }
    catch {
        Write-Host ""
        Write-Host "✗ エラー: $Description" -ForegroundColor $colors.Error
        Write-Host $_.Exception.Message -ForegroundColor $colors.Error
        Write-Host ""
        Pause
        return $false
    }
}

# スクリプト実行関数（自動モード - Pauseなし）
function Invoke-ScriptAuto {
    param(
        [string]$ScriptPath,
        [string]$Description
    )
    
    Write-Host ""
    Write-Host "▶ 実行中: $Description" -ForegroundColor Yellow
    
    if (-not (Test-Path $ScriptPath)) {
        Write-Host "✗ エラー: スクリプトが見つかりません: $ScriptPath" -ForegroundColor Red
        return $false
    }
    
    try {
        & $ScriptPath
        Write-Host "✓ 完了: $Description" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "✗ エラー: $Description" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        return $false
    }
}

# 全自動セットアップ
function Invoke-AutoSetup {
    Show-Banner
    Write-Host "  ┌────────────────────────────────────────────────────────────┐" -ForegroundColor DarkGray
    Write-Host "  │ " -NoNewline -ForegroundColor DarkGray
    Write-Host "🚀 全自動セットアップ" -NoNewline -ForegroundColor Green
    Write-Host "                                      │" -ForegroundColor DarkGray
    Write-Host "  └────────────────────────────────────────────────────────────┘" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  以下の処理を順番に実行します:" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "    1️⃣  MEDHOTデータ取得・処理" -ForegroundColor White
    Write-Host "    2️⃣  HOTコードマスター取得・処理" -ForegroundColor White
    Write-Host "    3️⃣  PMDAデータ取得・処理" -ForegroundColor White
    Write-Host "    4️⃣  JSON統合ファイル生成" -ForegroundColor White
    Write-Host "    5️⃣  スタンドアロン版ビューアー生成" -ForegroundColor White
    Write-Host ""
    Write-Host "  ⚠️  " -NoNewline -ForegroundColor Yellow
    Write-Host "処理には10分以上かかる場合があります" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    
    $confirm = Read-Host "実行しますか？ (Y/N)"
    if ($confirm -ne "Y" -and $confirm -ne "y") {
        Write-Host "キャンセルしました" -ForegroundColor Yellow
        Pause
        return
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " 全自動セットアップ開始" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $startTime = Get-Date
    $stepNumber = 1
    $totalSteps = 8
    
    # Step 1: MEDHOTデータ取得
    Write-Host "[$stepNumber/$totalSteps] MEDHOTデータ取得中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\medhot_get.ps1" "MEDHOTデータ取得"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 2: MEDHOTデータ処理
    Write-Host "[$stepNumber/$totalSteps] MEDHOTデータ処理中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\medhot_process.ps1" "MEDHOTデータ処理"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 3: HOTコードマスター取得
    Write-Host "[$stepNumber/$totalSteps] HOTコードマスター取得中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\hotcode_get.ps1" "HOTコードマスター取得"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 4: HOTコードマスター処理
    Write-Host "[$stepNumber/$totalSteps] HOTコードマスター処理中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\hotcode_process.ps1" "HOTコードマスター処理"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 5: PMDAデータ取得
    Write-Host "[$stepNumber/$totalSteps] PMDAデータ取得中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\pmda_get.ps1" "PMDAデータ取得"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 6: PMDAデータ処理
    Write-Host "[$stepNumber/$totalSteps] PMDAデータ処理中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\pmda_process.ps1" "PMDAデータ処理"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 7: JSON生成
    Write-Host "[$stepNumber/$totalSteps] JSON統合ファイル生成中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\generate_drug_json.ps1" "JSON生成"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    $stepNumber++
    
    # Step 8: スタンドアロン版生成
    Write-Host "[$stepNumber/$totalSteps] スタンドアロン版ビューアー生成中..." -ForegroundColor Cyan
    $result = Invoke-ScriptAuto ".\scripts\generate_standalone_viewer.ps1" "スタンドアロン版生成"
    if (-not $result) {
        Write-Host ""
        Write-Host "❌ セットアップを中断しました" -ForegroundColor Red
        Pause
        return
    }
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " ✅ 全自動セットアップ完了！" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "  所要時間: $([math]::Round($duration.TotalMinutes, 1)) 分" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  生成されたファイル:" -ForegroundColor Yellow
    Write-Host "    📄 output/drug_data.json (統合データ)" -ForegroundColor White
    Write-Host "    🌐 output/drug_viewer.html (通常版ビューアー)" -ForegroundColor White
    Write-Host "    📦 output/drug_viewer_standalone.html (スタンドアロン版)" -ForegroundColor White
    Write-Host ""
    Write-Host "  次のステップ:" -ForegroundColor Yellow
    Write-Host "    • メインメニュー [3] からビューアーを起動" -ForegroundColor White
    Write-Host "    • または output/drug_viewer_standalone.html をダブルクリック" -ForegroundColor White
    Write-Host ""
    
    Pause
}

# データ収集処理
function Invoke-DataCollection {
    while ($true) {
        Show-DataCollectionMenu
        $choice = Read-Host "選択してください"
        
        switch ($choice.ToUpper()) {
            "1" {
                Invoke-ScriptWithConfirmation ".\scripts\medhot_get.ps1" "MEDHOTデータ取得"
                Invoke-ScriptWithConfirmation ".\scripts\medhot_process.ps1" "MEDHOTデータ処理"
            }
            "2" {
                Invoke-ScriptWithConfirmation ".\scripts\hotcode_get.ps1" "HOTコードマスター取得"
                Invoke-ScriptWithConfirmation ".\scripts\hotcode_process.ps1" "HOTコードマスター処理"
            }
            "3" {
                Invoke-ScriptWithConfirmation ".\scripts\pmda_get.ps1" "PMDAデータ取得"
                Invoke-ScriptWithConfirmation ".\scripts\pmda_process.ps1" "PMDAデータ処理"
            }
            "4" {
                Write-Host ""
                Write-Host "すべてのデータを取得・処理します..." -ForegroundColor $colors.Highlight
                Write-Host "この処理には時間がかかる場合があります。" -ForegroundColor $colors.Warning
                Write-Host ""
                $confirm = Read-Host "実行しますか？ (Y/N)"
                
                if ($confirm -eq "Y" -or $confirm -eq "y") {
                    Invoke-ScriptWithConfirmation ".\scripts\medhot_get.ps1" "MEDHOTデータ取得"
                    Invoke-ScriptWithConfirmation ".\scripts\medhot_process.ps1" "MEDHOTデータ処理"
                    Invoke-ScriptWithConfirmation ".\scripts\hotcode_get.ps1" "HOTコードマスター取得"
                    Invoke-ScriptWithConfirmation ".\scripts\hotcode_process.ps1" "HOTコードマスター処理"
                    Invoke-ScriptWithConfirmation ".\scripts\pmda_get.ps1" "PMDAデータ取得"
                    Invoke-ScriptWithConfirmation ".\scripts\pmda_process.ps1" "PMDAデータ処理"
                }
            }
            "B" { return }
            default {
                Write-Host "無効な選択です" -ForegroundColor $colors.Error
                Start-Sleep -Seconds 1
            }
        }
    }
}

# ビューアー処理
function Invoke-Viewer {
    while ($true) {
        Show-ViewerMenu
        $choice = Read-Host "選択してください"
        
        switch ($choice.ToUpper()) {
            "1" {
                if (-not (Test-Path "output\drug_data.json")) {
                    Write-Host ""
                    Write-Host "エラー: drug_data.json が見つかりません" -ForegroundColor $colors.Error
                    Write-Host "先に [2] JSON生成 を実行してください" -ForegroundColor $colors.Warning
                    Write-Host ""
                    Pause
                }
                else {
                    Write-Host ""
                    Write-Host "HTTPサーバーを起動します..." -ForegroundColor $colors.Highlight
                    Write-Host "終了するには Ctrl+C を押してください" -ForegroundColor $colors.Warning
                    Write-Host ""
                    Start-Sleep -Seconds 2
                    & ".\scripts\start_viewer.ps1"
                }
            }
            "2" {
                if (-not (Test-Path "output\drug_data.json")) {
                    Write-Host ""
                    Write-Host "エラー: drug_data.json が見つかりません" -ForegroundColor $colors.Error
                    Write-Host "先に [2] JSON生成 を実行してください" -ForegroundColor $colors.Warning
                    Write-Host ""
                    Pause
                }
                else {
                    Invoke-ScriptWithConfirmation ".\scripts\generate_standalone_viewer.ps1" "スタンドアロン版生成"
                }
            }
            "3" {
                if (Test-Path "output\drug_viewer_standalone.html") {
                    Write-Host ""
                    Write-Host "スタンドアロン版を開きます..." -ForegroundColor $colors.Highlight
                    Start-Process (Resolve-Path "output\drug_viewer_standalone.html").Path
                    Start-Sleep -Seconds 1
                }
                else {
                    Write-Host ""
                    Write-Host "エラー: スタンドアロン版が見つかりません" -ForegroundColor $colors.Error
                    Write-Host "先に [2] スタンドアロン版を生成 を実行してください" -ForegroundColor $colors.Warning
                    Write-Host ""
                    Pause
                }
            }
            "B" { return }
            default {
                Write-Host "無効な選択です" -ForegroundColor $colors.Error
                Start-Sleep -Seconds 1
            }
        }
    }
}

# コード変換処理
function Invoke-CodeConversion {
    while ($true) {
        Show-CodeConversionMenu
        $choice = Read-Host "選択してください"
        
        switch ($choice.ToUpper()) {
            "1" {
                Write-Host ""
                Write-Host "コード変換" -ForegroundColor $colors.Highlight
                Write-Host ""
                Write-Host "入力可能なコード:" -ForegroundColor $colors.Menu
                Write-Host "  - 包装単位コード（14桁GS1コード）" -ForegroundColor $colors.Info
                Write-Host "  - 個別医薬品コード（12桁YJコード）" -ForegroundColor $colors.Info
                Write-Host ""
                $code = Read-Host "コードを入力してください"
                
                if ($code) {
                    Invoke-ScriptWithConfirmation ".\scripts\code_bidirectional.ps1" "コード変換" @("-Code", $code)
                }
            }
            "2" {
                Write-Host ""
                Write-Host "循環変換テスト（1000件）" -ForegroundColor $colors.Highlight
                Write-Host "ランダムに100件 × 10バッチのテストを実行します" -ForegroundColor $colors.Info
                Write-Host "この処理には約10分かかります" -ForegroundColor $colors.Warning
                Write-Host ""
                $confirm = Read-Host "実行しますか？ (Y/N)"
                
                if ($confirm -eq "Y" -or $confirm -eq "y") {
                    Invoke-ScriptWithConfirmation ".\scripts\code_bidirectional.ps1" "循環変換テスト" @("-Test")
                }
            }
            "B" { return }
            default {
                Write-Host "無効な選択です" -ForegroundColor $colors.Error
                Start-Sleep -Seconds 1
            }
        }
    }
}

# メインループ
while ($true) {
    Show-MainMenu
    $choice = Read-Host "選択してください"
    
    switch ($choice.ToUpper()) {
        "0" {
            Invoke-AutoSetup
        }
        "1" {
            Invoke-DataCollection
        }
        "2" {
            if (-not (Test-Path "csv\medhot.csv") -or 
                -not (Test-Path "csv\MEDIS20250930.csv") -or
                -not (Test-Path "csv\pmda.csv")) {
                Write-Host ""
                Write-Host "エラー: 必要なCSVファイルが見つかりません" -ForegroundColor $colors.Error
                Write-Host "先に [1] データ収集・更新 を実行してください" -ForegroundColor $colors.Warning
                Write-Host ""
                Pause
            }
            else {
                Invoke-ScriptWithConfirmation ".\scripts\generate_drug_json.ps1" "JSON生成"
            }
        }
        "3" {
            Invoke-Viewer
        }
        "4" {
            if (-not (Test-Path "csv\medhot.csv") -or 
                -not (Test-Path "csv\MEDIS20250930.csv")) {
                Write-Host ""
                Write-Host "エラー: 必要なCSVファイルが見つかりません" -ForegroundColor $colors.Error
                Write-Host "先に [1] データ収集・更新 を実行してください" -ForegroundColor $colors.Warning
                Write-Host ""
                Pause
            }
            else {
                Invoke-CodeConversion
            }
        }
        "5" {
            Show-SystemInfo
        }
        "Q" {
            Show-Banner
            Write-Host "医薬品データ統合システムを終了します" -ForegroundColor $colors.Highlight
            Write-Host ""
            exit 0
        }
        default {
            Write-Host "無効な選択です" -ForegroundColor $colors.Error
            Start-Sleep -Seconds 1
        }
    }
}

<#
.SYNOPSIS
医薬品データビューアー用のローカルHTTPサーバーを起動

.DESCRIPTION
outputフォルダでHTTPサーバーを起動し、ブラウザで drug_viewer.html を開きます。
これによりCORSエラーを回避してJSONファイルを読み込めます。

.EXAMPLE
.\start_viewer.ps1
# http://localhost:8080 でサーバーが起動し、ブラウザが開きます
#>

param(
    [int]$Port = 8080
)

# スクリプトの親ディレクトリ（プロジェクトルート）を取得
$ProjectRoot = Split-Path -Parent $PSScriptRoot

Write-Host "=== 医薬品データビューアー ===" -ForegroundColor Cyan
Write-Host ""

# outputディレクトリの存在確認
$outputPath = Join-Path $ProjectRoot "output"
if (-not (Test-Path $outputPath)) {
    Write-Error "output ディレクトリが見つかりません"
    exit 1
}

# drug_viewer.htmlの存在確認
$htmlPath = Join-Path $outputPath "drug_viewer.html"
if (-not (Test-Path $htmlPath)) {
    Write-Error "output/drug_viewer.html が見つかりません"
    exit 1
}

# drug_data.jsonの存在確認
$jsonPath = Join-Path $outputPath "drug_data.json"
if (-not (Test-Path $jsonPath)) {
    Write-Error "output/drug_data.json が見つかりません"
    Write-Host "先に generate_drug_json.ps1 を実行してください" -ForegroundColor Yellow
    exit 1
}

Write-Host "ローカルHTTPサーバーを起動中..." -ForegroundColor Yellow
Write-Host "  ポート: $Port" -ForegroundColor Gray
Write-Host "  ディレクトリ: output/" -ForegroundColor Gray
Write-Host ""

# HTTPサーバー起動用スクリプトブロック
$serverScript = {
    param($Port, $Path)
    
    # HttpListenerを使用したシンプルなHTTPサーバー
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add("http://localhost:$Port/")
    $listener.Start()
    
    Write-Host "✓ サーバー起動: http://localhost:$Port/" -ForegroundColor Green
    Write-Host ""
    Write-Host "終了するには Ctrl+C を押してください" -ForegroundColor Yellow
    Write-Host ""
    
    try {
        while ($listener.IsListening) {
            $context = $listener.GetContext()
            $request = $context.Request
            $response = $context.Response
            
            # URLパスを取得
            $urlPath = $request.Url.LocalPath
            if ($urlPath -eq '/') {
                $urlPath = '/drug_viewer.html'
            }
            
            # ファイルパス
            $filePath = Join-Path $Path $urlPath.TrimStart('/')
            
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - $($request.HttpMethod) $urlPath" -ForegroundColor Gray
            
            if (Test-Path $filePath) {
                # Content-Type設定
                $contentType = 'text/html'
                if ($filePath -match '\.json$') {
                    $contentType = 'application/json'
                }
                elseif ($filePath -match '\.css$') {
                    $contentType = 'text/css'
                }
                elseif ($filePath -match '\.js$') {
                    $contentType = 'application/javascript'
                }
                
                $response.ContentType = "$contentType; charset=utf-8"
                $content = [System.IO.File]::ReadAllBytes($filePath)
                $response.ContentLength64 = $content.Length
                $response.OutputStream.Write($content, 0, $content.Length)
            }
            else {
                # 404
                $response.StatusCode = 404
                $html = '<html><body><h1>404 Not Found</h1></body></html>'
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
            }
            
            $response.Close()
        }
    }
    finally {
        $listener.Stop()
    }
}

# バックグラウンドジョブとしてサーバーを起動
$outputFullPath = (Resolve-Path $outputPath).Path
$job = Start-Job -ScriptBlock $serverScript -ArgumentList $Port, $outputFullPath

# サーバー起動を少し待つ
Start-Sleep -Seconds 2

# ジョブの出力を表示
Receive-Job -Job $job

# ブラウザを開く
$url = "http://localhost:$Port/drug_viewer.html"
Write-Host "ブラウザを開いています..." -ForegroundColor Yellow
Write-Host "  URL: $url" -ForegroundColor Gray
Start-Process $url

Write-Host ""
Write-Host "=== サーバー実行中 ===" -ForegroundColor Green
Write-Host "終了するには Ctrl+C を押してください" -ForegroundColor Yellow
Write-Host ""

# ジョブの出力をリアルタイムで表示
try {
    while ($true) {
        Start-Sleep -Milliseconds 500
        Receive-Job -Job $job | ForEach-Object { Write-Host $_ }
        
        # ジョブが終了していたら抜ける
        if ($job.State -ne 'Running') {
            break
        }
    }
}
finally {
    # クリーンアップ
    Write-Host ""
    Write-Host "サーバーを停止しています..." -ForegroundColor Yellow
    Stop-Job -Job $job
    Remove-Job -Job $job
    Write-Host "✓ サーバー停止" -ForegroundColor Green
}

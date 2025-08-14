# RemoteWakeConnect ビルドスクリプト

Write-Host "RemoteWakeConnect ビルドスクリプト" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan

# プロジェクトディレクトリに移動
$projectDir = $PSScriptRoot
Set-Location $projectDir

# ビルド構成を選択
$buildConfig = "Release"
Write-Host "`nビルド構成: $buildConfig" -ForegroundColor Yellow

# クリーンビルド
Write-Host "`n古いビルド成果物をクリーンアップ中..." -ForegroundColor Yellow
dotnet clean src/RemoteWakeConnect.csproj -c $buildConfig

# リストア
Write-Host "`n依存関係を復元中..." -ForegroundColor Yellow
dotnet restore src/RemoteWakeConnect.csproj

# ビルド
Write-Host "`nプロジェクトをビルド中..." -ForegroundColor Yellow
dotnet build src/RemoteWakeConnect.csproj -c $buildConfig

if ($LASTEXITCODE -ne 0) {
    Write-Host "`nビルドに失敗しました。" -ForegroundColor Red
    exit 1
}

# 単一実行ファイルとして発行
Write-Host "`n単一実行ファイルを作成中..." -ForegroundColor Yellow
dotnet publish src/RemoteWakeConnect.csproj `
    -c $buildConfig `
    -r win-x64 `
    --self-contained false `
    -p:PublishSingleFile=true `
    -p:EnableCompressionInSingleFile=false `
    -o publish

if ($LASTEXITCODE -eq 0) {
    Write-Host "`nビルドが正常に完了しました！" -ForegroundColor Green
    Write-Host "出力先: $projectDir\publish\RemoteWakeConnect.exe" -ForegroundColor Green
    
    # ファイルサイズを表示
    $exePath = Join-Path $projectDir "publish\RemoteWakeConnect.exe"
    if (Test-Path $exePath) {
        $fileInfo = Get-Item $exePath
        $sizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
        Write-Host "ファイルサイズ: $sizeMB MB" -ForegroundColor Cyan
    }
} else {
    Write-Host "`n発行に失敗しました。" -ForegroundColor Red
    exit 1
}

Write-Host "`n注意: 実行には.NET 8ランタイムが必要です。" -ForegroundColor Yellow
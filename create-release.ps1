# RemoteWakeConnect リリースパッケージ作成スクリプト
# Usage: .\create-release.ps1 [-Version "1.1.0"]

param(
    [string]$Version = ""  # 空の場合はプロジェクトファイルから自動取得
)

# バージョン自動取得
if ([string]::IsNullOrEmpty($Version)) {
    Write-Host "Detecting version from project files..." -ForegroundColor Cyan
    
    # プロジェクトファイルからバージョン取得
    $csprojPath = Join-Path $PSScriptRoot "src\RemoteWakeConnect.csproj"
    if (Test-Path $csprojPath) {
        $csprojContent = Get-Content $csprojPath -Raw
        if ($csprojContent -match '<AssemblyVersion>([^<]+)</AssemblyVersion>') {
            $Version = $matches[1]
            Write-Host "✓ Version detected from csproj: $Version" -ForegroundColor Green
        }
    }
    
    if ([string]::IsNullOrEmpty($Version)) {
        Write-Host "✗ Could not detect version automatically" -ForegroundColor Red
        Write-Host "  Please specify version: .\create-release.ps1 -Version '1.2.1'" -ForegroundColor Yellow
        exit 1
    }
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " RemoteWakeConnect Release Builder" -ForegroundColor Cyan
Write-Host " Version: $Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# パス設定
$projectPath = $PSScriptRoot
$publishPath = Join-Path $projectPath "publish"
$outputPath = Join-Path $projectPath "releases"
$zipFileName = "RemoteWakeConnect-v$Version.zip"
$zipFilePath = Join-Path $outputPath $zipFileName

# リリースフォルダがなければ作成
if (!(Test-Path $outputPath)) {
    New-Item -ItemType Directory -Path $outputPath | Out-Null
    Write-Host "✓ Created releases folder" -ForegroundColor Green
}

# 既存のzipファイルがあれば削除
if (Test-Path $zipFilePath) {
    Remove-Item $zipFilePath -Force
    Write-Host "✓ Removed existing zip file" -ForegroundColor Yellow
}

# ビルド実行
Write-Host ""
Write-Host "Building Release version..." -ForegroundColor Cyan
$buildResult = & dotnet publish -c Release -p:PublishSingleFile=true -p:EnableCompressionInSingleFile=false 2>&1

if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ Build failed!" -ForegroundColor Red
    Write-Host $buildResult
    exit 1
}

Write-Host "✓ Build completed successfully" -ForegroundColor Green

# publishフォルダの確認
if (!(Test-Path $publishPath)) {
    Write-Host "✗ Publish folder not found at: $publishPath" -ForegroundColor Red
    exit 1
}

# 一時フォルダ作成
$tempPath = Join-Path $projectPath "temp_release"
if (Test-Path $tempPath) {
    Remove-Item $tempPath -Recurse -Force
}
New-Item -ItemType Directory -Path $tempPath | Out-Null

Write-Host ""
Write-Host "Preparing release files..." -ForegroundColor Cyan

# 除外するファイル/フォルダのパターン
$excludePatterns = @(
    "rdp_files",
    "debug.log",
    "*.pdb",
    "*.log",
    "connection_history.yaml"
)

# ファイルをコピー（除外パターンを適用）
$copiedCount = 0
$excludedCount = 0

Get-ChildItem -Path $publishPath -Recurse | ForEach-Object {
    $relativePath = $_.FullName.Substring($publishPath.Length + 1)
    $shouldExclude = $false
    
    # rdp_filesフォルダ内のファイルは除外
    if ($relativePath -like "rdp_files\*" -and !$_.PSIsContainer) {
        $shouldExclude = $true
        $excludedCount++
        Write-Host "  - Excluded: $relativePath" -ForegroundColor DarkGray
    }
    # その他の除外パターンチェック
    elseif (!$_.PSIsContainer) {
        foreach ($pattern in $excludePatterns) {
            if ($_.Name -like $pattern) {
                $shouldExclude = $true
                $excludedCount++
                Write-Host "  - Excluded: $relativePath" -ForegroundColor DarkGray
                break
            }
        }
    }
    
    if (!$shouldExclude) {
        $targetPath = Join-Path $tempPath $relativePath
        $targetDir = Split-Path $targetPath -Parent
        
        if (!(Test-Path $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
        }
        
        if (!$_.PSIsContainer) {
            Copy-Item $_.FullName -Destination $targetPath -Force
            $copiedCount++
            Write-Host "  + Added: $relativePath" -ForegroundColor Green
        }
    }
}

# rdp_filesフォルダを空で作成
$rdpFilesPath = Join-Path $tempPath "rdp_files"
if (!(Test-Path $rdpFilesPath)) {
    New-Item -ItemType Directory -Path $rdpFilesPath -Force | Out-Null
    Write-Host "  + Created empty folder: rdp_files" -ForegroundColor Green
}

Write-Host ""
Write-Host "Files included: $copiedCount" -ForegroundColor Green
Write-Host "Files excluded: $excludedCount" -ForegroundColor Yellow

# README.txt作成
$readmeContent = @"
RemoteWakeConnect v$Version
========================

【概要】
Wake On LANとリモートデスクトップ接続を統合した便利ツール

【必要環境】
- Windows 10/11
- .NET 8.0 Runtime

【使い方】
1. RemoteWakeConnect.exe を実行
2. RDPファイルをドラッグ&ドロップまたは選択
3. 必要に応じてWake On LAN設定を入力
4. 「接続」ボタンでリモート接続開始

【主な機能】
- Wake On LANでPCを起動
- マルチモニター対応
- 接続履歴管理
- セッション監視機能（v1.1.0）

【更新履歴】
v$Version - $(Get-Date -Format "yyyy-MM-dd")
- セッション監視機能の高速化
- 段階的リトライ機能実装
- OS情報のキャッシュ機能追加

詳細は https://github.com/aziproducer/RemoteWakeConnect を参照
"@

$readmeContent | Out-File -FilePath (Join-Path $tempPath "README.txt") -Encoding UTF8
Write-Host "✓ Created README.txt" -ForegroundColor Green

# ZIP作成
Write-Host ""
Write-Host "Creating ZIP archive..." -ForegroundColor Cyan

try {
    # .NET Framework のZipFileクラスを使用
    Add-Type -Assembly System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempPath, $zipFilePath, 
        [System.IO.Compression.CompressionLevel]::Optimal, $false)
    
    Write-Host "✓ ZIP file created: $zipFileName" -ForegroundColor Green
    
    # ファイルサイズ表示
    $zipInfo = Get-Item $zipFilePath
    $sizeMB = [math]::Round($zipInfo.Length / 1MB, 2)
    Write-Host "  Size: $sizeMB MB" -ForegroundColor Cyan
}
catch {
    Write-Host "✗ Failed to create ZIP file: $_" -ForegroundColor Red
    exit 1
}

# 一時フォルダ削除
Remove-Item $tempPath -Recurse -Force
Write-Host "✓ Cleaned up temporary files" -ForegroundColor Green

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host " Release package created successfully!" -ForegroundColor Green
Write-Host " Location: $zipFilePath" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Yellow
Write-Host "1. Test the package by extracting and running" -ForegroundColor White
Write-Host "2. Upload to GitHub Releases with:" -ForegroundColor White
Write-Host "   gh release create v$Version $zipFilePath --title 'Version $Version' --notes 'Release notes here'" -ForegroundColor DarkGray
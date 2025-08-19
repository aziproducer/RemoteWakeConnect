# GitHub Release 公開スクリプト
# Usage: .\publish-github-release.ps1 [-Version "1.1.0"] [-Draft]

param(
    [string]$Version = "",  # 空の場合はプロジェクトファイルから自動取得
    [switch]$Draft = $false,
    [switch]$AutoNotes = $true  # CHANGELOGから自動生成
)

Write-Host "========================================" -ForegroundColor Cyan
Write-Host " GitHub Release Publisher" -ForegroundColor Cyan
Write-Host " Version: $Version" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

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
    
    # CHANGELOGからも確認
    $changelogPath = Join-Path $PSScriptRoot "CHANGELOG.md"
    if (Test-Path $changelogPath) {
        $changelogContent = Get-Content $changelogPath -Raw
        if ($changelogContent -match '## \[([^\]]+)\] - \d{4}-\d{2}-\d{2}') {
            $changelogVersion = $matches[1]
            if ($changelogVersion -ne $Version) {
                Write-Host "⚠ Version mismatch: csproj($Version) vs changelog($changelogVersion)" -ForegroundColor Yellow
                Write-Host "Using changelog version: $changelogVersion" -ForegroundColor Yellow
                $Version = $changelogVersion
            }
        }
    }
    
    if ([string]::IsNullOrEmpty($Version)) {
        Write-Host "✗ Could not detect version automatically" -ForegroundColor Red
        Write-Host "  Please specify version: .\publish-github-release.ps1 -Version '1.2.1'" -ForegroundColor Yellow
        exit 1
    }
}

Write-Host " Version: $Version" -ForegroundColor Cyan

# パス設定
$projectPath = $PSScriptRoot
$releasesPath = Join-Path $projectPath "releases"
$zipFileName = "RemoteWakeConnect-v$Version.zip"
$zipFilePath = Join-Path $releasesPath $zipFileName

# ZIPファイルの確認
if (!(Test-Path $zipFilePath)) {
    Write-Host "✗ Release package not found: $zipFileName" -ForegroundColor Red
    Write-Host "  Please run .\create-release.ps1 first" -ForegroundColor Yellow
    exit 1
}

Write-Host "✓ Found release package: $zipFileName" -ForegroundColor Green
$zipInfo = Get-Item $zipFilePath
$sizeMB = [math]::Round($zipInfo.Length / 1MB, 2)
Write-Host "  Size: $sizeMB MB" -ForegroundColor Cyan

# リリースノート生成
if ($AutoNotes -and (Test-Path (Join-Path $PSScriptRoot "CHANGELOG.md"))) {
    Write-Host "Generating release notes from CHANGELOG.md..." -ForegroundColor Cyan
    
    $changelogContent = Get-Content (Join-Path $PSScriptRoot "CHANGELOG.md") -Raw
    
    # 該当バージョンの部分を抽出
    $versionPattern = "## \[$Version\] - (\d{4}-\d{2}-\d{2})(.*?)(?=## \[|\z)"
    if ($changelogContent -match $versionPattern) {
        $releaseDate = $matches[1]
        $versionChanges = $matches[2].Trim()
        
        $releaseNotes = @"
# RemoteWakeConnect v$Version

Released: $releaseDate

$versionChanges

## 📦 インストール方法

1. ``RemoteWakeConnect-v$Version.zip`` をダウンロード
2. 任意のフォルダに解凍
3. ``RemoteWakeConnect.exe`` を実行

## 🔧 動作環境

- Windows 10/11
- .NET 8.0 Runtime

## 📝 更新履歴

詳細な変更点は [CHANGELOG.md](https://github.com/aziproducer/RemoteWakeConnect/blob/master/CHANGELOG.md) をご覧ください。

## 🐛 バグ報告・機能要望

[Issues](https://github.com/aziproducer/RemoteWakeConnect/issues) にて受け付けています。

---
*This release was created with automated build tools.*
"@
        Write-Host "✓ Release notes generated from CHANGELOG.md" -ForegroundColor Green
    } else {
        Write-Host "⚠ Version $Version not found in CHANGELOG.md, using default template" -ForegroundColor Yellow
        $releaseNotes = @"
# RemoteWakeConnect v$Version

## 📦 インストール方法

1. ``RemoteWakeConnect-v$Version.zip`` をダウンロード
2. 任意のフォルダに解凍
3. ``RemoteWakeConnect.exe`` を実行

## 🔧 動作環境

- Windows 10/11
- .NET 8.0 Runtime

## 📝 更新履歴

詳細な変更点は [CHANGELOG.md](https://github.com/aziproducer/RemoteWakeConnect/blob/master/CHANGELOG.md) をご覧ください。

---
*This release was created with automated build tools.*
"@
    }
} else {
    # 手動指定の場合のデフォルトテンプレート
    $releaseNotes = @"
# RemoteWakeConnect v$Version

## 📦 インストール方法

1. ``RemoteWakeConnect-v$Version.zip`` をダウンロード
2. 任意のフォルダに解凍
3. ``RemoteWakeConnect.exe`` を実行

## 🔧 動作環境

- Windows 10/11
- .NET 8.0 Runtime

詳細な変更点は [CHANGELOG.md](https://github.com/aziproducer/RemoteWakeConnect/blob/master/CHANGELOG.md) をご覧ください。
"@
}

# 一時ファイルにリリースノートを保存
$releaseNotesPath = Join-Path $env:TEMP "release_notes_$Version.md"
$releaseNotes | Out-File -FilePath $releaseNotesPath -Encoding UTF8

Write-Host ""
Write-Host "Checking GitHub CLI..." -ForegroundColor Cyan

# GitHub CLIの確認
$ghVersion = & gh --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ GitHub CLI is not installed" -ForegroundColor Red
    Write-Host "  Please install from: https://cli.github.com/" -ForegroundColor Yellow
    exit 1
}
Write-Host "✓ GitHub CLI is available" -ForegroundColor Green

# 認証状態確認
Write-Host "Checking authentication..." -ForegroundColor Cyan
$authStatus = & gh auth status 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "✗ Not authenticated with GitHub" -ForegroundColor Red
    Write-Host "  Please run: gh auth login" -ForegroundColor Yellow
    exit 1
}
Write-Host "✓ Authenticated with GitHub" -ForegroundColor Green

# 既存のリリースをチェック
Write-Host ""
Write-Host "Checking existing releases..." -ForegroundColor Cyan
$existingRelease = & gh release view "v$Version" --json tagName 2>&1
if ($LASTEXITCODE -eq 0) {
    Write-Host "⚠ Release v$Version already exists" -ForegroundColor Yellow
    $response = Read-Host "Do you want to delete and recreate it? (y/n)"
    if ($response -eq 'y') {
        Write-Host "Deleting existing release..." -ForegroundColor Yellow
        & gh release delete "v$Version" -y
        Write-Host "✓ Deleted existing release" -ForegroundColor Green
    } else {
        Write-Host "Release creation cancelled" -ForegroundColor Yellow
        exit 0
    }
}

# リリース作成
Write-Host ""
Write-Host "Creating GitHub Release..." -ForegroundColor Cyan

$draftFlag = if ($Draft) { "--draft" } else { "" }

try {
    if ($draftFlag) {
        & gh release create "v$Version" $zipFilePath `
            --title "RemoteWakeConnect v$Version" `
            --notes-file $releaseNotesPath `
            --draft
    } else {
        & gh release create "v$Version" $zipFilePath `
            --title "RemoteWakeConnect v$Version" `
            --notes-file $releaseNotesPath
    }
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "========================================" -ForegroundColor Cyan
        Write-Host " ✓ Release published successfully!" -ForegroundColor Green
        Write-Host "========================================" -ForegroundColor Cyan
        
        if ($Draft) {
            Write-Host "Status: DRAFT" -ForegroundColor Yellow
            Write-Host "URL: https://github.com/aziproducer/RemoteWakeConnect/releases/tag/v$Version" -ForegroundColor Cyan
            Write-Host ""
            Write-Host "To publish the draft:" -ForegroundColor Yellow
            Write-Host "  gh release edit v$Version --draft=false" -ForegroundColor White
        } else {
            Write-Host "Status: PUBLISHED" -ForegroundColor Green
            Write-Host "URL: https://github.com/aziproducer/RemoteWakeConnect/releases/tag/v$Version" -ForegroundColor Cyan
        }
    } else {
        throw "Failed to create release"
    }
}
catch {
    Write-Host "✗ Failed to create release: $_" -ForegroundColor Red
    exit 1
}
finally {
    # 一時ファイル削除
    if (Test-Path $releaseNotesPath) {
        Remove-Item $releaseNotesPath -Force
    }
}
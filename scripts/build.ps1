# 本地构建：默认先跑测试，再打单文件 exe
# 用法: powershell -ExecutionPolicy Bypass -File scripts/build.ps1 [-SkipTests]
param([switch]$SkipTests)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not $SkipTests) {
    & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'test.ps1')
    if ($LASTEXITCODE -ne 0) { throw '测试未通过，已中止构建' }
}

python -m PyInstaller --onefile --windowed --name OfficeAutomationTool --noconfirm merge_gui.py
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller 构建失败（是否已安装 requirements-dev.txt？）' }

$exe = Join-Path $root 'dist\OfficeAutomationTool.exe'
Write-Host ''
Write-Host "构建完成: $exe" -ForegroundColor Green
Write-Host ("大小: {0:N1} MB" -f ((Get-Item $exe).Length / 1MB))

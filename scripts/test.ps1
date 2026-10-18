# 跑全部测试：任何一个失败就以非零退出
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$tests = @('tests/test_logic.py', 'tests/test_gui.py', 'tests/test_e2e.py')
$failed = 0

foreach ($t in $tests) {
    Write-Host "=== $t" -ForegroundColor Cyan
    & python $t
    if ($LASTEXITCODE -ne 0) { $failed++ }
}

if ($failed -gt 0) {
    Write-Host "$failed 个测试文件失败" -ForegroundColor Red
    exit 1
}
Write-Host '全部测试通过' -ForegroundColor Green
exit 0

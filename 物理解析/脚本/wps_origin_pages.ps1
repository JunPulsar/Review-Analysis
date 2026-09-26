# 用法（二选一）：
#   1) cd 到项目文件夹（内含 _toc_tmp、titles 清单等）后直接运行：
#        .\wps_origin_pages.ps1
#   2) 显式指定项目路径：
#        .\wps_origin_pages.ps1 -root "D:\你的项目路径"
param([string]$root = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$out = Join-Path $root 'wps_origin_pages.txt'

$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
$app.DisplayAlerts = 0

$total = 0
$lines = @()
Get-ChildItem -Path $root -Filter *.docx | Where-Object {
    $_.Name -ne '教师目录.docx' -and
    $_.Name -notlike '物理解析*' -and
    $_.Name -notlike '*_新.docx'
} | ForEach-Object {
    $d = $app.Documents.Open($_.FullName, $false, $true)
    $p = $d.ComputeStatistics(2)
    $d.Close($false)
    $total += $p
    $lines += ('{0}|{1}' -f $_.Name, $p)
}
$app.Quit()

$lines | Out-File $out -Encoding UTF8
Add-Content -Path $out -Value ('TOTAL|' + $total) -Encoding UTF8
Write-Output ('原始总页数: ' + $total)

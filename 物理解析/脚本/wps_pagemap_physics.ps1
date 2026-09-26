# 用法（二选一）：
#   1) cd 到项目文件夹（内含 _toc_tmp、titles 清单等）后直接运行：
#        .\wps_pagemap_physics.ps1
#   2) 显式指定项目路径：
#        .\wps_pagemap_physics.ps1 -root "D:\你的项目路径"
param([string]$root = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$tmp = Join-Path $root '_toc_tmp'

# 三本：大本 / 小本 / 合订本
$items = @(
    @('物理解析_大本.docx',        'titles_book.txt', 'pagemap_book.txt'),
    @('物理解析_小本.docx',        'titles_xb.txt',   'pagemap_xb.txt'),
    @('物理解析_大本小本合订本.docx', 'titles_cb.txt',   'pagemap_cb.txt')
)

$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
$app.DisplayAlerts = 0

foreach ($it in $items) {
    $docx = Join-Path $root $it[0]
    $titlesPath = Join-Path $tmp $it[1]
    $outPath = Join-Path $tmp $it[2]
    if (-not (Test-Path $docx)) { Write-Output ('skip(缺文件): ' + $it[0]); continue }
    if (-not (Test-Path $titlesPath)) { Write-Output ('skip(缺清单): ' + $it[1]); continue }

    $titles = @{}
    Get-Content $titlesPath -Encoding UTF8 | ForEach-Object {
        if ($_ -ne '') { $titles[$_] = -1 }
    }

    $doc = $app.Documents.Open($docx, $false, $true)
    foreach ($para in $doc.Paragraphs) {
        $t = $para.Range.Text
        if ($t.Length -ge 2) {
            $t = $t.Substring(0, $t.Length - 1)   # 去掉段落标记
            if ($titles.ContainsKey($t) -and $titles[$t] -eq -1) {
                $r = $para.Range
                $r.Collapse(1)
                $titles[$t] = $r.Information(3)   # wdActiveEndPageNumber
            }
        }
    }
    $doc.Close($false)

    $lines = @()
    foreach ($key in $titles.Keys) { $lines += ('{0}|{1}' -f $key, $titles[$key]) }
    $lines | Out-File $outPath -Encoding UTF8
    $miss = ($titles.Values | Where-Object { $_ -eq -1 }).Count
    Write-Output ('pagemap done: {0}  标题 {1}  未定位 {2}' -f $it[0], $titles.Count, $miss)
}

$app.Quit()
Write-Output 'ALL DONE'

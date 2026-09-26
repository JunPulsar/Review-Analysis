# 用法（二选一）：
#   1) cd 到项目文件夹（内含 _toc_tmp、titles 清单等）后直接运行：
#        .\wps_pagemap.ps1
#   2) 显式指定项目路径：
#        .\wps_pagemap.ps1 -root "D:\你的项目路径"
param([string]$root = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$tmp = Join-Path $root '_toc_tmp'

$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
$app.DisplayAlerts = 0

function Get-PageMap($docxPath, $titlesPath, $outPath) {
    $titles = @{}
    Get-Content $titlesPath -Encoding UTF8 | ForEach-Object {
        if ($_ -ne '') { $titles[$_] = -1 }
    }
    $doc = $app.Documents.Open($docxPath, $false, $true)
    foreach ($para in $doc.Paragraphs) {
        $t = $para.Range.Text
        if ($t.Length -ge 2) {
            $t = $t.Substring(0, $t.Length - 1)   # 去掉段落标记
            if ($titles.ContainsKey($t)) {
                $r = $para.Range
                $r.Collapse(1)                     # 定位到段首
                $pg = $r.Information(3)            # wdActiveEndPageNumber
                $titles[$t] = $pg
            }
        }
    }
    $doc.Close($false)
    $lines = @()
    foreach ($key in $titles.Keys) {
        $lines += ('{0}|{1}' -f $key, $titles[$key])
    }
    $lines | Out-File $outPath -Encoding UTF8
    Write-Output ('map done: ' + (Split-Path $docxPath -Leaf))
}

$items = @(
    @('成册解析\大本解析_全书.docx', 'titles_book.txt', 'pagemap_book.txt'),
    @('成册解析\小本解析_合集.docx', 'titles_xb.txt', 'pagemap_xb.txt'),
    @('成册解析\大本小本合订本.docx', 'titles_cb.txt', 'pagemap_cb.txt')
)
foreach ($it in $items) {
    $docxPath = Join-Path $root $it[0]
    $titlesPath = Join-Path $tmp $it[1]
    $outPath = Join-Path $tmp $it[2]
    if ((Test-Path $docxPath) -and (Test-Path $titlesPath)) {
        Get-PageMap $docxPath $titlesPath $outPath
    } else {
        Write-Output ('skip: ' + $it[0] + '  (titles 未生成: ' + $titlesPath + ')')
    }
}

$app.Quit()
Write-Output 'ALL DONE'

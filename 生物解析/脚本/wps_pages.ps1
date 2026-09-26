# 用法（二选一）：
#   1) cd 到项目文件夹（内含 _toc_tmp、titles 清单等）后直接运行：
#        .\wps_pages.ps1
#   2) 显式指定项目路径：
#        .\wps_pages.ps1 -root "D:\你的项目路径"
param([string]$root = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$result = @()

$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
$app.DisplayAlerts = 0

function Get-Pages($path) {
    $doc = $app.Documents.Open($path, $false, $true)
    $p = $doc.ComputeStatistics(2)   # wdStatisticPages = 2
    $doc.Close($false)
    return $p
}

# ── 原始 97 篇 ──
$origTotal = 0
$origCount = 0
$origLines = @()
Get-ChildItem $root -Directory -Filter '大概念*' | ForEach-Object {
    $day = $_
    Get-ChildItem $day.FullName -Filter *.docx | Where-Object { -not $_.Name.StartsWith('~$') -and -not $_.Name.EndsWith('_解析.docx') } | ForEach-Object {
        try {
            $p = Get-Pages $_.FullName
            $origTotal += $p
            $origCount += 1
            $origLines += ('{0}|{1}|{2}' -f $day.Name, $_.Name, $p)
            Write-Output ('OK 原始 {0}/{1} | {2} 页' -f $origCount, $_.Name, $p)
        } catch {
            Write-Output ('FAIL {0} | {1}' -f $_.Name, $_.Exception.Message)
            $origLines += ('{0}|{1}|FAIL {2}' -f $day.Name, $_.Name, $_.Exception.Message)
        }
    }
}

# ── 新整合全书 + 9 分册 ──
$bookPages = -1
try { $bookPages = Get-Pages (Join-Path $root '新整合_全书\解析合集_全书.docx') } catch { Write-Output ('FAIL 全书 | ' + $_.Exception.Message) }
$partTotal = 0
$partLines = @()
Get-ChildItem (Join-Path $root '新整合_分册') -Filter *.docx | ForEach-Object {
    try {
        $p = Get-Pages $_.FullName
        $partTotal += $p
        $partLines += ('{0}|{1}' -f $_.Name, $p)
        Write-Output ('OK 分册 {0} | {1} 页' -f $_.Name, $p)
    } catch {
        Write-Output ('FAIL 分册 {0} | {1}' -f $_.Name, $_.Exception.Message)
        $partLines += ('{0}|FAIL {1}' -f $_.Name, $_.Exception.Message)
    }
}

$app.Quit()

$lines = @()
$lines += 'WPS 实际打印页数统计'
$lines += '=' * 60
$lines += ('原始 97 篇合计页数: {0}  ({1} 篇成功)' -f $origTotal, $origCount)
$lines += ('新整合 全书页数: {0}' -f $bookPages)
$lines += ('新整合 9 分册合计页数: {0}' -f $partTotal)
$lines += ('对比: 原始 {0} 页 -> 全书 {1} 页, 节省 {2} 页 ({3:N1}%)' -f $origTotal, $bookPages, ($origTotal - $bookPages), ((1 - ($bookPages / $origTotal)) * 100))
$lines += ''
$lines += '=== 原始每篇页数 ==='
$lines += $origLines
$lines += ''
$lines += '=== 分册页数 ==='
$lines += $partLines

$lines | Out-File (Join-Path $root 'wps_pages_result.txt') -Encoding UTF8
Write-Output 'DONE'

# 用法（二选一）：
#   1) cd 到项目文件夹（内含 _toc_tmp、titles 清单等）后直接运行：
#        .\wps_zm.ps1
#   2) 显式指定项目路径：
#        .\wps_zm.ps1 -root "D:\你的项目路径"
param([string]$root = (Get-Location).Path)

$ErrorActionPreference = 'Stop'
$app = New-Object -ComObject KWPS.Application
$app.Visible = $false
$app.DisplayAlerts = 0
function Get-Pages($path) {
    $doc = $app.Documents.Open($path, $false, $true)
    $p = $doc.ComputeStatistics(2)
    $doc.Close($false)
    return $p
}
$zm = Get-Pages (Join-Path $root '新整合_周末必刷\周末必刷解析合集.docx')
$book = Get-Pages (Join-Path $root '新整合_全书\解析合集_全书.docx')
$app.Quit()
$lines = @()
$lines += ('orig_all_97: 1288')
$lines += ('book: {0}' -f $book)
$lines += ('zm: {0}' -f $zm)
$total = $book + $zm
$lines += ('total: {0}' -f $total)
$lines += ('saved: {0} ({1:N1}%)' -f (1288 - $total), ((1 - ($total / 1288)) * 100))
$lines | Out-File (Join-Path $root 'wps_final3.txt') -Encoding UTF8
Write-Output 'DONE'

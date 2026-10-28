$ErrorActionPreference = 'Stop'
$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open('D:\text\02_中世纪欧洲文明\讨论课\交付文件\查理曼_12分钟汇报讲稿.docx', $false, $true)
    $document.ExportAsFixedFormat('D:\text\02_中世纪欧洲文明\讨论课\.制作过程\讲稿预览\讲稿.pdf', 17)
    Write-Output ('Rendered pages: ' + $document.ComputeStatistics(2))
} finally {
    if ($null -ne $document) { $document.Close(0) }
    if ($null -ne $word) { $word.Quit() }
}

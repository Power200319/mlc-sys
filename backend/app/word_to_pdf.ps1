param([Parameter(Mandatory = $true)][string]$InputDirectory)

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    foreach ($file in Get-ChildItem -LiteralPath $InputDirectory -Filter *.docx) {
        $document = $word.Documents.Open($file.FullName, $false, $true)
        $pdfPath = [System.IO.Path]::ChangeExtension($file.FullName, '.pdf')
        $document.ExportAsFixedFormat($pdfPath, 17)
        $document.Close(0)
        $document = $null
    }
}
finally {
    if ($document) { $document.Close(0) }
    if ($word) { $word.Quit() }
}

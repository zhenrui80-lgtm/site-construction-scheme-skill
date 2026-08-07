param(
    [Parameter(Mandatory = $true)]
    [string]$InputDocx,

    [Parameter(Mandatory = $true)]
    [string]$OutputPdf
)

$inputPath = [System.IO.Path]::GetFullPath($InputDocx)
$outputPath = [System.IO.Path]::GetFullPath($OutputPdf)
if (-not (Test-Path -LiteralPath $inputPath)) {
    throw "DOCX not found: $inputPath"
}

$outputDirectory = Split-Path -Parent $outputPath
if ($outputDirectory) {
    New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null
}

$word = $null
$document = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($inputPath, $false, $false)
    foreach ($contents in $document.TablesOfContents) {
        $contents.Update()
    }
    $document.Fields.Update() | Out-Null
    $document.Repaginate()
    $pages = $document.ComputeStatistics(2)
    $document.Save()
    $document.ExportAsFixedFormat($outputPath, 17)
    [pscustomobject]@{
        docx = $inputPath
        pdf = $outputPath
        pages = $pages
    } | ConvertTo-Json -Compress
}
finally {
    if ($document) {
        try { $document.Close($false) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($word) {
        try { $word.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

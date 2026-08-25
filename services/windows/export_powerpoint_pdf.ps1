param(
    [Parameter(Mandatory = $true)][string]$InputPath,
    [Parameter(Mandatory = $true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'
$ppSaveAsPDF = 32
$msoTrue = -1
$msoFalse = 0
$inputFull = [System.IO.Path]::GetFullPath($InputPath)
$outputFull = [System.IO.Path]::GetFullPath($OutputPath)
New-Item -ItemType Directory -Force -Path (Split-Path $outputFull) | Out-Null

$wasAlreadyRunning = $null -ne (Get-Process POWERPNT -ErrorAction SilentlyContinue)
$ppt = $null
$pres = $null
try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $msoTrue
    $pres = $ppt.Presentations.Open($inputFull, $msoFalse, $msoFalse, $msoFalse)
    $pres.SaveAs($outputFull, $ppSaveAsPDF)
    $pres.Close()
    if (-not $wasAlreadyRunning) { $ppt.Quit() }
}
finally {
    if ($pres) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null }
    if ($ppt) { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

if (-not (Test-Path $outputFull)) {
    throw "PowerPoint did not create the PDF preview."
}

param(
    [Parameter(Mandatory=$true)][string]$XlsxPath,
    [Parameter(Mandatory=$true)][string]$PdfPath
)

$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($XlsxPath)
    $xl.CalculateFull()
    if (Test-Path $PdfPath) { Remove-Item $PdfPath -Force }
    # 0 = xlTypePDF, 0 = xlQualityStandard
    $wb.ExportAsFixedFormat(0, $PdfPath, 0, $true, $false)
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Write-Host "Exported PDF -> $PdfPath"

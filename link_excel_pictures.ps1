param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath
)

$ErrorActionPreference = "Stop"
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false

try {
    $wb = $xl.Workbooks.Open($WorkbookPath)
    $sheetMap = @{
        "2. Footing BBS (Lx & Ly)"        = "Pic_F_"
        "3. Column BBS (Main & Ties)"     = "Pic_C_"
        "4. Beam BBS (Top, Bot, Stirrup)" = "Pic_B_"
        "5. Slab BBS (Main, Dist, Crank)" = "Pic_S_"
    }

    foreach ($sname in $sheetMap.Keys) {
        $prefix = $sheetMap[$sname]
        $ws = $wb.Sheets.Item($sname)
        foreach ($shp in $ws.Shapes) {
            $row = $shp.TopLeftCell.Row
            if ($row -ge 5) {
                $fname = "=" + $prefix + $row
                $shp.DrawingObject.Formula = $fname
            }
        }
    }

    $wb.Save()
    $wb.Close($false)
}
finally {
    $xl.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
}
Write-Host "Linked all dynamic Shape Pictures in $WorkbookPath"

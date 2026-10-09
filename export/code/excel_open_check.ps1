# Opens a workbook in Excel through COM, read-only, and reports the sheets it found.
# Excel does not attempt recovery when a file is opened from the object model, so a workbook
# that would need repair fails here instead of being repaired silently.
# ASCII only: PowerShell 5.1 reads a BOM-less file as ANSI.
param([Parameter(Mandatory = $true)][string]$Path)
$ErrorActionPreference = 'Stop'
$excel = $null
$code = 0
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($Path, 0, $true)
    $found = @()
    foreach ($sheet in $workbook.Worksheets) {
        $found += ($sheet.Name + ' [' + $sheet.UsedRange.Address($false, $false) + ']')
    }
    Write-Output ('OPENED without repair: ' + ($found -join '; '))
    $workbook.Close($false)
} catch {
    Write-Output ('OPEN FAILED: ' + $_.Exception.Message)
    $code = 1
} finally {
    if ($excel -ne $null) {
        $excel.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
}
exit $code

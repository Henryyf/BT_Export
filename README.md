# BT_Export



# Save as: Save-ExcelSnapshot.ps1
param(
    [int]$Pid = 25308,       # <-- change if needed
    [int]$Iterations = 10,
    [int]$IntervalSec = 10
)

try {
    $proc = Get-Process -Id $Pid -ErrorAction Stop
} catch {
    throw "No process with PID $Pid was found."
}

$title = $proc.MainWindowTitle
if (-not $title) {
    throw "Excel PID $Pid has no main window title (no workbook open yet?)."
}

# Derive workbook name from window title like: 'MyBook - Excel'
$baseName = ($title -replace ' - Excel.*$','').Trim() -replace '\*$',''   # remove trailing '*' if unsaved

# Attach to the running Excel instance (first registered in ROT)
$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$null = $xl | Out-Null

# Try to find the workbook by common extensions (and raw name)
$extensions = @('xlsx','xlsm','xlsb','xls','csv','xlsm')  # a small set to try
$wb = $null
foreach ($ext in $extensions) {
    $candidate = if ($baseName -match "\.$ext$") { $baseName } else { "$baseName.$ext" }
    $wb = $xl.Workbooks | Where-Object { $_.Name -eq $candidate }
    if ($wb) { break }
}
if (-not $wb) {
    # Fallback: use ActiveWorkbook if we couldn't match by name
    $wb = $xl.ActiveWorkbook
    if (-not $wb) { throw "Could not locate a workbook for PID $Pid. Is a workbook open?" }
}

$sheet = $wb.ActiveSheet
if (-not $sheet) { throw "Active sheet not found in workbook '$($wb.Name)'." }

$desktop = [Environment]::GetFolderPath('Desktop')
$xl.DisplayAlerts = $false

for ($i = 1; $i -le $Iterations; $i++) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $desktop ("excel{0}.csv" -f $timestamp)

    # Copy values into a temporary workbook to avoid altering the source
    $src = $sheet.UsedRange
    $tempWb = $xl.Workbooks.Add()
    $dest = $tempWb.ActiveSheet.Range("A1").Resize($src.Rows.Count, $src.Columns.Count)
    $dest.Value2 = $src.Value2   # values only (no formulas, no formats)

    # Save as CSV (xlCSV = 6)
    $tempWb.SaveAs($csvPath, 6)
    $tempWb.Close($false)

    Write-Host "Saved $csvPath"
    if ($i -lt $Iterations) { Start-Sleep -Seconds $IntervalSec }
}


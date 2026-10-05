<#
Finishes the Excel report using Excel itself (Windows only):
  - adds two PivotTables on the tblWeekly history table
  - recalculates every formula and saves, so values are cached for any viewer
Usage: powershell -ExecutionPolicy Bypass -File scripts\excel_finalize.ps1 -Path outputs\pool_point_operations_report.xlsx
#>
param([Parameter(Mandatory = $true)][string]$Path)

$Path = (Resolve-Path $Path).Path
$missing = [System.Reflection.Missing]::Value
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

function Add-PivotSheet($wb, $after, $name, $title, $note) {
    $ws = $wb.Worksheets.Add($missing, $after)
    $ws.Name = $name
    $ws.Range("A1").Value2 = $title
    $ws.Range("A1").Font.Bold = $true
    $ws.Range("A1").Font.Size = 16
    $ws.Range("A2").Value2 = $note
    $ws.Range("A2").Font.Italic = $true
    $ws.Tab.Color = 0xD6782A   # BGR for #2A78D6
    return $ws
}

try {
    $wb = $excel.Workbooks.Open($Path)
    $history = $wb.Worksheets.Item("Weekly History")
    $cache = $wb.PivotCaches().Create(1, "tblWeekly")   # 1 = xlDatabase

    # Pivot 1: pallets shipped by pool point and year
    $ws1 = Add-PivotSheet $wb $history "Pivot Volume" "Pallets shipped by pool point and year" `
        "PivotTable on tblWeekly. Drag fields to slice by carrier or month."
    $pt1 = $cache.CreatePivotTable($ws1.Range("A4"), "ptVolumeByYear")
    $pt1.PivotFields("pool_point_id").Orientation = 1   # row
    $pt1.PivotFields("year").Orientation = 2            # column
    $f1 = $pt1.AddDataField($pt1.PivotFields("pallets_shipped"), "Pallets shipped ", -4157)  # sum
    $f1.NumberFormat = "#,##0"

    # Pivot 2: on-time rate by carrier and year, using a calculated field (weighted, not averaged)
    $ws2 = Add-PivotSheet $wb $ws1 "Pivot On-Time" "On-time delivery rate by carrier and year" `
        "Calculated field OnTimeRate = on_time_stops / stops, so each rate is weighted by deliveries."
    $pt2 = $cache.CreatePivotTable($ws2.Range("A4"), "ptOnTimeByCarrier")
    $pt2.PivotFields("carrier_id").Orientation = 1
    $pt2.PivotFields("year").Orientation = 2
    $calc = $pt2.CalculatedFields().Add("OnTimeRate", "=on_time_stops/stops", $true)
    $f2 = $pt2.AddDataField($calc, "On-time rate ", -4157)
    $f2.NumberFormat = "0.0%"

    foreach ($ws in @($ws1, $ws2)) { $ws.Columns("A:H").ColumnWidth = 16 }

    $wb.Worksheets.Item("Dashboard").Activate()
    $excel.CalculateFull()
    $wb.Save()
    $wb.Close($false)
    Write-Output "Finalized $Path"
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

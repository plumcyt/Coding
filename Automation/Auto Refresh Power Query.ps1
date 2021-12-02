$Excel = New-Object -COM "Excel.Application"
$Excel.Visible = $true
$Workbook = $Excel.Workbooks.Open("C:\Users\username\Desktop\excel.xlsx")

foreach ($sheet in $Workbook.Sheets) {
    $sheet.QueryTables | ForEach-Object {
        while ($_.QueryTable.Refreshing) {
            Start-Sleep -Seconds 1
        }
    }
}

$Excel.Save()
$Excel.Close()

# clear the COM object you have created after finishing with them to free memory:

$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook)
$null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel)
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

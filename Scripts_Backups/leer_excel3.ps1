$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastCol = $worksheet.Cells.SpecialCells(11).Column

for ($row = 1; $row -le 5; $row++) {
    $line = ""
    for ($col = 1; $col -le $lastCol; $col++) {
        $val = $worksheet.Cells.Item($row, $col).Value2
        $line += "$val | "
    }
    Write-Host $line
}
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

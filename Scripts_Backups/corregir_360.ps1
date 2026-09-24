$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.Cells.SpecialCells(11).Row

for ($row = 2; $row -le $lastRow; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $key = ([string]$ref).ToUpper().Trim()
    
    if ($key -eq "360") {
        $worksheet.Cells.Item($row, 6).Value2 = "Floral"
        $worksheet.Cells.Item($row, 7).Value2 = "melón, lirio, osmanto, mandarina, rosa, nenúfar, lirio de los valles, lavanda, salvia, almizcle, sándalo, vetiver, ámbar y vainilla."
        $worksheet.Cells.Item($row, 9).Value2 = "Dama"
    }
}
$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "360 corregido."

$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)

for ($row = 1; $row -le 5; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    $top = $worksheet.Cells.Item($row, 6).Value2
    Write-Host "$ref | $top"
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

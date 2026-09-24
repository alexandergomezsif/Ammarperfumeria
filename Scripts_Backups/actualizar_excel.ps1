$fotos_dir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"

if (-not (Test-Path $xlsx_path)) {
    Write-Host "ERROR: No se encuentra Base_Datos_Perfumes.xlsx"
    exit
}

$currentFiles = @{}
$brands = Get-ChildItem -Path $fotos_dir -Directory
foreach ($brand in $brands) {
    $files = Get-ChildItem -Path $brand.FullName -File -Filter "*.jpg"
    foreach ($file in $files) {
        $ref = $file.Name -replace '\.jpg$', '' -replace '\.JPG$', '' -replace '_', ' '
        $currentFiles[$file.Name] = @{
            Marca = $brand.Name
            Referencia = $ref
        }
    }
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)

$lastRow = $worksheet.Cells.SpecialCells(11).Row

$colFoto = 7
$colMarca = 1
$colRef = 2

$processedFiles = @()

for ($row = 2; $row -le $lastRow; $row++) {
    $fotoName = $worksheet.Cells.Item($row, $colFoto).Value2
    
    if ($fotoName -ne $null -and $currentFiles.ContainsKey($fotoName)) {
        $worksheet.Cells.Item($row, $colMarca).Value2 = $currentFiles[$fotoName].Marca
        $worksheet.Cells.Item($row, $colRef).Value2 = $currentFiles[$fotoName].Referencia
        $processedFiles += $fotoName
    }
}

foreach ($fotoName in $currentFiles.Keys) {
    if ($processedFiles -notcontains $fotoName) {
        $lastRow++
        $worksheet.Cells.Item($lastRow, $colMarca).Value2 = $currentFiles[$fotoName].Marca
        $worksheet.Cells.Item($lastRow, $colRef).Value2 = $currentFiles[$fotoName].Referencia
        $worksheet.Cells.Item($lastRow, $colFoto).Value2 = $fotoName
    }
}

$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel actualizado correctamente."

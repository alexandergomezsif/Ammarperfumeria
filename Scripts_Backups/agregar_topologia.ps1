$txtPath = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Skill_Diseño_Perfumero\7. Topologia y Arquitectura.txt"
$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.Cells.SpecialCells(11).Row

# Helper to normalize names
function Normalize-Name($name) {
    $n = $name -replace " \(.*?\)", ""
    $n = $n -replace "[\.-]", ""
    $n = $n -replace " ", ""
    return $n.ToUpper()
}

$content = Get-Content -Path $txtPath -Raw
$blocks = $content -split "(?m)^(?=[A-Za-z0-9])"

$dict = @{}
foreach ($block in $blocks) {
    if ($block -match "^([^\r\n]+)\r?\n+Clasificación:\s*([^\r\n]+)\r?\n+Detalle:\s*([^\r\n]+)") {
        $name = $matches[1].Trim()
        $clas = $matches[2].Trim()
        $det = $matches[3].Trim()
        $norm = Normalize-Name $name
        $dict[$norm] = "Clasificación: $clas`nDetalle: $det"
    }
}

for ($row = 2; $row -le $lastRow; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $normRef = Normalize-Name $ref
    if ($dict.ContainsKey($normRef)) {
        $worksheet.Cells.Item($row, 6).Value2 = $dict[$normRef]
    } else {
        # Try partial match or manual fixes
        if ($normRef -eq "CLOUD20INTENSE") { $worksheet.Cells.Item($row, 6).Value2 = $dict["CLOUD20INTENSE"] }
        if ($normRef -eq "CLUBDENUITINTENSE") { $worksheet.Cells.Item($row, 6).Value2 = $dict["CLUBDENUITINTENSE"] } # No exact match in txt for intense? Wait, let's see.
        
        # We will do a substring search if direct fails
        $found = $false
        foreach ($key in $dict.Keys) {
            if ($key -eq $normRef) {
                $worksheet.Cells.Item($row, 6).Value2 = $dict[$key]
                $found = $true
                break
            }
        }
    }
}

$worksheet.Columns.Item(6).ColumnWidth = 60
$worksheet.Columns.Item(6).WrapText = $true

$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel actualizado con la columna Topologia."

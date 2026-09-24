$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$fotosDir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$outJs = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\js\perfumes.js"

$images = Get-ChildItem -Path $fotosDir -Recurse -File -Filter "*.jpg"
$imgDict = @{}
foreach ($img in $images) {
    $refName = ($img.BaseName -replace "_", " ").ToUpper().Trim()
    $relPath = "fotos/" + $img.Directory.Name + "/" + $img.Name
    $imgDict[$refName] = $relPath
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.UsedRange.Rows.Count

$perfumes = @()
$lastMarca = "Desconocido"

for ($row = 2; $row -le $lastRow; $row++) {
    $marca = $worksheet.Cells.Item($row, 1).Value2
    if ($marca) { $lastMarca = $marca } else { $marca = $lastMarca }
    
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $gen = $worksheet.Cells.Item($row, 3).Value2
    $fam = $worksheet.Cells.Item($row, 4).Value2
    $not = $worksheet.Cells.Item($row, 5).Value2
    $top = $worksheet.Cells.Item($row, 6).Value2
    $tips = $worksheet.Cells.Item($row, 7).Value2
    
    $refKey = ([string]$ref).ToUpper().Trim()
    $imgPath = "https://via.placeholder.com/300x400?text=Falta+Foto"
    if ($imgDict.ContainsKey($refKey)) {
        $imgPath = $imgDict[$refKey]
    } else {
        foreach ($k in $imgDict.Keys) {
            if ($k -match [regex]::Escape($refKey)) {
                $imgPath = $imgDict[$k]
                break
            }
        }
    }
    
    # Random price
    $price = (Get-Random -Minimum 160 -Maximum 251) * 1000
    
    $perfumes += @{
        Marca = [string]$marca
        Referencia = [string]$ref
        Genero = [string]$gen
        Familia = [string]$fam
        Notas = [string]$not
        Topologia = [string]$top
        Tips = [string]$tips
        Imagen = [string]$imgPath
        Precio = $price
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$jsonStr = $perfumes | ConvertTo-Json -Depth 5 -Compress
$jsContent = "const perfumesData = $jsonStr;"

# Guardar como UTF-8 sin BOM para máxima compatibilidad web
$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllText($outJs, $jsContent, $utf8NoBom)
Write-Host "perfumes.js generado con éxito."

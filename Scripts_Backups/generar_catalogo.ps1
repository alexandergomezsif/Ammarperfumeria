$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$fotosDir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$outHtml = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Catalogo_Ammar.html"

$images = Get-ChildItem -Path $fotosDir -Recurse -File -Filter "*.jpg"
$imgDict = @{}
foreach ($img in $images) {
    $refName = ($img.BaseName -replace "_", " ").ToUpper().Trim()
    $imgPath = $img.FullName -replace "\\", "/"
    $imgDict[$refName] = "file:///$imgPath"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.Cells.SpecialCells(11).Row

$perfumes = @()

for ($row = 2; $row -le $lastRow; $row++) {
    $marca = $worksheet.Cells.Item($row, 1).Value2
    if (-not $marca) { $marca = $lastMarca } else { $lastMarca = $marca }
    
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $gen = $worksheet.Cells.Item($row, 3).Value2
    $fam = $worksheet.Cells.Item($row, 4).Value2
    $not = $worksheet.Cells.Item($row, 5).Value2
    $top = $worksheet.Cells.Item($row, 6).Value2
    $tips = $worksheet.Cells.Item($row, 7).Value2
    
    $refKey = ([string]$ref).ToUpper().Trim()
    $imgPath = ""
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
    
    $perfumes += [PSCustomObject]@{
        Marca = $marca
        Referencia = $ref
        Genero = $gen
        Familia = $fam
        Notas = $not
        Topologia = $top
        Tips = $tips
        Imagen = $imgPath
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<title>Catálogo AMMAR Perfumería</title>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600&family=Playfair+Display:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet">
<style>
    :root {
        --bg: #ffffff;
        --text-main: #111111;
        --text-muted: #555555;
        --border: #e0e0e0;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Montserrat', sans-serif;
        background-color: #f9f9f9;
        color: var(--text-main);
        line-height: 1.4;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
    .page {
        width: 210mm;
        height: 297mm;
        margin: 10mm auto;
        padding: 15mm;
        background: white;
        box-shadow: 0 0 10px rgba(0,0,0,0.1);
        page-break-after: always;
        position: relative;
    }
    .cover {
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        text-align: center;
        height: 100%;
        border: 2px solid var(--text-main);
        padding: 20px;
    }
    .cover h3 {
        font-size: 1rem;
        font-weight: 300;
        letter-spacing: 5px;
        margin-bottom: 20px;
        text-transform: uppercase;
    }
    .cover h1 {
        font-family: 'Playfair Display', serif;
        font-size: 4.5rem;
        font-weight: 400;
        line-height: 1;
        letter-spacing: 2px;
        margin-bottom: 30px;
    }
    .cover h2 {
        font-weight: 300;
        font-size: 1.2rem;
        letter-spacing: 3px;
        color: var(--text-muted);
        text-transform: uppercase;
    }
    .catalog-header {
        text-align: center;
        margin-bottom: 10mm;
        padding-bottom: 5mm;
        border-bottom: 1px solid var(--border);
    }
    .catalog-header h2 {
        font-family: 'Playfair Display', serif;
        font-weight: 400;
        letter-spacing: 2px;
        font-size: 1.8rem;
    }
    .catalog-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        grid-template-rows: repeat(3, 1fr);
        gap: 15mm;
        height: calc(100% - 25mm);
    }
    .card {
        display: flex;
        flex-direction: column;
    }
    .card-img-wrapper {
        height: 150px;
        width: 100%;
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 15px;
    }
    .card-img-wrapper img {
        max-height: 100%;
        max-width: 100%;
        object-fit: contain;
    }
    .card-brand {
        font-size: 0.65rem;
        text-transform: uppercase;
        letter-spacing: 2px;
        color: var(--text-muted);
        margin-bottom: 3px;
    }
    .card-title {
        font-family: 'Playfair Display', serif;
        font-size: 1.5rem;
        font-weight: 600;
        margin-bottom: 8px;
        line-height: 1.1;
    }
    .badge-container {
        display: flex;
        gap: 6px;
        margin-bottom: 10px;
        flex-wrap: wrap;
    }
    .badge {
        font-size: 0.55rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        padding: 3px 6px;
        border: 1px solid var(--border);
        color: var(--text-muted);
        border-radius: 2px;
    }
    .card-desc {
        font-size: 0.75rem;
        color: var(--text-main);
        margin-bottom: 10px;
        text-align: justify;
        font-style: italic;
    }
    .card-meta {
        font-size: 0.65rem;
        margin-top: auto;
        color: var(--text-muted);
        border-top: 1px solid var(--border);
        padding-top: 8px;
    }
    .card-meta span {
        display: block;
        margin-bottom: 4px;
    }
    .meta-label {
        font-weight: 600;
        color: var(--text-main);
    }
    
    @media print {
        body { background-color: white; }
        .page {
            margin: 0;
            border: initial;
            border-radius: initial;
            width: 210mm;
            height: 297mm;
            box-shadow: initial;
            background: initial;
            page-break-after: always;
        }
        @page {
            size: A4;
            margin: 0;
        }
    }
</style>
</head>
<body>

<div class="page">
    <div class="cover">
        <h3>Catálogo Oficial</h3>
        <h1>AMMAR</h1>
        <h2>Perfumería de Lujo</h2>
    </div>
</div>
"@

$itemsPerPage = 6
$totalPages = [math]::Ceiling($perfumes.Count / $itemsPerPage)

for ($p = 0; $p -lt $totalPages; $p++) {
    $html += "<div class='page'>"
    $html += "<div class='catalog-header'><h2>Colección AMMAR</h2></div>"
    $html += "<div class='catalog-grid'>"
    
    $startIdx = $p * $itemsPerPage
    $endIdx = [math]::Min($startIdx + $itemsPerPage - 1, $perfumes.Count - 1)
    
    for ($i = $startIdx; $i -le $endIdx; $i++) {
        $perf = $perfumes[$i]
        
        $topStr = ""
        if ($perf.Topologia) {
            $topLines = $perf.Topologia -split "`n"
            if ($topLines.Count -ge 2) {
                $clasif = ($topLines[0] -replace "Clasificación:\s*", "")
                $det = ($topLines[1] -replace "Detalle:\s*", "")
                $topStr = "<span class='meta-label'>Topología:</span> $clasif<br><span class='meta-label'>Estructura:</span> $det"
            }
        }
        
        $imgTag = ""
        if ($perf.Imagen) {
            $imgTag = "<img src='$($perf.Imagen)' alt='$($perf.Referencia)'>"
        } else {
            $imgTag = "<div style='color:#ccc; font-size:0.7rem'>Sin Imagen</div>"
        }
        
        $html += @"
        <div class='card'>
            <div class='card-img-wrapper'>
                $imgTag
            </div>
            <div class='card-brand'>$($perf.Marca)</div>
            <div class='card-title'>$($perf.Referencia)</div>
            <div class='badge-container'>
                <div class='badge'>$($perf.Genero)</div>
                <div class='badge'>$($perf.Familia)</div>
            </div>
            <div class='card-desc'>
                "$($perf.Tips)"
            </div>
            <div class='card-meta'>
                <span><span class='meta-label'>Notas:</span> $($perf.Notas)</span>
                $topStr
            </div>
        </div>
"@
    }
    
    $html += "</div></div>"
}

$html += @"
</body>
</html>
"@

Set-Content -Path $outHtml -Value $html -Encoding UTF8
Write-Host "Catálogo HTML generado en $outHtml"

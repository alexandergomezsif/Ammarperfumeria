$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$fotosDir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$outHtml = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Tienda_Ammar.html"
$bgImage = "file:///C:/Users/AXL/.gemini/antigravity/brain/7b566632-7a83-4b00-a9ee-98e462656def/fondo_marmol_blanco_1789860221062.jpg"

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
    
    # Escape quotes for JSON
    function Escape-JsonStr($str) {
        if (-not $str) { return "" }
        return ([string]$str).Replace("\", "\\").Replace("`"", "\`"").Replace("`n", "\n").Replace("`r", "")
    }

    $perfumes += @{
        Marca = Escape-JsonStr $marca
        Referencia = Escape-JsonStr $ref
        Genero = Escape-JsonStr $gen
        Familia = Escape-JsonStr $fam
        Notas = Escape-JsonStr $not
        Topologia = Escape-JsonStr $top
        Tips = Escape-JsonStr $tips
        Imagen = Escape-JsonStr $imgPath
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Build JSON array
$jsonArray = "["
foreach ($p in $perfumes) {
    $jsonArray += "`n  {"
    $jsonArray += "`"Marca`": `"$($p.Marca)`","
    $jsonArray += "`"Referencia`": `"$($p.Referencia)`","
    $jsonArray += "`"Genero`": `"$($p.Genero)`","
    $jsonArray += "`"Familia`": `"$($p.Familia)`","
    $jsonArray += "`"Notas`": `"$($p.Notas)`","
    $jsonArray += "`"Topologia`": `"$($p.Topologia)`","
    $jsonArray += "`"Tips`": `"$($p.Tips)`","
    $jsonArray += "`"Imagen`": `"$($p.Imagen)`""
    $jsonArray += "},"
}
$jsonArray = $jsonArray.TrimEnd(",") + "`n]"

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Tienda Interactiva - AMMAR by Karen Rico</title>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;500;600&family=Playfair+Display:ital,wght@0,400;0,600;0,700;1,400&display=swap" rel="stylesheet">
<style>
    :root {
        --text-main: #1a1a1a;
        --text-muted: #555555;
        --accent: #d4af37; /* Elegant gold */
        --border: #dddddd;
        --card-bg: rgba(255, 255, 255, 0.95);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
        font-family: 'Montserrat', sans-serif;
        color: var(--text-main);
        background-color: #f4f4f4;
        background-image: url('$bgImage');
        background-attachment: fixed;
        background-size: cover;
        background-position: center;
        line-height: 1.6;
    }
    
    /* Header & Navigation */
    header {
        background: rgba(255, 255, 255, 0.98);
        padding: 30px 20px;
        text-align: center;
        position: sticky;
        top: 0;
        z-index: 1000;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
        backdrop-filter: blur(5px);
    }
    header h1 {
        font-family: 'Playfair Display', serif;
        font-size: 2.8rem;
        letter-spacing: 4px;
        text-transform: uppercase;
        margin-bottom: 5px;
    }
    header h2 {
        font-weight: 300;
        font-size: 1rem;
        letter-spacing: 5px;
        color: var(--text-muted);
        text-transform: uppercase;
    }
    
    /* Controls */
    .controls {
        display: flex;
        justify-content: center;
        gap: 15px;
        margin-top: 20px;
        max-width: 800px;
        margin-left: auto;
        margin-right: auto;
    }
    input, select {
        padding: 12px 20px;
        border: 1px solid var(--border);
        border-radius: 30px;
        font-family: 'Montserrat', sans-serif;
        font-size: 0.9rem;
        outline: none;
        transition: border-color 0.3s;
    }
    input:focus, select:focus {
        border-color: var(--text-main);
    }
    input[type="text"] {
        flex: 1;
        min-width: 250px;
    }

    /* Container */
    .container {
        max-width: 1200px;
        margin: 40px auto;
        padding: 0 20px;
    }

    /* Brand Sections */
    .brand-section {
        margin-bottom: 60px;
    }
    .brand-header {
        text-align: center;
        margin-bottom: 40px;
        padding-bottom: 15px;
        border-bottom: 2px solid var(--text-main);
    }
    .brand-header h2 {
        font-family: 'Playfair Display', serif;
        font-size: 2.5rem;
        letter-spacing: 3px;
        text-transform: uppercase;
    }
    .brand-logo-placeholder {
        font-size: 0.8rem;
        color: var(--text-muted);
        letter-spacing: 2px;
        margin-top: 5px;
        text-transform: uppercase;
    }

    /* Perfume Cards (Alternating Layout) */
    .perfume-card {
        display: flex;
        background: var(--card-bg);
        border: 1px solid var(--border);
        border-radius: 8px;
        overflow: hidden;
        margin-bottom: 40px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.03);
        transition: transform 0.3s;
    }
    .perfume-card:hover {
        transform: translateY(-5px);
    }
    .perfume-card.reverse {
        flex-direction: row-reverse;
    }
    
    .card-image {
        flex: 1;
        max-width: 45%;
        padding: 30px;
        display: flex;
        justify-content: center;
        align-items: center;
        background: #ffffff;
        border-right: 1px solid var(--border);
    }
    .perfume-card.reverse .card-image {
        border-right: none;
        border-left: 1px solid var(--border);
    }
    .card-image img {
        max-width: 100%;
        max-height: 400px;
        object-fit: contain;
    }
    
    .card-content {
        flex: 1;
        padding: 40px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    .badge-container {
        display: flex;
        gap: 8px;
        margin-bottom: 15px;
    }
    .badge {
        font-size: 0.65rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        padding: 4px 10px;
        border: 1px solid var(--text-main);
        color: var(--text-main);
        border-radius: 20px;
    }
    
    .perfume-title {
        font-family: 'Playfair Display', serif;
        font-size: 2.2rem;
        font-weight: 600;
        margin-bottom: 15px;
        line-height: 1.1;
    }
    
    .perfume-pitch {
        font-size: 1.05rem;
        color: var(--text-muted);
        font-style: italic;
        margin-bottom: 25px;
        text-align: justify;
    }
    
    .perfume-details {
        margin-top: auto;
        padding-top: 20px;
        border-top: 1px solid var(--border);
    }
    
    .detail-group {
        margin-bottom: 10px;
        font-size: 0.85rem;
    }
    .detail-label {
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
        font-size: 0.75rem;
        display: block;
        margin-bottom: 2px;
        color: var(--text-main);
    }
    .topologia-text {
        white-space: pre-wrap;
    }

    /* Print Styles */
    @media print {
        body {
            background: white !important;
            color: black;
        }
        header {
            position: static;
            box-shadow: none;
            border-bottom: 2px solid black;
            padding: 20px 0;
            margin-bottom: 30px;
        }
        .controls {
            display: none; /* Hide interactive elements when printing */
        }
        .container {
            margin: 0;
            padding: 0;
            max-width: 100%;
        }
        .brand-section {
            page-break-before: always;
        }
        .brand-section:first-child {
            page-break-before: auto;
        }
        .perfume-card {
            box-shadow: none;
            border: 1px solid #ccc;
            page-break-inside: avoid; /* Prevents text bleeding across pages */
            margin-bottom: 20mm;
            border-radius: 0;
        }
        .card-image {
            max-width: 40%;
            padding: 10px;
        }
        .card-content {
            padding: 20px;
        }
        .perfume-title {
            font-size: 1.8rem;
        }
        .perfume-pitch {
            font-size: 0.9rem;
        }
        /* Optimize layout specifically for A4 Print */
        @page {
            size: A4;
            margin: 15mm;
        }
    }
</style>
</head>
<body>

<header>
    <h1>AMMAR</h1>
    <h2>By Karen Rico</h2>
    
    <div class="controls">
        <input type="text" id="searchInput" placeholder="Buscar fragancia, notas o estilo...">
        <select id="brandSelect">
            <option value="ALL">Todas las Marcas</option>
        </select>
        <select id="genderSelect">
            <option value="ALL">Todos los Géneros</option>
            <option value="Caballero">Caballero</option>
            <option value="Dama">Dama</option>
            <option value="Unisex">Unisex</option>
        </select>
    </div>
</header>

<div class="container" id="catalogContainer">
    <!-- Generated dynamically via JS -->
</div>

<script>
    const data = $jsonArray;
    
    const catalogContainer = document.getElementById('catalogContainer');
    const searchInput = document.getElementById('searchInput');
    const brandSelect = document.getElementById('brandSelect');
    const genderSelect = document.getElementById('genderSelect');
    
    // Extract unique brands and populate dropdown
    const uniqueBrands = [...new Set(data.map(p => p.Marca))].sort();
    uniqueBrands.forEach(brand => {
        const option = document.createElement('option');
        option.value = brand;
        option.textContent = brand;
        brandSelect.appendChild(option);
    });

    function renderCatalog() {
        catalogContainer.innerHTML = '';
        
        const searchTerm = searchInput.value.toLowerCase();
        const brandFilter = brandSelect.value;
        const genderFilter = genderSelect.value;
        
        // Filter Data
        let filtered = data.filter(p => {
            const matchesSearch = p.Referencia.toLowerCase().includes(searchTerm) || 
                                  p.Notas.toLowerCase().includes(searchTerm) ||
                                  p.Tips.toLowerCase().includes(searchTerm);
            const matchesBrand = brandFilter === 'ALL' || p.Marca === brandFilter;
            const matchesGender = genderFilter === 'ALL' || p.Genero === genderFilter;
            return matchesSearch && matchesBrand && matchesGender;
        });

        // Group by Brand
        const grouped = {};
        filtered.forEach(p => {
            if (!grouped[p.Marca]) grouped[p.Marca] = [];
            grouped[p.Marca].push(p);
        });

        // Sort Brands Alphabetically
        const sortedBrands = Object.keys(grouped).sort();

        if (sortedBrands.length === 0) {
            catalogContainer.innerHTML = '<p style="text-align:center; padding: 50px;">No se encontraron resultados.</p>';
            return;
        }

        sortedBrands.forEach(brand => {
            // Brand Section Header
            const section = document.createElement('div');
            section.className = 'brand-section';
            
            section.innerHTML = `
                <div class="brand-header">
                    <h2>\${brand}</h2>
                    <div class="brand-logo-placeholder">Colección Oficial</div>
                </div>
            `;
            
            // Perfumes
            grouped[brand].forEach((perf, index) => {
                const isReverse = index % 2 !== 0 ? 'reverse' : '';
                
                const imgSrc = perf.Imagen ? perf.Imagen : 'https://via.placeholder.com/400x500?text=Falta+Foto';
                
                // Format Topologia gracefully
                let topHTML = '';
                if (perf.Topologia) {
                    // It already has newlines \n, we can use pre-wrap CSS
                    topHTML = `<div class="detail-group"><span class="detail-label">Arquitectura</span><span class="topologia-text">\${perf.Topologia}</span></div>`;
                }

                section.innerHTML += `
                    <article class="perfume-card \${isReverse}">
                        <div class="card-image">
                            <img src="\${imgSrc}" alt="\${perf.Referencia}">
                        </div>
                        <div class="card-content">
                            <div class="badge-container">
                                <span class="badge">\${perf.Genero}</span>
                                <span class="badge">\${perf.Familia}</span>
                            </div>
                            <h3 class="perfume-title">\${perf.Referencia}</h3>
                            <p class="perfume-pitch">"\${perf.Tips}"</p>
                            
                            <div class="perfume-details">
                                <div class="detail-group">
                                    <span class="detail-label">Notas Principales</span>
                                    \${perf.Notas}
                                </div>
                                \${topHTML}
                            </div>
                        </div>
                    </article>
                `;
            });
            
            catalogContainer.appendChild(section);
        });
    }

    // Event Listeners
    searchInput.addEventListener('input', renderCatalog);
    brandSelect.addEventListener('change', renderCatalog);
    genderSelect.addEventListener('change', renderCatalog);

    // Initial Render
    renderCatalog();
</script>

</body>
</html>
"@

Set-Content -Path $outHtml -Value $html -Encoding UTF8
Write-Host "Tienda generada en $outHtml"

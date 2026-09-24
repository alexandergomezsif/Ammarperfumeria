$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$fotosDir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$outHtml = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\index.html"

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

$html = @"
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AMMAR by Karen Rico - Catalogo Oficial</title>
    
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500;600&family=Montserrat:ital,wght@0,300;0,400;0,500;1,300&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --bg: #fafafa;
            --surface: #ffffff;
            --text-dark: #111111;
            --text-light: #666666;
            --border: #eaeaea;
            --font-heading: 'Cinzel', serif;
            --font-body: 'Montserrat', sans-serif;
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body { font-family: var(--font-body); background-color: var(--bg); color: var(--text-dark); line-height: 1.6; }
        .sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
        
        .main-header { background-color: var(--surface); padding: 20px 40px; display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid var(--border); position: sticky; top: 0; z-index: 100; }
        .brand-logo { font-family: var(--font-heading); font-size: 2rem; font-weight: 500; letter-spacing: 4px; text-transform: uppercase; }
        .brand-subtitle { font-size: 0.7rem; letter-spacing: 3px; text-transform: uppercase; color: var(--text-light); display: block; text-align: center; margin-top: -5px; }
        .header-meta { font-size: 0.8rem; letter-spacing: 2px; text-transform: uppercase; color: var(--text-light); }
        
        .hero-section { text-align: center; padding: 80px 20px; background: linear-gradient(to bottom, var(--surface) 0%, var(--bg) 100%); }
        .hero-section h1 { font-family: var(--font-heading); font-size: 3.5rem; font-weight: 400; margin-bottom: 20px; }
        .hero-section p { color: var(--text-light); max-width: 600px; margin: 0 auto; font-size: 1.1rem; }
        
        .catalog-search { max-width: 1400px; margin: 0 auto 40px auto; padding: 0 20px; display: flex; gap: 20px; justify-content: center; flex-wrap: wrap; }
        .filter-control { padding: 12px 20px; border: 1px solid var(--border); background-color: var(--surface); font-family: var(--font-body); font-size: 0.9rem; color: var(--text-dark); outline: none; min-width: 200px; transition: border-color 0.3s; }
        .filter-control:focus { border-color: var(--text-dark); }
        #searchInput { flex: 1; max-width: 400px; }
        
        .product-grid { max-width: 1400px; margin: 0 auto; padding: 0 20px 80px 20px; display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 40px; }
        .empty-state { grid-column: 1 / -1; text-align: center; padding: 50px; color: var(--text-light); }
        
        .product-card { background-color: var(--surface); border: 1px solid var(--border); padding: 30px; display: flex; flex-direction: column; transition: transform 0.4s ease, box-shadow 0.4s ease; }
        .product-card:hover { transform: translateY(-5px); box-shadow: 0 15px 35px rgba(0,0,0,0.05); }
        .product-image { width: 100%; height: 300px; object-fit: contain; margin-bottom: 25px; mix-blend-mode: multiply; }
        .product-image.error-fallback { content: url("https://via.placeholder.com/300x400?text=AMMAR"); }
        
        .product-brand { font-size: 0.75rem; letter-spacing: 2px; text-transform: uppercase; color: var(--text-light); margin-bottom: 5px; }
        .product-title { font-family: var(--font-heading); font-size: 1.5rem; font-weight: 600; margin-bottom: 10px; line-height: 1.2; }
        .product-price { font-size: 1.3rem; font-weight: 600; color: #111; margin-bottom: 15px; }
        
        .product-tags { display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap; }
        .tag { font-size: 0.65rem; letter-spacing: 1px; text-transform: uppercase; padding: 4px 8px; background-color: var(--bg); border: 1px solid var(--border); color: var(--text-light); }
        
        .product-pitch { font-size: 0.85rem; font-style: italic; color: var(--text-light); margin-bottom: 20px; flex-grow: 1; }
        
        .product-meta { font-size: 0.75rem; border-top: 1px solid var(--border); padding-top: 15px; }
        .meta-line { margin-bottom: 5px; }
        .meta-label { font-weight: 500; color: var(--text-dark); text-transform: uppercase; font-size: 0.65rem; letter-spacing: 1px; }
        
        .btn-action { display: inline-block; margin-top: 15px; padding: 12px 24px; background-color: var(--text-dark); color: var(--surface); text-decoration: none; font-family: var(--font-body); font-size: 0.85rem; font-weight: 500; letter-spacing: 2px; text-transform: uppercase; text-align: center; border: 1px solid var(--text-dark); transition: all 0.3s ease; cursor: pointer; width: 100%; }
        .btn-action:hover { background-color: var(--surface); color: var(--text-dark); outline: 2px solid var(--text-dark); outline-offset: 2px; }
        
        .main-footer { background-color: var(--text-dark); color: white; text-align: center; padding: 40px 20px; font-size: 0.85rem; }
        
        @media print {
            body { background: white !important; }
            .main-header, .hero-section, [role="search"], .main-footer, .btn-action { display: none !important; }
            .product-grid { display: block; padding: 0; }
            .product-card { page-break-inside: avoid; margin-bottom: 30px; border: 1px solid #ccc; box-shadow: none; display: flex; flex-direction: row; align-items: center; gap: 20px; padding: 20px; }
            .product-image { width: 30%; height: 200px; margin-bottom: 0; }
            .product-info-wrapper { width: 70%; }
        }
    </style>
</head>
<body>
    <div id="ariaLiveRegion" class="sr-only" aria-live="polite"></div>

    <header class="main-header">
        <div>
            <div class="brand-logo">AMMAR</div>
            <span class="brand-subtitle">By Karen Rico</span>
        </div>
        <div class="header-meta">Catalogo Oficial</div>
    </header>

    <section class="hero-section" aria-labelledby="hero-title">
        <h1 id="hero-title">Nuestra Coleccion</h1>
        <p>Descubre un universo de fragancias exclusivas. Filtra por marca, genero o notas olfativas para encontrar el aroma perfecto que defina tu estilo.</p>
    </section>

    <div role="search">
        <form class="catalog-search" onsubmit="event.preventDefault();">
            <div>
                <label for="searchInput" class="sr-only">Buscar fragancia o notas</label>
                <input type="search" id="searchInput" class="filter-control" placeholder="Buscar fragancia, notas o inspiracion...">
            </div>
            <div>
                <label for="brandSelect" class="sr-only">Filtrar por Marca</label>
                <select id="brandSelect" class="filter-control">
                    <option value="ALL">Todas las Marcas</option>
                </select>
            </div>
            <div>
                <label for="genderSelect" class="sr-only">Filtrar por Genero</label>
                <select id="genderSelect" class="filter-control">
                    <option value="ALL">Todos los Generos</option>
                    <option value="Caballero">Caballero</option>
                    <option value="Dama">Dama</option>
                    <option value="Unisex">Unisex</option>
                </select>
            </div>
        </form>
    </div>

    <main class="product-grid" id="productGrid" aria-live="polite">
        <div class="empty-state">Inicializando catalogo...</div>
    </main>

    <footer class="main-footer">
        <p>&copy; 2026 AMMAR by Karen Rico. Todos los derechos reservados.</p>
    </footer>

    <script>
        const perfumesData = $jsonStr;
        
        document.addEventListener('DOMContentLoaded', () => {
            const grid = document.getElementById('productGrid');
            const searchInput = document.getElementById('searchInput');
            const brandSelect = document.getElementById('brandSelect');
            const genderSelect = document.getElementById('genderSelect');
            const liveRegion = document.getElementById('ariaLiveRegion');
            
            if (!perfumesData || perfumesData.length === 0) {
                grid.innerHTML = '<div class="empty-state">No se pudieron cargar los datos. Revisa la base de datos de Excel.</div>';
                return;
            }

            function init() {
                populateBrandDropdown();
                bindEvents();
                renderProducts();
            }

            function populateBrandDropdown() {
                const uniqueBrands = [...new Set(perfumesData.map(p => p.Marca))].sort();
                uniqueBrands.forEach(brand => {
                    const option = document.createElement('option');
                    option.value = brand;
                    option.textContent = brand;
                    brandSelect.appendChild(option);
                });
            }

            function bindEvents() {
                searchInput.addEventListener('input', renderProducts);
                brandSelect.addEventListener('change', renderProducts);
                genderSelect.addEventListener('change', renderProducts);
                
                grid.addEventListener('error', (e) => {
                    if (e.target.tagName && e.target.tagName.toLowerCase() === 'img') {
                        e.target.classList.add('error-fallback');
                        e.target.alt = 'Imagen no disponible';
                    }
                }, true);
            }

            function renderProducts() {
                grid.innerHTML = '';
                
                const term = searchInput.value.toLowerCase();
                const brand = brandSelect.value;
                const gender = genderSelect.value;

                const filtered = perfumesData.filter(p => {
                    const ref = (p.Referencia || "").toLowerCase();
                    const not = (p.Notas || "").toLowerCase();
                    const tip = (p.Tips || "").toLowerCase();
                    
                    const matchSearch = ref.includes(term) || not.includes(term) || tip.includes(term);
                    const matchBrand = brand === 'ALL' || p.Marca === brand;
                    const matchGender = gender === 'ALL' || p.Genero === gender;
                    return matchSearch && matchBrand && matchGender;
                });

                liveRegion.textContent = 'Mostrando ' + filtered.length + ' fragancias.';

                if (filtered.length === 0) {
                    grid.innerHTML = '<div class="empty-state">No se encontraron fragancias con esos filtros.</div>';
                    return;
                }

                const fragment = document.createDocumentFragment();
                
                const formatter = new Intl.NumberFormat('es-CO', {
                    style: 'currency',
                    currency: 'COP',
                    minimumFractionDigits: 0
                });

                filtered.forEach(p => {
                    let topHTML = '';
                    if (p.Topologia && p.Topologia.trim() !== '') {
                        const cleanTopologia = p.Topologia.replace('Clasificación:', '').replace('Detalle:', '<br>');
                        topHTML = '<div class="meta-line"><span class="meta-label">Arquitectura:</span> ' + cleanTopologia + '</div>';
                    }
                    
                    const precioFmt = formatter.format(p.Precio);
                    
                    const wpText = 'Hola, me interesa el perfume ' + encodeURIComponent(p.Marca) + ' - ' + encodeURIComponent(p.Referencia);
                    const wpUrl = 'https://wa.me/573000000000?text=' + wpText;

                    const article = document.createElement('article');
                    article.className = 'product-card';
                    
                    article.innerHTML = `
                        <img src="\${p.Imagen}" alt="\${p.Referencia}" class="product-image" loading="lazy">
                        
                        <div class="product-info-wrapper">
                            <div class="product-brand">\${p.Marca}</div>
                            <h2 class="product-title">\${p.Referencia}</h2>
                            
                            <div class="product-price">
                                \${precioFmt}
                            </div>
                            
                            <div class="product-tags">
                                <span class="tag">\${p.Genero}</span>
                                <span class="tag">\${p.Familia}</span>
                            </div>
                            
                            <p class="product-pitch">"\${p.Tips}"</p>
                            
                            <div class="product-meta">
                                <div class="meta-line"><span class="meta-label">Notas:</span> \${p.Notas}</div>
                                \${topHTML}
                            </div>
                            
                            <a href="\${wpUrl}" target="_blank" rel="noopener noreferrer" class="btn-action" aria-label="Solicitar \${p.Referencia} por WhatsApp">
                                Lo Quiero
                            </a>
                        </div>
                    `;
                    fragment.appendChild(article);
                });

                grid.appendChild(fragment);
            }

            init();
        });
    </script>
</body>
</html>
"@

# Usamos Out-File con encoding utf8 para evitar problemas de BOM en WriteAllText
$html | Out-File -FilePath "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\index.html" -Encoding utf8
Write-Host "Index monolithico generado"

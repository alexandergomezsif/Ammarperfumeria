# -*- coding: utf-8 -*-
import pandas as pd
import os, random, json, sys

# Forzar salida UTF-8
sys.stdout.reconfigure(encoding='utf-8')

excel_path = 'Base_Datos_Perfumes.xlsx'
fotos_dir  = 'fotos'
output     = 'index.html'

# 1. Mapeo de imagenes relativas
img_dict = {}
for root, dirs, files in os.walk(fotos_dir):
    for f in files:
        if f.lower().endswith('.jpg'):
            ref_name = f[:-4].replace('_', ' ').upper().strip()
            rel_path = 'fotos/' + os.path.basename(root) + '/' + f
            img_dict[ref_name] = rel_path

# 2. Leer Excel
df = pd.read_excel(excel_path).fillna('')
perfumes = []
last_marca = 'Desconocido'
random.seed(42)

for _, row in df.iterrows():
    marca = str(row.iloc[0]).strip() if str(row.iloc[0]).strip() else last_marca
    last_marca = marca
    ref = str(row.iloc[1]).strip()
    if not ref: continue
    if ref.endswith('.0'): ref = ref[:-2]

    gen   = str(row.iloc[2]).strip()
    fam   = str(row.iloc[3]).strip()
    notas = str(row.iloc[4]).strip()
    top   = str(row.iloc[5]).strip()
    tips  = str(row.iloc[6]).strip()

    ref_key = ref.upper()
    img = img_dict.get(ref_key, '')
    if not img:
        for k, v in img_dict.items():
            if ref_key in k:
                img = v; break
    if not img:
        img = 'https://via.placeholder.com/300x400/f5f5f5/111111?text=' + ref.replace(' ', '+')

    precio = random.randint(160, 250) * 1000

    perfumes.append({
        'marca': marca, 'referencia': ref, 'genero': gen,
        'familia': fam, 'notas': notas, 'topologia': top,
        'tips': tips, 'imagen': img, 'precio': precio
    })

# 3. Convertir a JSON limpio UTF-8
json_data = json.dumps(perfumes, ensure_ascii=False)

# 4. Construir HTML completo (sin Jinja2, concatenacion directa)
html = '''<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AMMAR by Karen Rico - Catalogo Oficial</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500;600&family=Montserrat:wght@300;400;500;600&display=swap" rel="stylesheet">
<style>
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#fafafa;--surface:#fff;--dark:#111;--muted:#666;--border:#e8e8e8;
  --accent:#c8a97e;
  --font-h:'Cinzel',serif;--font-b:'Montserrat',sans-serif;
}
body{font-family:var(--font-b);background:var(--bg);color:var(--dark);line-height:1.6}
.sr-only{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0,0,0,0)}

/* NAV */
nav{background:var(--surface);border-bottom:1px solid var(--border);padding:16px 40px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:200}
.logo{font-family:var(--font-h);font-size:1.8rem;letter-spacing:6px;font-weight:500}
.logo span{display:block;font-family:var(--font-b);font-size:.65rem;letter-spacing:4px;color:var(--muted);font-weight:300;text-transform:uppercase;text-align:center}
.nav-info{font-size:.75rem;letter-spacing:2px;text-transform:uppercase;color:var(--muted)}

/* HERO */
.hero{padding:80px 20px 60px;text-align:center;background:linear-gradient(180deg,#fff 0%,var(--bg) 100%)}
.hero h1{font-family:var(--font-h);font-size:3rem;font-weight:400;margin-bottom:16px;letter-spacing:2px}
.hero p{color:var(--muted);max-width:540px;margin:0 auto;font-size:1rem;font-weight:300}

/* FILTERS */
.filters{padding:0 20px 50px;display:flex;gap:16px;justify-content:center;flex-wrap:wrap;max-width:1200px;margin:0 auto}
.f-input,.f-select{
  padding:11px 18px;border:1px solid var(--border);background:var(--surface);
  font-family:var(--font-b);font-size:.85rem;color:var(--dark);outline:none;
  min-width:180px;transition:border-color .25s
}
.f-input:focus,.f-select:focus{border-color:var(--dark)}
.f-input{flex:1;max-width:360px}

/* STATS */
.stats{text-align:center;font-size:.75rem;letter-spacing:1px;text-transform:uppercase;color:var(--muted);margin-bottom:40px}

/* GRID */
.grid{
  max-width:1400px;margin:0 auto;padding:0 20px 100px;
  display:grid;grid-template-columns:repeat(auto-fill,minmax(290px,1fr));gap:32px
}
.empty{grid-column:1/-1;text-align:center;padding:60px;color:var(--muted)}

/* CARD */
.card{
  background:var(--surface);border:1px solid var(--border);
  display:flex;flex-direction:column;cursor:default;
  transition:transform .35s,box-shadow .35s
}
.card:hover{transform:translateY(-6px);box-shadow:0 20px 40px rgba(0,0,0,.06)}
.card-img-wrap{padding:24px;display:flex;align-items:center;justify-content:center;height:280px;background:#fcfcfc;border-bottom:1px solid var(--border)}
.card-img-wrap img{max-height:100%;max-width:100%;object-fit:contain;mix-blend-mode:multiply}
.card-body{padding:24px;display:flex;flex-direction:column;flex:1}
.card-brand{font-size:.65rem;letter-spacing:2.5px;text-transform:uppercase;color:var(--muted);margin-bottom:6px}
.card-title{font-family:var(--font-h);font-size:1.3rem;font-weight:600;margin-bottom:12px;line-height:1.2}
.card-price{font-size:1.2rem;font-weight:600;color:var(--dark);margin-bottom:14px}
.tags{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:16px}
.tag{font-size:.6rem;letter-spacing:1px;text-transform:uppercase;padding:3px 8px;border:1px solid var(--border);color:var(--muted)}
.tag-g{border-color:var(--accent);color:var(--accent)}
.card-pitch{font-size:.82rem;font-style:italic;color:var(--muted);margin-bottom:auto;padding-bottom:16px;line-height:1.7}
.card-meta{border-top:1px solid var(--border);padding-top:14px;margin-top:16px;font-size:.72rem;color:var(--muted)}
.meta-row{margin-bottom:6px;line-height:1.5}
.meta-label{font-weight:600;color:var(--dark);text-transform:uppercase;font-size:.6rem;letter-spacing:1px}

/* BUTTON */
.btn{
  display:block;margin-top:18px;padding:13px;background:var(--dark);color:#fff;
  text-align:center;text-decoration:none;font-family:var(--font-b);font-size:.78rem;
  font-weight:500;letter-spacing:2.5px;text-transform:uppercase;
  border:1px solid var(--dark);transition:all .25s
}
.btn:hover{background:#fff;color:var(--dark)}

/* FOOTER */
footer{background:var(--dark);color:#aaa;text-align:center;padding:40px 20px;font-size:.78rem;letter-spacing:1px}

/* MODAL (detalle) */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:500;display:none;align-items:center;justify-content:center;padding:20px}
.modal-overlay.open{display:flex}
.modal{background:#fff;max-width:800px;width:100%;max-height:90vh;overflow-y:auto;display:flex;gap:0;position:relative}
.modal-close{position:absolute;top:16px;right:20px;background:none;border:none;font-size:1.5rem;cursor:pointer;color:var(--dark);z-index:10}
.modal-img{width:340px;min-width:280px;background:#fcfcfc;display:flex;align-items:center;justify-content:center;padding:30px;border-right:1px solid var(--border)}
.modal-img img{max-width:100%;max-height:400px;object-fit:contain;mix-blend-mode:multiply}
.modal-content{padding:40px;flex:1}
.modal-brand{font-size:.65rem;letter-spacing:2px;text-transform:uppercase;color:var(--muted);margin-bottom:8px}
.modal-title{font-family:var(--font-h);font-size:2rem;font-weight:600;margin-bottom:8px;line-height:1.1}
.modal-price{font-size:1.5rem;font-weight:600;margin-bottom:20px}
.modal-tags{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:20px}
.modal-pitch{font-size:.95rem;font-style:italic;color:var(--muted);line-height:1.8;margin-bottom:24px;border-left:3px solid var(--accent);padding-left:16px}
.modal-section{margin-bottom:16px}
.modal-section-label{font-size:.65rem;letter-spacing:2px;text-transform:uppercase;font-weight:600;color:var(--dark);margin-bottom:4px}
.modal-section-text{font-size:.85rem;color:var(--muted);line-height:1.6}
.modal-btn{display:block;width:100%;margin-top:24px;padding:14px;background:var(--dark);color:#fff;text-align:center;text-decoration:none;font-family:var(--font-b);font-size:.82rem;font-weight:500;letter-spacing:2px;text-transform:uppercase;border:1px solid var(--dark);transition:all .25s}
.modal-btn:hover{background:#fff;color:var(--dark)}

@media(max-width:600px){
  .modal{flex-direction:column}
  .modal-img{width:100%;min-width:unset;border-right:none;border-bottom:1px solid var(--border);height:240px}
  .logo{font-size:1.4rem}
  .hero h1{font-size:2rem}
  nav{padding:14px 20px}
}

@media print{
  nav,.filters,.btn,.modal-overlay{display:none!important}
  .card{break-inside:avoid;border:1px solid #ccc;box-shadow:none!important;transform:none!important}
  body{background:#fff}
}
</style>
</head>
<body>

<div id="ariaLive" class="sr-only" aria-live="polite"></div>
<div id="modalOverlay" class="modal-overlay" role="dialog" aria-modal="true" aria-label="Detalle del perfume">
  <div class="modal" id="modalBox">
    <button class="modal-close" id="modalClose" aria-label="Cerrar">&times;</button>
    <div class="modal-img"><img id="modalImg" src="" alt=""></div>
    <div class="modal-content">
      <div class="modal-brand" id="modalBrand"></div>
      <h2 class="modal-title" id="modalTitle"></h2>
      <div class="modal-price" id="modalPrice"></div>
      <div class="modal-tags" id="modalTags"></div>
      <p class="modal-pitch" id="modalPitch"></p>
      <div class="modal-section"><div class="modal-section-label">Notas Principales</div><div class="modal-section-text" id="modalNotas"></div></div>
      <div class="modal-section" id="modalTopWrap"><div class="modal-section-label">Arquitectura Olfativa</div><div class="modal-section-text" id="modalTop"></div></div>
      <a href="" target="_blank" rel="noopener noreferrer" class="modal-btn" id="modalBtn">Lo Quiero &rarr; WhatsApp</a>
    </div>
  </div>
</div>

<nav>
  <div>
    <div class="logo">AMMAR<span>By Karen Rico</span></div>
  </div>
  <div class="nav-info">Catalogo Oficial</div>
</nav>

<header class="hero">
  <h1>Nuestra Coleccion</h1>
  <p>Descubre un universo de fragancias exclusivas. Filtra, busca y encuentra tu aroma perfecto.</p>
</header>

<section>
  <form class="filters" onsubmit="return false;" role="search">
    <label for="searchInput" class="sr-only">Buscar</label>
    <input type="search" id="searchInput" class="f-input" placeholder="Buscar por nombre, notas o estilo...">
    <label for="brandSel" class="sr-only">Marca</label>
    <select id="brandSel" class="f-select">
      <option value="ALL">Todas las Marcas</option>
    </select>
    <label for="genderSel" class="sr-only">Genero</label>
    <select id="genderSel" class="f-select">
      <option value="ALL">Todos los Generos</option>
      <option value="Caballero">Caballero</option>
      <option value="Dama">Dama</option>
      <option value="Unisex">Unisex</option>
    </select>
    <label for="famSel" class="sr-only">Familia</label>
    <select id="famSel" class="f-select">
      <option value="ALL">Todas las Familias</option>
    </select>
  </form>
  <div class="stats" id="statsBar"></div>
</section>

<main class="grid" id="grid"></main>

<footer>
  <p>&copy; 2026 AMMAR by Karen Rico &mdash; Todos los derechos reservados</p>
</footer>

<script>
const DATA = ''' + json_data + ''';

(function(){
  const grid = document.getElementById('grid');
  const searchInput = document.getElementById('searchInput');
  const brandSel = document.getElementById('brandSel');
  const genderSel = document.getElementById('genderSel');
  const famSel = document.getElementById('famSel');
  const statsBar = document.getElementById('statsBar');
  const ariaLive = document.getElementById('ariaLive');

  // Formato COP
  const fmt = new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',minimumFractionDigits:0});

  // --- MODAL ---
  const overlay = document.getElementById('modalOverlay');
  const modalClose = document.getElementById('modalClose');

  function openModal(p){
    document.getElementById('modalImg').src = p.imagen;
    document.getElementById('modalImg').alt = p.referencia;
    document.getElementById('modalBrand').textContent = p.marca;
    document.getElementById('modalTitle').textContent = p.referencia;
    document.getElementById('modalPrice').textContent = fmt.format(p.precio);
    document.getElementById('modalNotas').textContent = p.notas;
    
    const tagsEl = document.getElementById('modalTags');
    tagsEl.innerHTML = '';
    [p.genero, p.familia].filter(Boolean).forEach(t=>{
      const s = document.createElement('span');
      s.className = 'tag';
      s.textContent = t;
      tagsEl.appendChild(s);
    });
    
    document.getElementById('modalPitch').textContent = '"' + p.tips + '"';
    
    const topWrap = document.getElementById('modalTopWrap');
    if(p.topologia && p.topologia.trim()){
      document.getElementById('modalTop').textContent = p.topologia;
      topWrap.style.display = '';
    } else {
      topWrap.style.display = 'none';
    }
    
    const wpText = 'Hola, me interesa el perfume ' + p.marca + ' - ' + p.referencia;
    document.getElementById('modalBtn').href = 'https://wa.me/573000000000?text=' + encodeURIComponent(wpText);
    
    overlay.classList.add('open');
    document.body.style.overflow = 'hidden';
    modalClose.focus();
  }

  function closeModal(){
    overlay.classList.remove('open');
    document.body.style.overflow = '';
  }

  modalClose.addEventListener('click', closeModal);
  overlay.addEventListener('click', e => { if(e.target === overlay) closeModal(); });
  document.addEventListener('keydown', e => { if(e.key === 'Escape') closeModal(); });

  // --- POBLAR SELECTS ---
  const brands = [...new Set(DATA.map(p=>p.marca))].sort();
  brands.forEach(b=>{ const o=document.createElement('option'); o.value=b; o.textContent=b; brandSel.appendChild(o); });

  const fams = [...new Set(DATA.map(p=>p.familia).filter(Boolean))].sort();
  fams.forEach(f=>{ const o=document.createElement('option'); o.value=f; o.textContent=f; famSel.appendChild(o); });

  // --- RENDER ---
  function render(){
    const term = searchInput.value.toLowerCase().trim();
    const brand = brandSel.value;
    const gender = genderSel.value;
    const fam = famSel.value;

    const filtered = DATA.filter(p=>{
      if(brand !== 'ALL' && p.marca !== brand) return false;
      if(gender !== 'ALL' && p.genero !== gender) return false;
      if(fam !== 'ALL' && p.familia !== fam) return false;
      if(term){
        const haystack = (p.referencia+' '+p.notas+' '+p.tips+' '+p.familia+' '+p.genero).toLowerCase();
        if(!haystack.includes(term)) return false;
      }
      return true;
    });

    statsBar.textContent = filtered.length + ' fragancia' + (filtered.length !== 1 ? 's' : '') + ' encontrada' + (filtered.length !== 1 ? 's' : '');
    ariaLive.textContent = 'Mostrando ' + filtered.length + ' fragancias.';

    grid.innerHTML = '';

    if(!filtered.length){
      grid.innerHTML = '<div class="empty">No se encontraron fragancias con esos criterios.</div>';
      return;
    }

    const frag = document.createDocumentFragment();
    filtered.forEach(p=>{
      const art = document.createElement('article');
      art.className = 'card';
      art.innerHTML = `
        <div class="card-img-wrap">
          <img src="${p.imagen}" alt="${p.referencia}" loading="lazy" onerror="this.src='https://via.placeholder.com/300x400/f5f5f5/999?text=AMMAR'">
        </div>
        <div class="card-body">
          <div class="card-brand">${p.marca}</div>
          <div class="card-title">${p.referencia}</div>
          <div class="card-price">${fmt.format(p.precio)}</div>
          <div class="tags">
            <span class="tag tag-g">${p.genero}</span>
            <span class="tag">${p.familia}</span>
          </div>
          <p class="card-pitch">"${p.tips.length > 160 ? p.tips.substring(0,160)+'...' : p.tips}"</p>
          <div class="card-meta">
            <div class="meta-row"><span class="meta-label">Notas: </span>${p.notas}</div>
          </div>
          <button class="btn" onclick="openCard(${DATA.indexOf(p)})">Ver Detalle</button>
        </div>
      `;
      frag.appendChild(art);
    });
    grid.appendChild(frag);

    window._filtered = filtered;
  }

  window.openCard = function(globalIdx){
    openModal(DATA[globalIdx]);
  };

  searchInput.addEventListener('input', render);
  brandSel.addEventListener('change', render);
  genderSel.addEventListener('change', render);
  famSel.addEventListener('change', render);

  render();
})();
</script>
</body>
</html>'''

with open(output, 'w', encoding='utf-8') as f:
    f.write(html)

print("Generados " + str(len(perfumes)) + " perfumes. index.html listo.")
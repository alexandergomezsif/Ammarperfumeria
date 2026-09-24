# -*- coding: utf-8 -*-
# AMMAR Perfumeria - Site Builder DESIGN 2
# Fonts: Cinzel + Montserrat

import pandas as pd
import os, random, json, sys

EXCEL = 'Base_Datos_Perfumes.xlsx'
FOTOS = 'fotos'
OUT   = 'index.html'
WP_NUM = '573128586188'
WP_DISP = '+57 312 858 6188'
ADDR = 'Cra. 86b #47a-40, Barrio Floresta, Medellín'

img_map = {}
for root, _, files in os.walk(FOTOS):
    for f in files:
        if f.lower().endswith('.jpg'):
            key = f[:-4].replace('_',' ').upper().strip()
            img_map[key] = 'fotos/' + os.path.basename(root) + '/' + f

def find_img(ref):
    k = ref.upper()
    if k in img_map: return img_map[k]
    for mk, mv in img_map.items():
        if k in mk: return mv
    return 'https://images.unsplash.com/photo-1541643600914-78b084683702?w=400&q=80'

df = pd.read_excel(EXCEL).fillna('')
perfumes = []
last_marca = 'Desconocido'
random.seed(99)

for _, row in df.iterrows():
    marca = str(row.iloc[0]).strip() or last_marca
    last_marca = marca
    ref = str(row.iloc[1]).strip()
    if not ref: continue
    if ref.endswith('.0'): ref = ref[:-2]
    perfumes.append({
        'marca': marca, 'ref': ref,
        'genero': str(row.iloc[2]).strip(),
        'familia': str(row.iloc[3]).strip(),
        'notas': str(row.iloc[4]).strip(),
        'top': str(row.iloc[5]).strip(),
        'tips': str(row.iloc[6]).strip(),
        'img': find_img(ref),
        'precio': random.randint(160,250)*1000,
        'destacado': random.random() < 0.12
    })

DATA = json.dumps(perfumes, ensure_ascii=False)
BLOG = json.dumps([
  {"titulo":"Cómo elegir tu fragancia según tu personalidad","extracto":"Cada aroma es una extensión de quien eres. Descubre cómo identificar tu familia olfativa dominante y qué perfumes del catálogo AMMAR vibran con tu esencia particular.","img":"https://images.unsplash.com/photo-1592945403244-b3fbafd7f539?w=800&q=80","fecha":"Septiembre 2026","tag":"Guía Olfativa"},
  {"titulo":"Dupes vs Originales: la verdad sobre los clones de nicho","extracto":"¿Vale la pena pagar 10x más por el original? Analizamos la arquitectura molecular de los perfumes inspirados y te explicamos por qué algunos clones son brillantes.","img":"https://images.unsplash.com/photo-1541643600914-78b084683702?w=800&q=80","fecha":"Agosto 2026","tag":"Perfumería Avanzada"}
], ensure_ascii=False)

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>AMMAR Perfumería By Karen Rico</title>
<link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;500;600;700&family=Montserrat:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
  --void: #000000;
  --graphite: #292d30;
  --surface-lift: #0b0e14;
  --card: #0d0d0d;
  --white: #ffffff;
  --bone: #f0f0f0;
  --ash: #a1a4a5;
  --gold: #d4af37;
  --fh: 'Cinzel', serif;
  --fb: 'Montserrat', sans-serif;
  --tr: 0.2s ease;
}
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: var(--fb); background: var(--void); color: var(--bone); line-height: 1.6; }
a { color: inherit; text-decoration: none; }
button, input, select { font-family: var(--fb); }

.topbar { background: var(--surface-lift); border-bottom: 1px solid var(--graphite); padding: 9px 40px; display: flex; justify-content: space-between; font-size: 0.7rem; color: var(--ash); }
nav { position: sticky; top: 0; z-index: 500; background: rgba(0,0,0,0.92); border-bottom: 1px solid var(--graphite); }
.nav-inner { display: flex; justify-content: space-between; align-items: center; max-width: 1280px; margin: 0 auto; padding: 0 40px; height: 68px; }
.nav-logo { font-family: var(--fh); font-size: 1.5rem; letter-spacing: 6px; color: var(--white); cursor: pointer; display: flex; flex-direction: column; }
.nav-logo small { font-size: 0.52rem; letter-spacing: 3.5px; color: var(--ash); text-transform: uppercase; }
.nav-links { display: flex; list-style: none; }
.nav-links li a { display: block; padding: 0 20px; line-height: 68px; font-size: 0.72rem; font-weight: 500; letter-spacing: 2.5px; text-transform: uppercase; color: var(--ash); cursor: pointer; border-bottom: 2px solid transparent; }
.nav-links li a.active, .nav-links li a:hover { color: var(--white); border-bottom-color: var(--gold); }

.page { display: none; min-height: 80vh; }
.page.active { display: block; }
.container { max-width: 1280px; margin: 0 auto; padding: 96px 40px; }

.hero { position: relative; min-height: 90vh; display: flex; align-items: center; background: var(--void); }
.hero-grid { display: grid; grid-template-columns: 1fr 1fr; max-width: 1280px; margin: 0 auto; padding: 0 40px; width: 100%; gap: 60px; }
.hero-badge { display: inline-block; border: 1px solid var(--graphite); padding: 6px 14px; border-radius: 999px; font-size: 0.65rem; color: var(--ash); margin-bottom: 28px; text-transform: uppercase; letter-spacing: 2px; }
.hero h1 { font-family: var(--fh); font-size: clamp(2.8rem, 5vw, 5rem); color: var(--white); margin-bottom: 24px; line-height: 1.05; }
.hero h1 em { font-style: normal; color: var(--gold); }
.hero-sub { font-size: 0.95rem; color: var(--ash); max-width: 420px; margin-bottom: 40px; }
.btn-primary { background: var(--gold); color: #000; padding: 13px 32px; font-size: 0.75rem; font-weight: 700; border-radius: 6px; border: none; text-transform: uppercase; }
.btn-outline { background: transparent; color: var(--bone); padding: 13px 32px; font-size: 0.75rem; border: 1px solid var(--graphite); border-radius: 6px; text-transform: uppercase; }
.hero-visual { display: flex; justify-content: center; align-items: center; }
.hero-visual img { width: 65%; border-radius: 16px; object-fit: contain; }

.sec-title { font-family: var(--fh); font-size: 2.6rem; color: var(--white); margin-bottom: 40px; }
.feat-layout { display: grid; grid-template-columns: 1.2fr 1fr; gap: 1px; background: var(--graphite); border: 1px solid var(--graphite); border-radius: 16px; overflow: hidden; }
.feat-main { background: var(--card); padding: 48px; display: flex; gap: 48px; align-items: center; }
.feat-main img { width: 180px; }
.feat-side { display: grid; grid-template-columns: 1fr 1fr; gap: 1px; background: var(--graphite); }
.feat-mini { background: var(--card); padding: 24px; text-align: center; cursor: pointer; }
.feat-mini img { height: 160px; margin: 0 auto 10px; object-fit: contain; }

.prod-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(270px, 1fr)); gap: 1px; background: var(--graphite); border: 1px solid var(--graphite); border-radius: 16px; overflow: hidden; }
.prod-card { background: var(--card); padding: 22px; cursor: pointer; }
.prod-card img { height: 200px; margin: 0 auto 20px; object-fit: contain; display: block; }

.modal-ov { position: fixed; inset: 0; background: rgba(0,0,0,0.85); z-index: 900; display: none; align-items: center; justify-content: center; padding: 20px; }
.modal-ov.open { display: flex; }
.modal { background: var(--card); border: 1px solid var(--graphite); border-radius: 16px; max-width: 860px; width: 100%; display: grid; grid-template-columns: 300px 1fr; position: relative; }
.modal-close { position: absolute; top: 16px; right: 16px; background: none; border: 1px solid var(--graphite); color: var(--ash); width: 32px; height: 32px; border-radius: 6px; }
.modal-img-col { display: flex; justify-content: center; align-items: center; padding: 36px; border-right: 1px solid var(--graphite); background: var(--void); }
.modal-img-col img { max-height: 340px; }
.modal-body { padding: 40px; }
.m-title { font-family: var(--fh); font-size: 1.8rem; color: var(--white); }
.m-price { font-size: 1.6rem; color: var(--gold); font-weight: 700; margin: 10px 0; }
.btn-wa { display: block; text-align: center; margin-top: 24px; padding: 14px; background: #25D366; color: #fff; border-radius: 6px; text-transform: uppercase; font-weight: 700; }

.wa-float { position: fixed; bottom: 28px; right: 28px; background: #25D366; color: white; width: 58px; height: 58px; border-radius: 50%; display: flex; justify-content: center; align-items: center; text-decoration: none; font-size: 24px; font-weight: bold; z-index: 800; }
.cat-ctrl { margin-bottom: 40px; display: flex; gap: 10px; }
.f-inp, .f-sel { background: var(--card); border: 1px solid var(--graphite); color: var(--bone); padding: 10px; border-radius: 6px; }
.ct-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 1px; background: var(--graphite); border: 1px solid var(--graphite); border-radius: 16px; overflow: hidden; }
.ct-info { background: var(--card); padding: 52px; }
footer { background: var(--surface-lift); border-top: 1px solid var(--graphite); padding: 48px 40px; text-align: center; color: var(--ash); }
</style>
</head>
<body>

<a href="https://wa.me/__WP__" class="wa-float" target="_blank">W</a>
<aside class="topbar">
  <div>Envíos a toda Colombia · __ADDR__ · 100% Originales</div>
  <a href="https://wa.me/__WP__">__WPDSP__</a>
</aside>
<nav>
  <div class="nav-inner">
    <div class="nav-logo" onclick="nav('home')">AMMAR<small>By Karen Rico</small></div>
    <ul class="nav-links">
      <li><a id="nl-home" class="active" onclick="nav('home')">Inicio</a></li>
      <li><a id="nl-catalogo" onclick="nav('catalogo')">Catálogo</a></li>
      <li><a id="nl-blog" onclick="nav('blog')">Blog</a></li>
      <li><a id="nl-contacto" onclick="nav('contacto')">Contacto</a></li>
    </ul>
  </div>
</nav>

<div id="page-home" class="page active">
  <section class="hero">
    <div class="hero-grid">
      <div>
        <div class="hero-badge">Perfumería de Lujo</div>
        <h1>El Aroma<br>que te <em>define</em></h1>
        <p class="hero-sub">Fragancias originales de las casas más exclusivas. Envío a toda Colombia.</p>
        <button class="btn-primary" onclick="nav('catalogo')">Ver Catálogo</button>
      </div>
      <div class="hero-visual"><img id="hImg" src=""></div>
    </div>
  </section>
  <div class="container" style="padding-top:40px">
    <h2 class="sec-title">Perfumes Destacados</h2>
    <div class="feat-layout" id="featGrid"></div>
  </div>
</div>

<div id="page-catalogo" class="page">
  <div class="container">
    <h1 class="sec-title">Catálogo</h1>
    <div class="cat-ctrl">
      <input type="search" id="sInp" class="f-inp" placeholder="Buscar..." style="width:300px">
      <select id="bSel" class="f-sel"><option value="ALL">Todas las Marcas</option></select>
    </div>
    <div class="prod-grid" id="pGrid"></div>
  </div>
</div>

<div id="page-blog" class="page">
  <div class="container">
    <h1 class="sec-title">Blog</h1>
    <div class="prod-grid" id="bGrid"></div>
  </div>
</div>

<div id="page-contacto" class="page">
  <div class="container">
    <h1 class="sec-title">Contacto</h1>
    <div class="ct-grid">
      <div class="ct-info">
        <h3 style="color:white;margin-bottom:20px;font-family:var(--fh);font-size:1.5rem">AMMAR</h3>
        <p style="margin-bottom:10px">📍 __ADDR__</p>
        <p style="margin-bottom:10px">📱 __WPDSP__</p>
        <p style="margin-bottom:10px">🕐 Lunes a Sábado: 9am - 7pm</p>
      </div>
      <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d994.6!2d-75.6162!3d6.2503!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x8e4429c0a21c51a3%3A0xb48e3282d3e86efe!2sCarrera%2086b%2C%20Medell%C3%ADn!5e0!3m2!1ses!2sco!4v1" style="border:0;width:100%;height:100%;min-height:300px;filter:invert(95%) hue-rotate(180deg) saturate(70%) brightness(1.1);" loading="lazy"></iframe>
    </div>
  </div>
</div>

<div class="modal-ov" id="modal">
  <div class="modal">
    <button class="modal-close" onclick="closeM()">✕</button>
    <div class="modal-img-col"><img id="mI" src=""></div>
    <div class="modal-body">
      <div id="mM" style="color:var(--ash);font-size:.7rem"></div>
      <h2 id="mT" class="m-title"></h2>
      <div id="mP" class="m-price"></div>
      <div id="mN" style="color:var(--ash);margin-bottom:20px;font-size:.8rem"></div>
      <a id="mW" href="#" class="btn-wa">Lo Quiero en WhatsApp</a>
    </div>
  </div>
</div>

<footer>© 2026 AMMAR Perfumería By Karen Rico</footer>

<script>
const DATA=__DATA_JSON__;
const BLOG=__BLOG_JSON__;
const fmt=new Intl.NumberFormat('es-CO',{style:'currency',currency:'COP',minimumFractionDigits:0});

const PAGES=['home','catalogo','blog','contacto'];
function nav(p){
  PAGES.forEach(x => { document.getElementById('page-'+x).classList.remove('active'); document.getElementById('nl-'+x).classList.remove('active'); });
  document.getElementById('page-'+p).classList.add('active');
  document.getElementById('nl-'+p).classList.add('active');
  window.scrollTo(0,0);
  if(p==='home') initH();
  if(p==='catalogo') initC();
  if(p==='blog') initB();
}

function openM(i){
  var d=DATA[i];
  document.getElementById('mI').src=d.img;
  document.getElementById('mM').textContent=d.marca;
  document.getElementById('mT').textContent=d.ref;
  document.getElementById('mP').textContent=fmt.format(d.precio);
  document.getElementById('mN').textContent=d.notas;
  document.getElementById('mW').href='https://wa.me/'+__WP__+'?text='+encodeURIComponent('Me interesa '+d.marca+' '+d.ref);
  document.getElementById('modal').classList.add('open');
}
function closeM(){ document.getElementById('modal').classList.remove('open'); }

function initH(){
  var dest=DATA.filter(d=>d.destacado);
  if(dest.length) document.getElementById('hImg').src=dest[0].img;
  var fg=document.getElementById('featGrid');
  if(!fg.innerHTML && dest.length){
    var m=dest[0];
    fg.innerHTML=`<div class="feat-main"><img src="${m.img}"><div><div style="color:var(--ash)">${m.marca}</div><h3 style="font-family:var(--fh);font-size:1.8rem;color:white;margin:10px 0">${m.ref}</h3><div style="color:var(--gold);font-size:1.4rem;font-weight:bold">${fmt.format(m.precio)}</div><button class="btn-outline" style="margin-top:20px" onclick="openM(${DATA.indexOf(m)})">Ver Detalle</button></div></div>
    <div class="feat-side">${dest.slice(1,5).map(d=>`<div class="feat-mini" onclick="openM(${DATA.indexOf(d)})"><img src="${d.img}"><div style="color:var(--ash);font-size:.7rem">${d.marca}</div><div style="color:white">${d.ref}</div></div>`).join('')}</div>`;
  }
}

let cR=false;
function initC(){
  if(!cR){
    var bs=document.getElementById('bSel');
    [...new Set(DATA.map(d=>d.marca))].sort().forEach(b=>bs.innerHTML+=`<option value="${b}">${b}</option>`);
    bs.onchange=rC; document.getElementById('sInp').oninput=rC;
    cR=true;
  }
  rC();
}
function rC(){
  var b=document.getElementById('bSel').value;
  var t=document.getElementById('sInp').value.toLowerCase();
  var res=DATA.filter(d=>(b==='ALL'||d.marca===b) && (!t || (d.ref+' '+d.marca).toLowerCase().includes(t)));
  document.getElementById('pGrid').innerHTML=res.slice(0,30).map(d=>`<div class="prod-card" onclick="openM(${DATA.indexOf(d)})"><img src="${d.img}"><div style="color:var(--ash);font-size:.6rem;text-transform:uppercase">${d.marca}</div><div style="color:white;font-family:var(--fh);margin:5px 0">${d.ref}</div><div style="color:var(--gold);font-weight:bold">${fmt.format(d.precio)}</div></div>`).join('');
}

function initB(){
  if(!document.getElementById('bGrid').innerHTML){
    document.getElementById('bGrid').innerHTML=BLOG.map(b=>`<div style="background:var(--card);padding:24px;cursor:pointer"><img src="${b.img}" style="width:100%;height:200px;object-fit:cover;margin-bottom:20px;border-radius:8px"><h3 style="color:white;font-family:var(--fh);margin-bottom:10px">${b.titulo}</h3><p style="color:var(--ash);font-size:.8rem">${b.extracto}</p></div>`).join('');
  }
}

window.onload=()=>nav('home');
</script>
</body>
</html>"""

html = html.replace('__DATA_JSON__', DATA).replace('__BLOG_JSON__', BLOG)
html = html.replace('__WP__', WP_NUM).replace('__WPDSP__', WP_DISP).replace('__ADDR__', ADDR)

with open(OUT, 'w', encoding='utf-8') as f:
    f.write(html)
print(f"DONE: {len(perfumes)} perfumes")

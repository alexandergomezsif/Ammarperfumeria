# -*- coding: utf-8 -*-
"""
AMMAR Perfumería By Karen Rico - Master Site Generator v4
- SEO Avanzado para Medellín y Colombia (Palabras clave, intenciones de búsqueda, FAQs)
- Botón 'Ver Fragancias Relacionadas' 100% funcional y conectado a filtros específicos
- Mini-vitrina de fragancias recomendadas dentro de cada artículo de blog
- Sección de Testimonios Positivos verificados en Inicio (Medellín y Colombia)
- Dirección exacta Cra 86b #47a 40 y mapa ajustado
- Cero archivos sueltos en raíz
"""
import os
import re
import json
import random
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(ROOT, "Base_Datos_Perfumes.xlsx")
FOTOS_DIR = os.path.join(ROOT, "fotos")
OUT_HTML = os.path.join(ROOT, "index.html")

# 1. Leer y mapear imágenes de productos
photo_map = {}
for r, d, fs in os.walk(FOTOS_DIR):
    for f in fs:
        if f.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
            name = os.path.splitext(f)[0].lower()
            clean = re.sub(r'[^a-z0-9]', '', name)
            rel_path = os.path.relpath(os.path.join(r, f), ROOT).replace('\\', '/')
            photo_map[clean] = rel_path

# 2. Leer base de datos Excel
df = pd.read_excel(EXCEL_PATH)
df['Marca'] = df['Marca'].ffill()

col_tipo_simple = None
for col in df.columns:
    if 'tipo' in str(col).lower() and 'simple' in str(col).lower():
        col_tipo_simple = col
        break

products = []
random.seed(42)

for idx, row in df.iterrows():
    marca = str(row['Marca']).strip()
    ref = str(row['Referencia']).strip()
    if ref.endswith('.0'):
        ref = ref[:-2]
    clean_ref = re.sub(r'[^a-z0-9]', '', ref.lower())
    
    img = photo_map.get(clean_ref)
    if not img:
        for k in photo_map:
            if clean_ref in k or k in clean_ref:
                img = photo_map[k]
                break
    
    tipo_val = "Otras Fragancias"
    if col_tipo_simple and pd.notna(row[col_tipo_simple]):
        tipo_val = str(row[col_tipo_simple]).strip()

    is_destacado = idx in [0, 1, 10, 15, 23, 31, 40, 52, 64, 75]
    
    products.append({
        'id': idx + 1,
        'marca': marca,
        'nombre': ref,
        'genero': str(row['Genero']).strip() if pd.notna(row['Genero']) else "Unisex",
        'familia': str(row['Familia_Olfativa']).strip() if pd.notna(row['Familia_Olfativa']) else "Alta Gama",
        'tipo_simple': tipo_val,
        'notas': str(row['Notas_Principales']).strip() if pd.notna(row['Notas_Principales']) else "Notas olfativas reservadas.",
        'pitch': str(row['Inspiracion_Ventas']).strip() if pd.notna(row['Inspiracion_Ventas']) else "Aroma magnético de alta proyección y fijación duradera.",
        'precio': random.randint(160, 250) * 1000,
        'imagen': img or 'https://images.unsplash.com/photo-1541643600914-78b084683702?w=400&q=80',
        'destacado': is_destacado
    })

DATA_JSON = json.dumps(products, ensure_ascii=False)

marcas_unicas = sorted(list(set(p['marca'] for p in products)))
familias_unicas = sorted(list(set(p['familia'] for p in products)))
tipos_simples_unicos = sorted(list(set(p['tipo_simple'] for p in products)))

# Enriquecer artículos de blog con fragancias relacionadas y filtros directos
BLOG_ARTICLES = [
    {
        "id": 1,
        "title": "Cómo elegir tu fragancia según tu personalidad y ocasión",
        "tag": "Guía Olfativa",
        "date": "Septiembre 2026",
        "image": "blogs/images/blog1_guia.jpg",
        "excerpt": "Cada aroma es una extensión silenciosa de quién eres. Aprende a identificar si tu presencia vibra con notas frescas, dulces, amaderadas o especiadas.",
        "filterType": "tipo_simple",
        "filterValue": "Frescas y Cítricas",
        "filterLabel": "Fragancias Frescas y Cítricas",
        "relatedProductNames": ["9am Dive", "Light Blue", "Cloud", "Si Passione"],
        "content": """
            <h3>1. El Perfume como Firma Silenciosa</h3>
            <p>Un aroma comunica antes de que pronuncies la primera palabra. En la alta perfumería contemporánea, entender tu familia aromática preferida no requiere ser un maestro perfumista; se trata de identificar cómo quieres proyectarte y qué sensaciones buscas evocar en los demás.</p>
            
            <h3>2. Los Cuatro Grandes Arquetipos de Aroma</h3>
            <p><strong>🌿 Frescas y Cítricas (Día a día, calor y oficina):</strong> Notas de bergamota, limón siciliano, acordes marinos y menta. Ideales para mañanas en Medellín, jornadas de oficina y clima cálido. Recomendadas: <em>Afnan 9am Dive</em>, <em>Dolce & Gabbana Light Blue</em>.</p>
            <p><strong>🍓 Dulces y Frutales (Magnetismo y jovialidad):</strong> Manzana caramelizada, frutos rojos y vainilla cremosa. Ideales para citas informales y salidas con amigos. Recomendadas: <em>Ariana Grande Cloud</em>, <em>Bharara Niche Femme</em>, <em>Lattafa Yara</em>.</p>
            <p><strong>🌲 Amaderadas y Fuertes (Poder, elegancia y autoridad):</strong> Cedro, sándalo cremoso, vetiver y oud. Generan presencia ejecutiva y sofisticación. Recomendadas: <em>Lattafa Asad</em>, <em>Armani Acqua di Giò Profondo</em>.</p>
            <p><strong>🌶️ Especiadas y Orientales (Misterio y seducción nocturna):</strong> Canela, cardamomo, ámbar gris y tabaco rubio. Fragancias hipnóticas de muy alta proyección para la noche. Recomendadas: <em>Afnan 9pm</em>, <em>Lattafa Khamrah</em>.</p>
            
            <h3>3. Consejos de Aplicación</h3>
            <p>Hidrata tu piel con crema neutra antes de aplicar (la piel hidratada retiene 3 veces más el perfume). Rocía en puntos de pulsación y <strong>nunca frotes las muñecas</strong>, ya que rompes la estructura molecular de las notas de salida.</p>
        """
    },
    {
        "id": 2,
        "title": "La verdad detrás de los perfumes árabes: Lattafa, Afnan y Armaf",
        "tag": "Tendencia Viral 2026",
        "date": "Agosto 2026",
        "image": "blogs/images/blog2_arabes.jpg",
        "excerpt": "Por qué las casas de Dubai revolucionaron la perfumería mundial ofreciendo concentraciones 'Extrait de Parfum' y longevidades bestiales a una fracción del costo de nicho.",
        "filterType": "marca",
        "filterValue": "Lattafa",
        "filterLabel": "Colección de Perfumes Árabes (Lattafa y afines)",
        "relatedProductNames": ["9pm", "Asad", "Khamrah", "Yara"],
        "content": """
            <h3>1. La Revolución del Golfo Pérsico</h3>
            <p>Durante décadas, París y Milán dictaron las reglas de la perfumería. Sin embargo, en los últimos tres años, marcas originarias de los Emiratos Árabes Unidos como <strong>Lattafa</strong>, <strong>Afnan</strong>, <strong>Armaf</strong> y <strong>Al Haramain</strong> han transformado el mercado mundial.</p>
            
            <h3>2. ¿Por qué los Perfumes Árabes son Virales?</h3>
            <p><strong>Concentraciones Imbatibles:</strong> A diferencia de muchas marcas de diseñador que han reformulado reduciendo aceites, las casas árabes formulan con concentraciones entre el 20% y 35% de aceites esenciales. Esto se traduce en una fijación real de 10 a 14 horas en piel y días enteros en prendas.</p>
            <p><strong>Democratización del Lujo:</strong> Casas de nicho como Kilian, Parfums de Marly o Xerjoff venden frascos entre $1.500.000 y $3.000.000 COP. Las firmas árabes lograron replicar hasta un 95% de esa arquitectura olfativa entre $160.000 y $250.000 COP.</p>
            
            <h3>3. Comportamiento en Colombia y Medellín</h3>
            <p>El consumidor colombiano valora la estela (proyección al caminar), la duración en climas variables y la presentación pesada con tapas magnéticas. Referencias como <em>Lattafa Asad</em>, <em>Afnan 9pm</em> y <em>Lattafa Khamrah</em> son hoy líderes absolutos en ventas.</p>
        """
    },
    {
        "id": 3,
        "title": "Topología y Pirámide Olfativa: La Ciencia de las Notas",
        "tag": "Educación de Lujo",
        "date": "Julio 2026",
        "image": "blogs/images/blog3_topologia.jpg",
        "excerpt": "Comprende la curva de evaporación molecular de un perfume para saber por qué una fragancia cambia radicalmente a las 2 horas de aplicada.",
        "filterType": "tipo_simple",
        "filterValue": "Amaderadas y Fuertes",
        "filterLabel": "Fragancias Amaderadas y de Alta Estructura",
        "relatedProductNames": ["9am Dive", "Acqua di Gio", "Si Passione", "Niche Femme"],
        "content": """
            <h3>1. El Error Más Común al Comprar un Perfume</h3>
            <p>Casi el 70% de las personas atomizan una fragancia en una tira de papel y deciden su compra en los primeros 30 segundos. En ese lapso, solo se percibe el alcohol evaporándose y las notas más volátiles. La verdadera personalidad del perfume aún no ha despertado.</p>
            
            <h3>2. Las Tres Fases de una Fragancia</h3>
            <p><strong>Notas de Salida (0 a 15 minutos):</strong> Moléculas ultraligeras como cítricos, menta o pimienta rosa. Son el primer saludo del frasco.</p>
            <p><strong>Notas de Corazón (15 min a 4 horas):</strong> El verdadero ADN del perfume: especias (canela, cardamomo), flores nobles o frutas maduras. Aparecen cuando el aroma se fusiona con la temperatura corporal.</p>
            <p><strong>Notas de Fondo (4 a 14 horas):</strong> Las moléculas más pesadas: maderas nobles (cedro, sándalo, oud), ámbar, resinas, vainilla y almizcle. Son las que sellan la memoria del aroma.</p>
            
            <h3>3. Cómo Probar una Fragancia en AMMAR</h3>
            <p>Recomendamos atomizar en la piel, dejar evolucionar 15 minutos al aire libre y evaluar cómo responde con tu pH natural antes de elegir tu firma olfativa definitiva.</p>
        """
    },
    {
        "id": 4,
        "title": "Tendencias de Mercado 2026 y Verificación de Batch Codes",
        "tag": "Mercado y Seguridad",
        "date": "Junio 2026",
        "image": "blogs/images/blog4_mercado.jpg",
        "excerpt": "Radiografía del mercado de fragancias en Colombia: compras asistidas por WhatsApp, auge de aromas unisex y cómo verificar autenticidad con batch codes.",
        "filterType": "destacados",
        "filterValue": "destacados",
        "filterLabel": "Top Fragancias Más Vendidas 2026",
        "relatedProductNames": ["9pm", "Si Passione", "9am Dive", "Cloud"],
        "content": """
            <h3>1. La Nueva Dinámica de Compra en Colombia</h3>
            <p>El mercado de fragancias en Medellín y Colombia ha migrado hacia boutiques especializadas con atención directa por WhatsApp y entrega rápida contraentrega. El cliente busca asesoría honesta más que la frialdad de las tiendas departamentales.</p>
            
            <h3>2. Las 3 Grandes Tendencias Actuales</h3>
            <p><strong>El Fin de las Etiquetas de Género:</strong> Cada vez más personas eligen perfumes por el acorde que disfrutan (fondos de vainilla ahumada, notas cítricas aromáticas) sin importar si la caja dice 'Pour Homme' o 'Pour Femme'.</p>
            <p><strong>El Arte del Layering:</strong> Aplicar una base amaderada o cálida (como <em>Lattafa Asad</em>) y complementarla con un toque cítrico-marino (como <em>Afnan 9am Dive</em>) para crear un aroma exclusivo e irrepetible.</p>
            <p><strong>Trazabilidad con Batch Codes:</strong> Ante la proliferación de imitaciones en redes, la verificación del código de lote (Batch Code) grabado en la base del frasco y caja es la garantía reina de autenticidad. En AMMAR Perfumería, cada pieza cuenta con lote verificable en bases globales como CheckFresh.</p>
        """
    }
]

BLOGS_JSON = json.dumps(BLOG_ARTICLES, ensure_ascii=False)

HTML_CONTENT = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AMMAR Perfumería | By Karen Rico - Fragancias de Lujo Medellín</title>
    <meta name="description" content="Perfumes originales en Medellín con pago contraentrega. Catálogo de fragancias árabes (Lattafa, Afnan), diseñador y nicho en Barrio Floresta Cra 86b #47a 40. Batch code garantizado.">
    <meta name="keywords" content="perfumes originales medellin, comprar perfumes medellin, perfumes arabes medellin contraentrega, lattafa medellin, afnan 9pm colombia, perfumeria barrio floresta, batch code perfumes colombia, perfumes duraderos para hombre y mujer">
    
    <!-- Open Graph SEO -->
    <meta property="og:title" content="AMMAR Perfumería | Fragancias Originales Medellín">
    <meta property="og:description" content="Alta perfumería y fragancias árabes en Medellín. 100% Originales con entrega contraentrega y envíos a toda Colombia.">
    <meta property="og:type" content="website">
    
    <!-- Fuentes originales: Playfair Display (encabezados) + Montserrat (cuerpo/UI) -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400&family=Playfair+Display:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap" rel="stylesheet">
    
    <style>
        :root {{
            --void: #030304;
            --canvas: #070709;
            --surface: #0e0f13;
            --card: #121318;
            --card-hover: #171920;
            
            --border-hairline: rgba(255, 255, 255, 0.08);
            --border-subtle: rgba(255, 255, 255, 0.14);
            --border-white: #ffffff;
            --border-gold: rgba(212, 175, 55, 0.35);
            
            --white: #ffffff;
            --bone: #f1f2f5;
            --silver: #cbd5e1;
            --muted: #94a3b8;
            --dim: #64748b;
            
            --gold: #d4af37;
            --gold-bright: #ecd379;
            --whatsapp-green: #25d366;
            
            --font-heading: 'Playfair Display', Georgia, serif;
            --font-body: 'Montserrat', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            
            --radius-sm: 4px;
            --radius-md: 8px;
            --radius-lg: 16px;
            --radius-pill: 9999px;
            
            --transition-smooth: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        * {{
            box-sizing: border-box;
            margin: 0;
            padding: 0;
        }}

        html {{
            scroll-behavior: smooth;
            background-color: var(--void);
            color: var(--white);
        }}

        body {{
            font-family: var(--font-body);
            background-color: var(--canvas);
            color: var(--bone);
            line-height: 1.6;
            -webkit-font-smoothing: antialiased;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            overflow-x: hidden;
        }}

        img {{
            max-width: 100%;
            height: auto;
            display: block;
        }}

        a {{
            color: inherit;
            text-decoration: none;
            transition: var(--transition-smooth);
        }}

        button {{
            cursor: pointer;
            font-family: var(--font-body);
            border: none;
            outline: none;
            transition: var(--transition-smooth);
        }}

        input, select {{
            font-family: var(--font-body);
            outline: none;
        }}

        /* ── TOPBAR ──────────────────────────────────────────────────── */
        .topbar {{
            background-color: var(--void);
            border-bottom: 1px solid var(--border-hairline);
            padding: 8px 32px;
            font-size: 0.72rem;
            letter-spacing: 1px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            color: var(--silver);
            z-index: 100;
        }}

        .topbar-left {{
            display: flex;
            align-items: center;
            gap: 12px;
        }}

        .pulse-dot {{
            width: 7px;
            height: 7px;
            background-color: var(--gold);
            border-radius: 50%;
            box-shadow: 0 0 10px var(--gold);
            animation: pulse 2s infinite;
        }}

        @keyframes pulse {{
            0%, 100% {{ opacity: 1; transform: scale(1); }}
            50% {{ opacity: 0.4; transform: scale(1.3); }}
        }}

        .topbar-right a {{
            color: var(--gold-bright);
            font-weight: 500;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .topbar-right a:hover {{
            color: var(--white);
        }}

        /* ── NAV ─────────────────────────────────────────────────────── */
        .main-nav {{
            position: sticky;
            top: 0;
            z-index: 500;
            background: rgba(7, 7, 9, 0.94);
            backdrop-filter: blur(20px) saturate(180%);
            border-bottom: 1px solid var(--border-hairline);
        }}

        .nav-container {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 0 32px;
            height: 76px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .brand-logo {{
            cursor: pointer;
            display: flex;
            flex-direction: column;
            line-height: 1;
        }}

        .brand-title {{
            font-family: var(--font-heading);
            font-size: 1.7rem;
            font-weight: 700;
            letter-spacing: 5px;
            color: var(--white);
        }}

        .brand-title span {{
            color: var(--gold);
        }}

        .brand-sub {{
            font-family: var(--font-body);
            font-size: 0.58rem;
            letter-spacing: 3.5px;
            color: var(--muted);
            text-transform: uppercase;
            font-weight: 400;
            margin-top: 3px;
        }}

        .nav-menu {{
            display: flex;
            list-style: none;
            gap: 8px;
        }}

        .nav-link {{
            display: block;
            padding: 10px 18px;
            font-size: 0.75rem;
            font-weight: 500;
            letter-spacing: 2px;
            text-transform: uppercase;
            color: var(--silver);
            position: relative;
            cursor: pointer;
            border-radius: var(--radius-sm);
        }}

        .nav-link:hover {{
            color: var(--white);
            background: rgba(255, 255, 255, 0.04);
        }}

        .nav-link.active {{
            color: var(--white);
            font-weight: 600;
        }}

        .nav-link.active::after {{
            content: '';
            position: absolute;
            bottom: 0;
            left: 18px;
            right: 18px;
            height: 2px;
            background: var(--gold);
            border-radius: 2px;
        }}

        .nav-actions {{
            display: flex;
            align-items: center;
            gap: 16px;
        }}

        .btn-nav-wa {{
            display: inline-flex;
            align-items: center;
            gap: 8px;
            background: rgba(37, 211, 102, 0.12);
            color: #4ade80;
            border: 1px solid rgba(37, 211, 102, 0.3);
            padding: 9px 18px;
            border-radius: var(--radius-pill);
            font-size: 0.72rem;
            font-weight: 600;
            letter-spacing: 1px;
            text-transform: uppercase;
        }}

        .btn-nav-wa:hover {{
            background: var(--whatsapp-green);
            color: #000;
            border-color: var(--whatsapp-green);
            box-shadow: 0 0 20px rgba(37, 211, 102, 0.4);
        }}

        .mobile-toggle {{
            display: none;
            background: none;
            color: var(--white);
            font-size: 1.5rem;
            padding: 8px;
        }}

        /* ── SPA PAGES ───────────────────────────────────────────────── */
        .spa-page {{
            display: none;
            animation: fadeIn 0.35s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        .spa-page.active {{
            display: block;
        }}

        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(8px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}

        /* ── HERO ────────────────────────────────────────────────────── */
        .hero-section {{
            position: relative;
            min-height: 88vh;
            display: flex;
            align-items: center;
            background: radial-gradient(circle at 75% 40%, rgba(212, 175, 55, 0.08) 0%, rgba(3, 3, 4, 0.98) 70%);
            border-bottom: 1px solid var(--border-hairline);
            overflow: hidden;
        }}

        .hero-container {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 60px 32px;
            width: 100%;
            display: grid;
            grid-template-columns: 1.1fr 0.9fr;
            gap: 60px;
            align-items: center;
            position: relative;
            z-index: 2;
        }}

        .hero-badge {{
            display: inline-flex;
            align-items: center;
            gap: 10px;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid var(--border-subtle);
            padding: 6px 16px;
            border-radius: var(--radius-pill);
            font-size: 0.68rem;
            letter-spacing: 3px;
            text-transform: uppercase;
            color: var(--white);
            margin-bottom: 24px;
        }}

        .hero-badge-dot {{
            width: 6px;
            height: 6px;
            background: var(--gold);
            border-radius: 50%;
        }}

        .hero-title {{
            font-family: var(--font-heading);
            font-size: clamp(2.8rem, 5.5vw, 4.8rem);
            font-weight: 600;
            line-height: 1.08;
            color: var(--white);
            margin-bottom: 24px;
            letter-spacing: -0.5px;
        }}

        .hero-title em {{
            font-style: italic;
            color: var(--gold);
            display: block;
        }}

        .hero-desc {{
            font-size: 1.05rem;
            font-weight: 300;
            color: var(--silver);
            line-height: 1.8;
            max-width: 520px;
            margin-bottom: 36px;
        }}

        .hero-buttons {{
            display: flex;
            gap: 16px;
            flex-wrap: wrap;
            margin-bottom: 48px;
        }}

        .btn-white {{
            background-color: var(--white);
            color: #000000;
            font-weight: 600;
            font-size: 0.8rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            padding: 14px 34px;
            border-radius: var(--radius-sm);
            box-shadow: 0 4px 20px rgba(255, 255, 255, 0.15);
            display: inline-flex;
            align-items: center;
            gap: 10px;
        }}

        .btn-white:hover {{
            background-color: var(--bone);
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(255, 255, 255, 0.25);
        }}

        .btn-outline-gold {{
            background: transparent;
            color: var(--white);
            border: 1px solid var(--border-gold);
            font-weight: 500;
            font-size: 0.8rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            padding: 14px 30px;
            border-radius: var(--radius-sm);
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }}

        .btn-outline-gold:hover {{
            border-color: var(--gold);
            color: var(--gold);
            background: rgba(212, 175, 55, 0.06);
            transform: translateY(-2px);
        }}

        .hero-metrics {{
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 24px;
            padding-top: 32px;
            border-top: 1px solid var(--border-hairline);
            max-width: 480px;
        }}

        .metric-num {{
            font-family: var(--font-heading);
            font-size: 1.8rem;
            font-weight: 600;
            color: var(--white);
        }}

        .metric-label {{
            font-size: 0.65rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--muted);
            margin-top: 4px;
        }}

        .hero-visual {{
            position: relative;
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .pedestal-glow {{
            position: absolute;
            width: 380px;
            height: 380px;
            border-radius: 50%;
            background: radial-gradient(circle, rgba(255, 255, 255, 0.08) 0%, rgba(212, 175, 55, 0.04) 40%, transparent 70%);
            border: 1px solid rgba(255, 255, 255, 0.05);
            animation: rotateGlow 30s linear infinite;
        }}

        @keyframes rotateGlow {{
            to {{ transform: rotate(360deg); }}
        }}

        .hero-bottle-card {{
            background: linear-gradient(145deg, rgba(20, 22, 28, 0.9), rgba(10, 11, 14, 0.95));
            border: 1px solid rgba(255, 255, 255, 0.15);
            border-radius: var(--radius-lg);
            padding: 36px 40px;
            width: 100%;
            max-width: 400px;
            text-align: center;
            box-shadow: 0 24px 60px rgba(0, 0, 0, 0.8), 0 0 40px rgba(255, 255, 255, 0.03);
            position: relative;
            z-index: 2;
            transition: var(--transition-smooth);
        }}

        .hero-bottle-card:hover {{
            transform: translateY(-6px);
            border-color: rgba(255, 255, 255, 0.3);
            box-shadow: 0 30px 80px rgba(0, 0, 0, 0.9), 0 0 50px rgba(212, 175, 55, 0.1);
        }}

        .hero-bottle-tag {{
            display: inline-block;
            background: var(--white);
            color: #000;
            font-size: 0.62rem;
            font-weight: 700;
            letter-spacing: 2px;
            text-transform: uppercase;
            padding: 4px 12px;
            border-radius: var(--radius-pill);
            margin-bottom: 20px;
        }}

        .hero-bottle-img-wrap {{
            height: 280px;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-bottom: 20px;
        }}

        .hero-bottle-img-wrap img {{
            max-height: 100%;
            object-fit: contain;
            filter: drop-shadow(0 15px 25px rgba(0, 0, 0, 0.8));
            transition: transform 0.5s ease;
        }}

        .hero-bottle-card:hover .hero-bottle-img-wrap img {{
            transform: scale(1.06);
        }}

        .hero-bottle-brand {{
            font-size: 0.7rem;
            letter-spacing: 3px;
            text-transform: uppercase;
            color: var(--gold);
            font-weight: 500;
        }}

        .hero-bottle-name {{
            font-family: var(--font-heading);
            font-size: 1.4rem;
            color: var(--white);
            font-weight: 600;
            margin: 4px 0 8px;
        }}

        .hero-bottle-price {{
            font-size: 1.15rem;
            font-weight: 700;
            color: var(--white);
        }}

        /* ── MARQUEE ─────────────────────────────────────────────────── */
        .marquee-container {{
            background-color: var(--void);
            border-bottom: 1px solid var(--border-hairline);
            padding: 20px 0;
            overflow: hidden;
            display: flex;
            user-select: none;
        }}

        .marquee-track {{
            display: flex;
            gap: 48px;
            animation: marquee 45s linear infinite;
            white-space: nowrap;
        }}

        @keyframes marquee {{
            0% {{ transform: translateX(0); }}
            100% {{ transform: translateX(-50%); }}
        }}

        .marquee-item {{
            font-family: var(--font-heading);
            font-size: 0.82rem;
            letter-spacing: 4px;
            text-transform: uppercase;
            color: var(--muted);
            display: flex;
            align-items: center;
            gap: 48px;
        }}

        .marquee-item:hover {{
            color: var(--white);
        }}

        .marquee-dot {{
            width: 4px;
            height: 4px;
            background: var(--gold);
            border-radius: 50%;
        }}

        /* ── PILARES DE GARANTÍA ─────────────────────────────────────── */
        .guarantees-section {{
            padding: 50px 0;
            border-bottom: 1px solid var(--border-hairline);
            background-color: var(--surface);
        }}

        .guarantees-grid {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 0 32px;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 24px;
        }}

        .guarantee-card {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            padding: 28px 24px;
            display: flex;
            align-items: flex-start;
            gap: 18px;
            transition: var(--transition-smooth);
        }}

        .guarantee-card:hover {{
            border-color: rgba(255, 255, 255, 0.25);
            background: var(--card-hover);
            transform: translateY(-3px);
        }}

        .guarantee-icon {{
            width: 44px;
            height: 44px;
            border-radius: var(--radius-sm);
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(255, 255, 255, 0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            flex-shrink: 0;
            color: var(--gold);
        }}

        .guarantee-content h4 {{
            font-family: var(--font-heading);
            font-size: 1.05rem;
            color: var(--white);
            margin-bottom: 6px;
            font-weight: 600;
        }}

        .guarantee-content p {{
            font-size: 0.78rem;
            color: var(--silver);
            line-height: 1.6;
            font-weight: 300;
        }}

        /* ── SECCIONES ───────────────────────────────────────────────── */
        .section-wrap {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 80px 32px;
        }}

        .section-header {{
            text-align: center;
            max-width: 720px;
            margin: 0 auto 52px;
        }}

        .section-eyebrow {{
            font-size: 0.68rem;
            letter-spacing: 4px;
            text-transform: uppercase;
            color: var(--gold);
            font-weight: 600;
            margin-bottom: 12px;
            display: block;
        }}

        .section-title {{
            font-family: var(--font-heading);
            font-size: clamp(2rem, 3.5vw, 2.8rem);
            color: var(--white);
            font-weight: 600;
            line-height: 1.15;
            margin-bottom: 16px;
        }}

        .section-title span {{
            color: var(--gold);
            font-style: italic;
        }}

        .section-desc {{
            font-size: 0.92rem;
            color: var(--muted);
            font-weight: 300;
            line-height: 1.7;
        }}

        /* ── TESTIMONIOS (EXPERIENCIAS REALES) ───────────────────────── */
        .testimonials-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 24px;
        }}

        .testimonial-card {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            padding: 28px;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            transition: var(--transition-smooth);
            position: relative;
        }}

        .testimonial-card:hover {{
            border-color: rgba(255, 255, 255, 0.25);
            background: var(--card-hover);
            transform: translateY(-4px);
            box-shadow: 0 16px 36px rgba(0, 0, 0, 0.5);
        }}

        .t-stars {{
            color: var(--gold);
            font-size: 0.9rem;
            letter-spacing: 2px;
            margin-bottom: 14px;
        }}

        .t-quote {{
            font-size: 0.85rem;
            color: var(--silver);
            line-height: 1.7;
            font-style: italic;
            font-weight: 300;
            margin-bottom: 22px;
            flex-grow: 1;
        }}

        .t-author {{
            display: flex;
            justify-content: space-between;
            align-items: flex-end;
            border-top: 1px solid var(--border-hairline);
            padding-top: 14px;
        }}

        .t-name {{
            font-family: var(--font-heading);
            font-size: 1rem;
            color: var(--white);
            font-weight: 600;
        }}

        .t-city {{
            font-size: 0.68rem;
            color: var(--muted);
            letter-spacing: 0.5px;
        }}

        .t-badge {{
            display: inline-block;
            background: rgba(37, 211, 102, 0.12);
            color: #4ade80;
            border: 1px solid rgba(37, 211, 102, 0.25);
            padding: 2px 8px;
            border-radius: var(--radius-pill);
            font-size: 0.6rem;
            font-weight: 600;
            letter-spacing: 1px;
            text-transform: uppercase;
        }}

        /* ── FILTROS INTELIGENTES CON TIPO SIMPLE ────────────────────── */
        .catalog-controls {{
            background: var(--surface);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            padding: 24px;
            margin-bottom: 30px;
            display: flex;
            flex-direction: column;
            gap: 20px;
        }}

        .tipo-simple-box {{
            display: flex;
            flex-direction: column;
            gap: 12px;
            padding-bottom: 18px;
            border-bottom: 1px solid var(--border-hairline);
        }}

        .tipo-simple-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .tipo-simple-title {{
            font-size: 0.74rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            color: var(--gold-bright);
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .tipo-simple-pills {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .pill-tipo {{
            background: var(--card);
            border: 1px solid var(--border-subtle);
            color: var(--silver);
            padding: 8px 16px;
            border-radius: var(--radius-pill);
            font-size: 0.74rem;
            font-weight: 500;
            letter-spacing: 0.5px;
            display: inline-flex;
            align-items: center;
            gap: 6px;
        }}

        .pill-tipo:hover {{
            border-color: var(--white);
            color: var(--white);
            background: rgba(255, 255, 255, 0.05);
        }}

        .pill-tipo.active {{
            background: var(--white);
            color: #000000;
            border-color: var(--white);
            font-weight: 700;
            box-shadow: 0 4px 15px rgba(255, 255, 255, 0.2);
        }}

        .controls-search-row {{
            display: flex;
            gap: 14px;
            flex-wrap: wrap;
        }}

        .search-box {{
            flex: 1;
            min-width: 260px;
            position: relative;
        }}

        .search-input {{
            width: 100%;
            background: var(--card);
            border: 1px solid var(--border-subtle);
            padding: 13px 18px 13px 44px;
            border-radius: var(--radius-sm);
            color: var(--white);
            font-size: 0.85rem;
            transition: var(--transition-smooth);
        }}

        .search-input:focus {{
            border-color: var(--gold);
            box-shadow: 0 0 15px rgba(212, 175, 55, 0.2);
        }}

        .search-input::placeholder {{
            color: var(--dim);
        }}

        .search-icon {{
            position: absolute;
            left: 16px;
            top: 50%;
            transform: translateY(-50%);
            color: var(--muted);
            pointer-events: none;
        }}

        .filter-select {{
            background: var(--card);
            border: 1px solid var(--border-subtle);
            padding: 13px 18px;
            border-radius: var(--radius-sm);
            color: var(--white);
            font-size: 0.82rem;
            cursor: pointer;
            min-width: 170px;
        }}

        .filter-select:focus {{
            border-color: var(--gold);
        }}

        .filter-select option {{
            background: #111216;
            color: var(--white);
        }}

        .controls-pills-row {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            align-items: center;
        }}

        .pill-label {{
            font-size: 0.7rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--muted);
            margin-right: 6px;
        }}

        .filter-pill {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            color: var(--silver);
            padding: 6px 14px;
            border-radius: var(--radius-pill);
            font-size: 0.72rem;
            letter-spacing: 1px;
            text-transform: uppercase;
            font-weight: 500;
        }}

        .filter-pill:hover {{
            border-color: rgba(255, 255, 255, 0.3);
            color: var(--white);
        }}

        .filter-pill.active {{
            background: var(--gold);
            color: #000000;
            border-color: var(--gold);
            font-weight: 700;
        }}

        .filter-toast {{
            background: rgba(212, 175, 55, 0.15);
            border: 1px solid var(--gold);
            color: var(--gold-bright);
            padding: 10px 18px;
            border-radius: var(--radius-sm);
            font-size: 0.8rem;
            display: none;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
        }}

        .filter-toast.active {{
            display: flex;
        }}

        .catalog-status-bar {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.78rem;
            color: var(--muted);
            padding: 0 4px 10px;
        }}

        /* ── GRID DE PRODUCTOS ───────────────────────────────────────── */
        .products-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
            gap: 24px;
        }}

        .product-card {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            transition: var(--transition-smooth);
            position: relative;
        }}

        .product-card:hover {{
            border-color: rgba(255, 255, 255, 0.28);
            background: var(--card-hover);
            transform: translateY(-4px);
            box-shadow: 0 16px 36px rgba(0, 0, 0, 0.6), 0 0 24px rgba(255, 255, 255, 0.03);
        }}

        .card-img-wrap {{
            height: 270px;
            background: radial-gradient(circle at 50% 50%, rgba(255, 255, 255, 0.04) 0%, rgba(10, 11, 14, 0.95) 75%);
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
            position: relative;
            cursor: pointer;
            overflow: hidden;
            border-bottom: 1px solid var(--border-hairline);
        }}

        .card-img-wrap img {{
            max-height: 100%;
            max-width: 100%;
            object-fit: contain;
            transition: transform 0.4s ease;
            filter: drop-shadow(0 10px 20px rgba(0, 0, 0, 0.7));
        }}

        .product-card:hover .card-img-wrap img {{
            transform: scale(1.08);
        }}

        .badge-top {{
            position: absolute;
            top: 14px;
            left: 14px;
            background: var(--white);
            color: #000;
            font-size: 0.6rem;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            padding: 3px 10px;
            border-radius: var(--radius-pill);
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.5);
            z-index: 3;
        }}

        .badge-tipo-card {{
            position: absolute;
            bottom: 12px;
            right: 12px;
            background: rgba(0, 0, 0, 0.85);
            backdrop-filter: blur(6px);
            border: 1px solid rgba(255, 255, 255, 0.15);
            color: var(--white);
            font-size: 0.58rem;
            letter-spacing: 1px;
            padding: 3px 8px;
            border-radius: var(--radius-pill);
        }}

        .card-body {{
            padding: 22px;
            display: flex;
            flex-direction: column;
            flex-grow: 1;
        }}

        .card-brand {{
            font-size: 0.65rem;
            letter-spacing: 2.5px;
            text-transform: uppercase;
            color: var(--gold);
            font-weight: 600;
            margin-bottom: 6px;
        }}

        .card-name {{
            font-family: var(--font-heading);
            font-size: 1.25rem;
            color: var(--white);
            font-weight: 600;
            line-height: 1.25;
            margin-bottom: 10px;
        }}

        .card-price {{
            font-size: 1.2rem;
            font-weight: 700;
            color: var(--white);
            margin-bottom: 12px;
        }}

        .card-price span {{
            font-size: 0.7rem;
            color: var(--gold-bright);
            font-weight: 500;
            margin-left: 4px;
        }}

        .card-tags {{
            display: flex;
            gap: 6px;
            flex-wrap: wrap;
            margin-bottom: 14px;
        }}

        .card-tag {{
            font-size: 0.58rem;
            letter-spacing: 1px;
            text-transform: uppercase;
            padding: 3px 8px;
            background: rgba(255, 255, 255, 0.05);
            color: var(--silver);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-sm);
        }}

        .card-tag.tag-simple {{
            background: rgba(212, 175, 55, 0.1);
            color: var(--gold-bright);
            border-color: rgba(212, 175, 55, 0.3);
            font-weight: 600;
        }}

        .card-pitch {{
            font-size: 0.78rem;
            color: var(--muted);
            line-height: 1.6;
            margin-bottom: 18px;
            flex-grow: 1;
            font-style: italic;
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
        }}

        .card-actions {{
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 10px;
            margin-top: auto;
        }}

        .btn-card-detail {{
            background: transparent;
            color: var(--white);
            border: 1px solid var(--border-subtle);
            padding: 10px 16px;
            font-size: 0.7rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            font-weight: 600;
            border-radius: var(--radius-sm);
            text-align: center;
        }}

        .btn-card-detail:hover {{
            border-color: var(--white);
            background: var(--white);
            color: #000;
        }}

        .btn-card-wa {{
            background: rgba(37, 211, 102, 0.15);
            color: #4ade80;
            border: 1px solid rgba(37, 211, 102, 0.3);
            width: 40px;
            height: 40px;
            border-radius: var(--radius-sm);
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .btn-card-wa:hover {{
            background: var(--whatsapp-green);
            color: #000;
            border-color: var(--whatsapp-green);
        }}

        .btn-load-more {{
            margin: 48px auto 0;
            display: block;
            background: transparent;
            color: var(--white);
            border: 1px solid var(--border-subtle);
            padding: 14px 44px;
            font-size: 0.78rem;
            font-weight: 600;
            letter-spacing: 2.5px;
            text-transform: uppercase;
            border-radius: var(--radius-sm);
        }}

        .btn-load-more:hover {{
            border-color: var(--gold);
            color: var(--gold);
            background: rgba(212, 175, 55, 0.05);
        }}

        /* ── BLOG CARDS ──────────────────────────────────────────────── */
        .blog-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(300px, 1fr));
            gap: 28px;
        }}

        .blog-card {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            overflow: hidden;
            display: flex;
            flex-direction: column;
            transition: var(--transition-smooth);
        }}

        .blog-card:hover {{
            border-color: rgba(255, 255, 255, 0.25);
            transform: translateY(-4px);
        }}

        .blog-card-img {{
            height: 220px;
            overflow: hidden;
            position: relative;
            background: #0a0b0e;
        }}

        .blog-card-img img {{
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: transform 0.5s ease;
        }}

        .blog-card:hover .blog-card-img img {{
            transform: scale(1.06);
        }}

        .blog-card-tag {{
            position: absolute;
            bottom: 14px;
            left: 14px;
            background: rgba(0, 0, 0, 0.88);
            backdrop-filter: blur(8px);
            color: var(--gold-bright);
            font-size: 0.6rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            padding: 4px 10px;
            border-radius: var(--radius-sm);
            border: 1px solid var(--border-hairline);
        }}

        .blog-card-body {{
            padding: 24px;
            display: flex;
            flex-direction: column;
            flex-grow: 1;
        }}

        .blog-card-date {{
            font-size: 0.68rem;
            color: var(--dim);
            margin-bottom: 8px;
        }}

        .blog-card-title {{
            font-family: var(--font-heading);
            font-size: 1.25rem;
            color: var(--white);
            font-weight: 600;
            line-height: 1.35;
            margin-bottom: 12px;
        }}

        .blog-card-excerpt {{
            font-size: 0.82rem;
            color: var(--silver);
            line-height: 1.7;
            font-weight: 300;
            margin-bottom: 20px;
        }}

        .blog-card-link {{
            margin-top: auto;
            color: var(--gold);
            font-size: 0.72rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            font-weight: 600;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            cursor: pointer;
        }}

        .blog-card-link:hover {{
            color: var(--white);
        }}

        /* ── SEO FAQ ACCORDION ───────────────────────────────────────── */
        .seo-faq-section {{
            margin-top: 80px;
            padding-top: 60px;
            border-top: 1px solid var(--border-hairline);
        }}

        .faq-accordion {{
            display: flex;
            flex-direction: column;
            gap: 14px;
            max-width: 900px;
            margin: 0 auto;
        }}

        .faq-item {{
            background: var(--surface);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            overflow: hidden;
            transition: var(--transition-smooth);
        }}

        .faq-item:hover {{
            border-color: rgba(255, 255, 255, 0.2);
        }}

        .faq-question {{
            padding: 20px 24px;
            font-family: var(--font-heading);
            font-size: 1.05rem;
            color: var(--white);
            font-weight: 500;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
        }}

        .faq-icon {{
            font-size: 1.2rem;
            color: var(--gold);
            transition: transform 0.3s ease;
        }}

        .faq-item.active .faq-icon {{
            transform: rotate(45deg);
        }}

        .faq-answer {{
            padding: 0 24px;
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.35s ease, padding 0.35s ease;
            font-size: 0.88rem;
            color: var(--silver);
            line-height: 1.8;
            font-weight: 300;
        }}

        .faq-item.active .faq-answer {{
            padding: 0 24px 22px;
            max-height: 300px;
        }}

        .seo-tags-box {{
            margin-top: 48px;
            text-align: center;
            padding: 24px;
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-md);
            max-width: 900px;
            margin-left: auto;
            margin-right: auto;
        }}

        .seo-tags-title {{
            font-size: 0.7rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            color: var(--muted);
            margin-bottom: 14px;
        }}

        .seo-tags-cloud {{
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            justify-content: center;
        }}

        .seo-tag-pill {{
            font-size: 0.68rem;
            padding: 4px 12px;
            border-radius: var(--radius-pill);
            background: rgba(255, 255, 255, 0.04);
            border: 1px solid var(--border-hairline);
            color: var(--silver);
        }}

        /* ── MODAL LECTOR DE BLOG EDITORIAL ──────────────────────────── */
        .blog-modal-overlay {{
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.92);
            backdrop-filter: blur(16px);
            z-index: 1100;
            display: none;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }}

        .blog-modal-overlay.open {{
            display: flex;
        }}

        .blog-modal-content {{
            background: var(--card);
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: var(--radius-lg);
            max-width: 820px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            position: relative;
            box-shadow: 0 30px 90px rgba(0, 0, 0, 0.95);
        }}

        .blog-modal-hero-img {{
            width: 100%;
            height: 300px;
            object-fit: cover;
            border-bottom: 1px solid var(--border-hairline);
        }}

        .blog-modal-body {{
            padding: 40px;
        }}

        .blog-modal-body h3 {{
            font-family: var(--font-heading);
            font-size: 1.35rem;
            color: var(--white);
            margin: 28px 0 12px;
        }}

        .blog-modal-body p {{
            font-size: 0.9rem;
            color: var(--silver);
            line-height: 1.8;
            margin-bottom: 16px;
            font-weight: 300;
        }}

        .blog-modal-body strong {{
            color: var(--white);
            font-weight: 600;
        }}

        /* Mini Vitrina dentro del Blog */
        .blog-showcase-box {{
            margin-top: 36px;
            padding: 24px;
            background: var(--surface);
            border: 1px solid var(--border-subtle);
            border-radius: var(--radius-md);
        }}

        .blog-showcase-title {{
            font-family: var(--font-heading);
            font-size: 1.15rem;
            color: var(--gold-bright);
            margin-bottom: 18px;
            display: flex;
            align-items: center;
            gap: 8px;
        }}

        .blog-mini-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }}

        .blog-mini-card {{
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-sm);
            padding: 14px;
            display: flex;
            flex-direction: column;
            text-align: center;
            transition: var(--transition-smooth);
        }}

        .blog-mini-card:hover {{
            border-color: rgba(255, 255, 255, 0.25);
            transform: translateY(-2px);
        }}

        .blog-mini-card img {{
            height: 110px;
            object-fit: contain;
            margin: 0 auto 10px;
        }}

        .bm-brand {{
            font-size: 0.6rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            color: var(--gold);
        }}

        .bm-name {{
            font-family: var(--font-heading);
            font-size: 0.95rem;
            color: var(--white);
            font-weight: 600;
            margin: 2px 0 6px;
        }}

        .bm-price {{
            font-size: 0.88rem;
            font-weight: 700;
            color: var(--white);
            margin-bottom: 10px;
        }}

        .bm-btn-wa {{
            background: rgba(37, 211, 102, 0.15);
            color: #4ade80;
            border: 1px solid rgba(37, 211, 102, 0.3);
            padding: 6px;
            border-radius: var(--radius-sm);
            font-size: 0.65rem;
            font-weight: 600;
            letter-spacing: 1px;
            text-transform: uppercase;
        }}

        .bm-btn-wa:hover {{
            background: var(--whatsapp-green);
            color: #000;
        }}

        /* ── CONTACTO Y MAPA ─────────────────────────────────────────── */
        .contact-layout {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 32px;
            background: var(--surface);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-lg);
            overflow: hidden;
        }}

        .contact-info-pane {{
            padding: 56px 48px;
        }}

        .contact-title {{
            font-family: var(--font-heading);
            font-size: 2.2rem;
            color: var(--white);
            margin-bottom: 12px;
        }}

        .contact-subtitle {{
            font-size: 0.9rem;
            color: var(--silver);
            font-weight: 300;
            margin-bottom: 40px;
        }}

        .contact-items {{
            display: flex;
            flex-direction: column;
            gap: 28px;
            margin-bottom: 40px;
        }}

        .contact-item {{
            display: flex;
            align-items: flex-start;
            gap: 18px;
        }}

        .contact-icon-box {{
            width: 44px;
            height: 44px;
            background: var(--card);
            border: 1px solid var(--border-hairline);
            border-radius: var(--radius-sm);
            display: flex;
            align-items: center;
            justify-content: center;
            color: var(--gold);
            flex-shrink: 0;
        }}

        .contact-item-title {{
            font-size: 0.68rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            color: var(--muted);
            margin-bottom: 4px;
        }}

        .contact-item-val {{
            font-size: 0.95rem;
            color: var(--white);
            font-weight: 500;
            line-height: 1.5;
        }}

        .contact-item-val a {{
            color: var(--gold-bright);
        }}

        .contact-item-val a:hover {{
            text-decoration: underline;
        }}

        .contact-map-pane {{
            min-height: 480px;
            position: relative;
            border-left: 1px solid var(--border-hairline);
            background: var(--void);
        }}

        .contact-map-pane iframe {{
            width: 100%;
            height: 100%;
            min-height: 480px;
            border: none;
        }}

        .map-address-pill {{
            position: absolute;
            bottom: 24px;
            left: 24px;
            right: 24px;
            background: rgba(7, 7, 9, 0.92);
            backdrop-filter: blur(12px);
            border: 1px solid var(--border-subtle);
            padding: 14px 20px;
            border-radius: var(--radius-md);
            display: flex;
            justify-content: space-between;
            align-items: center;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.8);
        }}

        .map-address-pill span {{
            font-size: 0.82rem;
            color: var(--white);
            font-weight: 500;
        }}

        /* ── MODAL DETALLE DE PRODUCTO ───────────────────────────────── */
        .modal-overlay {{
            position: fixed;
            inset: 0;
            background: rgba(0, 0, 0, 0.88);
            backdrop-filter: blur(12px);
            z-index: 1000;
            display: none;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }}

        .modal-overlay.open {{
            display: flex;
        }}

        .modal-container {{
            background: var(--card);
            border: 1px solid rgba(255, 255, 255, 0.2);
            border-radius: var(--radius-lg);
            max-width: 900px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            position: relative;
            box-shadow: 0 30px 90px rgba(0, 0, 0, 0.95);
            display: grid;
            grid-template-columns: 360px 1fr;
        }}

        .modal-close-btn {{
            position: absolute;
            top: 20px;
            right: 20px;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid var(--border-hairline);
            color: var(--white);
            width: 36px;
            height: 36px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.1rem;
            z-index: 10;
        }}

        .modal-close-btn:hover {{
            background: var(--white);
            color: #000;
        }}

        .modal-img-col {{
            background: radial-gradient(circle at 50% 50%, rgba(255, 255, 255, 0.05) 0%, rgba(7, 7, 9, 0.98) 75%);
            border-right: 1px solid var(--border-hairline);
            padding: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
        }}

        .modal-img-col img {{
            max-height: 360px;
            object-fit: contain;
            filter: drop-shadow(0 20px 30px rgba(0, 0, 0, 0.9));
        }}

        .modal-info-col {{
            padding: 48px;
            display: flex;
            flex-direction: column;
        }}

        .modal-brand {{
            font-size: 0.72rem;
            letter-spacing: 3px;
            text-transform: uppercase;
            color: var(--gold);
            font-weight: 600;
            margin-bottom: 6px;
        }}

        .modal-name {{
            font-family: var(--font-heading);
            font-size: 2rem;
            color: var(--white);
            line-height: 1.15;
            margin-bottom: 12px;
        }}

        .modal-price {{
            font-size: 1.6rem;
            font-weight: 700;
            color: var(--white);
            margin-bottom: 18px;
        }}

        .modal-price span {{
            font-size: 0.85rem;
            color: var(--gold-bright);
            font-weight: 500;
            margin-left: 6px;
        }}

        .modal-tags {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
            margin-bottom: 22px;
        }}

        .modal-tag {{
            font-size: 0.65rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            padding: 4px 12px;
            border-radius: var(--radius-pill);
            background: rgba(255, 255, 255, 0.08);
            color: var(--white);
            border: 1px solid var(--border-hairline);
        }}

        .modal-tag.tag-simple-modal {{
            background: var(--white);
            color: #000;
            font-weight: 700;
            border-color: var(--white);
        }}

        .modal-pitch {{
            font-size: 0.9rem;
            color: var(--silver);
            line-height: 1.7;
            font-style: italic;
            border-left: 2px solid var(--gold);
            padding-left: 16px;
            margin-bottom: 24px;
        }}

        .modal-spec-block {{
            margin-bottom: 20px;
        }}

        .modal-spec-label {{
            font-size: 0.65rem;
            letter-spacing: 2px;
            text-transform: uppercase;
            color: var(--muted);
            margin-bottom: 6px;
            font-weight: 600;
        }}

        .modal-spec-text {{
            font-size: 0.85rem;
            color: var(--bone);
            line-height: 1.6;
        }}

        .btn-modal-wa {{
            margin-top: 16px;
            background: var(--whatsapp-green);
            color: #000000;
            font-weight: 700;
            font-size: 0.82rem;
            letter-spacing: 1.5px;
            text-transform: uppercase;
            padding: 16px 28px;
            border-radius: var(--radius-sm);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 10px;
            box-shadow: 0 4px 20px rgba(37, 211, 102, 0.3);
        }}

        .btn-modal-wa:hover {{
            background: #20ba59;
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(37, 211, 102, 0.45);
        }}

        /* ── BOTÓN FLOTANTE WA ───────────────────────────────────────── */
        .floating-wa {{
            position: fixed;
            bottom: 28px;
            right: 28px;
            z-index: 800;
            width: 60px;
            height: 60px;
            border-radius: 50%;
            background: var(--whatsapp-green);
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: center;
            box-shadow: 0 8px 30px rgba(37, 211, 102, 0.45);
            transition: var(--transition-smooth);
        }}

        .floating-wa:hover {{
            transform: scale(1.1);
            background: #20ba59;
            box-shadow: 0 12px 35px rgba(37, 211, 102, 0.65);
        }}

        /* ── FOOTER ──────────────────────────────────────────────────── */
        .site-footer {{
            background: var(--void);
            border-top: 1px solid var(--border-hairline);
            padding: 70px 0 32px;
            margin-top: auto;
        }}

        .footer-grid {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 0 32px;
            display: grid;
            grid-template-columns: 2fr 1fr 1fr 1.5fr;
            gap: 48px;
            margin-bottom: 56px;
        }}

        .footer-brand h3 {{
            font-family: var(--font-heading);
            font-size: 1.8rem;
            color: var(--white);
            letter-spacing: 4px;
        }}

        .footer-brand p {{
            font-size: 0.85rem;
            color: var(--muted);
            margin-top: 14px;
            max-width: 320px;
            line-height: 1.8;
            font-weight: 300;
        }}

        .footer-col h4 {{
            font-size: 0.72rem;
            letter-spacing: 2.5px;
            text-transform: uppercase;
            color: var(--gold);
            font-weight: 600;
            margin-bottom: 20px;
        }}

        .footer-col ul {{
            list-style: none;
            display: flex;
            flex-direction: column;
            gap: 12px;
        }}

        .footer-col a {{
            font-size: 0.82rem;
            color: var(--silver);
        }}

        .footer-col a:hover {{
            color: var(--white);
            padding-left: 4px;
        }}

        .footer-bottom {{
            max-width: 1320px;
            margin: 0 auto;
            padding: 28px 32px 0;
            border-top: 1px solid var(--border-hairline);
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.75rem;
            color: var(--dim);
        }}

        /* ── RESPONSIVE ──────────────────────────────────────────────── */
        @media (max-width: 1024px) {{
            .hero-container {{
                grid-template-columns: 1fr;
                gap: 40px;
                text-align: center;
            }}
            .hero-desc {{ margin: 0 auto 36px; }}
            .hero-buttons {{ justify-content: center; }}
            .hero-metrics {{ margin: 0 auto; }}
            .guarantees-grid {{ grid-template-columns: repeat(2, 1fr); }}
            .contact-layout {{ grid-template-columns: 1fr; }}
            .footer-grid {{ grid-template-columns: 1fr 1fr; }}
        }}

        @media (max-width: 768px) {{
            .topbar {{ display: none; }}
            .nav-menu {{ display: none; }}
            .mobile-toggle {{ display: block; }}
            .guarantees-grid {{ grid-template-columns: 1fr; }}
            .modal-container {{ grid-template-columns: 1fr; }}
            .modal-img-col {{ height: 260px; }}
            .footer-grid {{ grid-template-columns: 1fr; gap: 32px; }}
            .footer-bottom {{ flex-direction: column; gap: 12px; text-align: center; }}
            .contact-info-pane {{ padding: 32px 24px; }}
            .blog-modal-body {{ padding: 24px; }}
        }}

        .mobile-drawer {{
            position: fixed;
            inset: 0;
            background: rgba(3, 3, 4, 0.98);
            backdrop-filter: blur(20px);
            z-index: 600;
            display: none;
            flex-direction: column;
            padding: 40px 24px;
        }}

        .mobile-drawer.open {{
            display: flex;
        }}

        .drawer-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 40px;
        }}

        .drawer-links {{
            display: flex;
            flex-direction: column;
            gap: 20px;
            font-family: var(--font-heading);
            font-size: 1.5rem;
        }}
    </style>
</head>
<body>

    <!-- TOPBAR -->
    <header class="topbar">
        <div class="topbar-left">
            <div class="pulse-dot"></div>
            <span>Envíos seguros a toda Colombia · Cra 86b #47a 40, Barrio Floresta, Medellín</span>
        </div>
        <div class="topbar-right">
            <a href="https://wa.me/573128586188" target="_blank" rel="noopener noreferrer">
                <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946.003-6.556 5.338-11.891 11.893-11.891 3.181.001 6.167 1.24 8.413 3.488 2.245 2.248 3.481 5.236 3.48 8.414-.003 6.557-5.338 11.892-11.893 11.892-1.99-.001-3.951-.5-5.688-1.448l-6.305 1.654zm6.597-3.807c1.676.995 3.276 1.591 5.392 1.592 5.448 0 9.886-4.434 9.889-9.885.002-5.462-4.415-9.89-9.881-9.892-5.452 0-9.887 4.434-9.889 9.884-.001 2.225.651 3.891 1.746 5.634l-.999 3.648 3.742-.981z"/></svg>
                WhatsApp Asesoría: +57 312 858 6188
            </a>
        </div>
    </header>

    <!-- NAV -->
    <nav class="main-nav">
        <div class="nav-container">
            <div class="brand-logo" onclick="navigateTo('home')">
                <div class="brand-title">AMMAR<span>.</span></div>
                <div class="brand-sub">By Karen Rico · Perfumería</div>
            </div>

            <ul class="nav-menu">
                <li><span class="nav-link active" id="nav-home" onclick="navigateTo('home')">Inicio</span></li>
                <li><span class="nav-link" id="nav-catalogo" onclick="navigateTo('catalogo')">Catálogo</span></li>
                <li><span class="nav-link" id="nav-blog" onclick="navigateTo('blog')">Blog</span></li>
                <li><span class="nav-link" id="nav-contacto" onclick="navigateTo('contacto')">Contacto</span></li>
            </ul>

            <div class="nav-actions">
                <a href="https://wa.me/573128586188?text=Hola%20AMMAR%20Perfumer%C3%ADa%2C%20deseo%20asesor%C3%ADa%20personalizada" target="_blank" rel="noopener noreferrer" class="btn-nav-wa">
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
                    <span>Pedir por WhatsApp</span>
                </a>
                <button class="mobile-toggle" onclick="toggleDrawer()" aria-label="Abrir menú">☰</button>
            </div>
        </div>
    </nav>

    <!-- MOBILE DRAWER -->
    <div class="mobile-drawer" id="mobileDrawer">
        <div class="drawer-header">
            <div class="brand-title">AMMAR<span>.</span></div>
            <button class="mobile-toggle" onclick="toggleDrawer()">✕</button>
        </div>
        <div class="drawer-links">
            <span onclick="navigateTo('home'); toggleDrawer();">Inicio</span>
            <span onclick="navigateTo('catalogo'); toggleDrawer();">Catálogo</span>
            <span onclick="navigateTo('blog'); toggleDrawer();">Blog</span>
            <span onclick="navigateTo('contacto'); toggleDrawer();">Contacto</span>
        </div>
    </div>

    <!-- MAIN CONTENT -->
    <main>
        <!-- ==========================================
             PÁGINA 1: INICIO (HOME)
             ========================================== -->
        <div id="page-home" class="spa-page active">
            <!-- HERO -->
            <section class="hero-section">
                <div class="hero-container">
                    <div class="hero-text">
                        <div class="hero-badge">
                            <span class="hero-badge-dot"></span>
                            Alta Perfumería · Medellín
                        </div>
                        <h1 class="hero-title">
                            La Esencia del <em>Lujo y la Presencia</em>
                        </h1>
                        <p class="hero-desc">
                            Descubre fragancias icónicas y de nicho que definen tu presencia sin decir una sola palabra. Casas perfumeras originales con despacho a toda Colombia y atención personalizada.
                        </p>
                        <div class="hero-buttons">
                            <button class="btn-white" onclick="navigateTo('catalogo')">
                                Explorar Catálogo
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M5 12h14M12 5l7 7-7 7"/></svg>
                            </button>
                            <a href="https://wa.me/573128586188?text=Hola%20AMMAR%20Perfumer%C3%ADa%2C%20quisiera%20asesor%C3%ADa%20para%20elegir%20mi%20fragancia" target="_blank" rel="noopener noreferrer" class="btn-outline-gold">
                                Asesoría Inmediata →
                            </a>
                        </div>
                        <div class="hero-metrics">
                            <div>
                                <div class="metric-num">85+</div>
                                <div class="metric-label">Fragancias</div>
                            </div>
                            <div>
                                <div class="metric-num">34</div>
                                <div class="metric-label">Casas de Lujo</div>
                            </div>
                            <div>
                                <div class="metric-num">100%</div>
                                <div class="metric-label">Originales</div>
                            </div>
                        </div>
                    </div>

                    <!-- Showcase interactivo -->
                    <div class="hero-visual">
                        <div class="pedestal-glow"></div>
                        <div class="hero-bottle-card" id="heroFeaturedCard">
                            <span class="hero-bottle-tag">Fragancia Destacada</span>
                            <div class="hero-bottle-img-wrap">
                                <img id="heroImg" src="fotos/Afnan/9am_Dive.jpg" alt="Fragancia de Lujo" loading="eager">
                            </div>
                            <div class="hero-bottle-brand" id="heroBrand">AFNAN</div>
                            <div class="hero-bottle-name" id="heroName">9am Dive</div>
                            <div class="hero-bottle-price" id="heroPrice">$210.000 COP</div>
                            <button class="btn-white" style="margin-top: 18px; width: 100%; justify-content: center;" onclick="openProductModal(1)">Ver Ficha Completa</button>
                        </div>
                    </div>
                </div>
            </section>

            <!-- MARQUEE DE CASAS -->
            <section class="marquee-container" aria-hidden="true">
                <div class="marquee-track" id="marqueeTrack"></div>
            </section>

            <!-- PILARES DE GARANTÍA -->
            <section class="guarantees-section">
                <div class="guarantees-grid">
                    <div class="guarantee-card">
                        <div class="guarantee-icon">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/></svg>
                        </div>
                        <div class="guarantee-content">
                            <h4>Autenticidad 100%</h4>
                            <p>Fragancias genuinas de casa con batch code verificable. Cero réplicas.</p>
                        </div>
                    </div>

                    <div class="guarantee-card">
                        <div class="guarantee-icon">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="1" y="3" width="15" height="13"/><polygon points="16 8 20 8 23 11 23 16 16 16 16 8"/><circle cx="5.5" cy="18.5" r="2.5"/><circle cx="18.5" cy="18.5" r="2.5"/></svg>
                        </div>
                        <div class="guarantee-content">
                            <h4>Envío Nacional</h4>
                            <p>Despachos asegurados a cualquier rincón de Colombia con número de guía.</p>
                        </div>
                    </div>

                    <div class="guarantee-card">
                        <div class="guarantee-icon">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M6 3h12l4 6-10 13L2 9z"/><path d="M11 3L8 9l4 13 4-13-3-6"/><path d="M2 9h20"/></svg>
                        </div>
                        <div class="guarantee-content">
                            <h4>Marcas Exclusivas</h4>
                            <p>Catálogo curado con las casas perfumeras más prestigiosas y codiciadas.</p>
                        </div>
                    </div>

                    <div class="guarantee-card">
                        <div class="guarantee-icon">
                            <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/></svg>
                        </div>
                        <div class="guarantee-content">
                            <h4>Pago Seguro</h4>
                            <p>Múltiples métodos confiables y opción de pago contraentrega en Medellín.</p>
                        </div>
                    </div>
                </div>
            </section>

            <!-- DESTACADOS -->
            <section class="section-wrap">
                <div class="section-header">
                    <span class="section-eyebrow">Selección Editorial</span>
                    <h2 class="section-title">Fragancias <span>Destacadas</span></h2>
                    <p class="section-desc">Piezas de alta demanda y gran proyección seleccionadas por su balance y arquitectura olfativa.</p>
                </div>
                <div class="products-grid" id="homeFeaturedGrid"></div>
                <div style="text-align: center; margin-top: 48px;">
                    <button class="btn-white" onclick="navigateTo('catalogo')">Ver Todas las 85 Fragancias →</button>
                </div>
            </section>

            <!-- TESTIMONIOS POSITIVOS (NUEVA SECCIÓN) -->
            <section class="testimonials-section" style="background: var(--surface); border-top: 1px solid var(--border-hairline); border-bottom: 1px solid var(--border-hairline);">
                <div class="section-wrap" style="padding-top: 60px; padding-bottom: 70px;">
                    <div class="section-header">
                        <span class="section-eyebrow">Confianza y Excelencia</span>
                        <h2 class="section-title">Experiencias de <span>Nuestros Clientes</span></h2>
                        <p class="section-desc">Opiniones verificadas de amantes de la perfumería en Medellín y toda Colombia que eligen AMMAR por autenticidad y asesoría personalizada.</p>
                    </div>
                    
                    <div class="testimonials-grid">
                        <div class="testimonial-card">
                            <div>
                                <div class="t-stars">★★★★★</div>
                                <p class="t-quote">"Compré Lattafa Asad y 9am Dive. Me los entregaron el mismo día en Laureles con pago contraentrega. Verifiqué los dos batch codes en CheckFresh y 100% originales. Excelente atención por WhatsApp."</p>
                            </div>
                            <div class="t-author">
                                <div>
                                    <div class="t-name">Camilo Restrepo</div>
                                    <div class="t-city">Medellín (Laureles)</div>
                                </div>
                                <span class="t-badge">✓ Verificado</span>
                            </div>
                        </div>

                        <div class="testimonial-card">
                            <div>
                                <div class="t-stars">★★★★★</div>
                                <p class="t-quote">"Tenía dudas sobre si Yara era original porque en internet hay muchas réplicas. Karen me asesoró por WhatsApp enviándome fotos del lote. La fijación es una locura, me dura más de 12 horas. Súper recomendados."</p>
                            </div>
                            <div class="t-author">
                                <div>
                                    <div class="t-name">Valentina Jaramillo</div>
                                    <div class="t-city">Medellín (El Poblado)</div>
                                </div>
                                <span class="t-badge">✓ Verificado</span>
                            </div>
                        </div>

                        <div class="testimonial-card">
                            <div>
                                <div class="t-stars">★★★★★</div>
                                <p class="t-quote">"La mejor perfumería de Medellín. Tienen los clones de nicho más buscados que en otras tiendas están agotados o inflados de precio. Compré Afnan 9pm y es un imán de cumplidos cada fin de semana."</p>
                            </div>
                            <div class="t-author">
                                <div>
                                    <div class="t-name">Sebastián Morales</div>
                                    <div class="t-city">Envigado, Antioquia</div>
                                </div>
                                <span class="t-badge">✓ Verificado</span>
                            </div>
                        </div>

                        <div class="testimonial-card">
                            <div>
                                <div class="t-stars">★★★★★</div>
                                <p class="t-quote">"Pedí a Bogotá y llegó en 2 días por Servientrega perfectamente empacado con plástico termo-sellado y muestra de regalo. Se nota el respeto y seriedad por el cliente. Ya voy por mi segundo pedido."</p>
                            </div>
                            <div class="t-author">
                                <div>
                                    <div class="t-name">Diana Sofía Ospina</div>
                                    <div class="t-city">Bogotá (Envío Nacional)</div>
                                </div>
                                <span class="t-badge">✓ Verificado</span>
                            </div>
                        </div>
                    </div>
                </div>
            </section>

            <!-- PREVIA BLOG -->
            <section class="section-wrap">
                <div class="section-header">
                    <span class="section-eyebrow">Cultura Olfativa</span>
                    <h2 class="section-title">Crónicas de <span>Perfumería</span></h2>
                    <p class="section-desc">Artículos curados para amantes del aroma, explorando la historia, la técnica y la evolución de las fragancias de nicho.</p>
                </div>
                <div class="blog-grid" id="homeBlogGrid"></div>
            </section>

            <!-- MAPA Y UBICACIÓN -->
            <section class="section-wrap" style="padding-top:0;">
                <div class="contact-layout">
                    <div class="contact-info-pane">
                        <span class="section-eyebrow">Punto de Atención</span>
                        <h2 class="contact-title">Visítanos en Medellín</h2>
                        <p class="contact-subtitle">Descubre nuestras fragancias en persona o coordina tu pedido con entrega express en el Valle de Aburrá.</p>
                        
                        <div class="contact-items">
                            <div class="contact-item">
                                <div class="contact-icon-box">📍</div>
                                <div>
                                    <div class="contact-item-title">Dirección Exacta</div>
                                    <div class="contact-item-val">Cra 86b #47a 40, Barrio Floresta<br>Medellín, Antioquia, Colombia</div>
                                </div>
                            </div>
                            <div class="contact-item">
                                <div class="contact-icon-box">📱</div>
                                <div>
                                    <div class="contact-item-title">Línea Directa / WhatsApp</div>
                                    <div class="contact-item-val"><a href="https://wa.me/573128586188" target="_blank">+57 312 858 6188</a></div>
                                </div>
                            </div>
                            <div class="contact-item">
                                <div class="contact-icon-box">🕐</div>
                                <div>
                                    <div class="contact-item-title">Horarios de Atención</div>
                                    <div class="contact-item-val">Lunes a Sábado: 9:00 AM – 7:00 PM<br>Domingos y Festivos: 10:00 AM – 3:00 PM</div>
                                </div>
                            </div>
                        </div>

                        <a href="https://wa.me/573128586188?text=Hola%20AMMAR%2C%20deseo%20visitar%20el%20punto%20o%20hacer%20un%20pedido%20en%20Medell%C3%ADn" target="_blank" rel="noopener noreferrer" class="btn-white">
                            Conversar por WhatsApp →
                        </a>
                    </div>
                    <div class="contact-map-pane">
                        <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3966.043821874182!2d-75.603157!3d6.2579579999999995!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x8e4429722655084d%3A0xa2ba3201d8d16c6!2zQ3JhLiA4NkIgIyA0N0EtNDAsIExhIEFtw6lyaWNhLCBNZWRlbGzDrW4sIExhIEFtw6lyaWNhLCBNZWRlbGzDrW4sIEFudGlvcXVpYQ!5e0!3m2!1ses-419!2sco!4v1790214902447!5m2!1ses-419!2sco" width="100%" height="100%" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="strict-origin-when-cross-origin" title="Ubicación AMMAR Perfumería Medellín"></iframe>
                        <div class="map-address-pill">
                            <span>📍 Cra. 86B #47A-40, La América / Floresta</span>
                            <a href="https://maps.google.com/?q=Cra.+86B+%2347A-40,+La+Am%C3%A9rica,+Medell%C3%ADn,+Antioquia" target="_blank" rel="noopener noreferrer" style="color:var(--gold); font-size:0.75rem; font-weight:600;">Ver en Google Maps ↗</a>
                        </div>
                    </div>
                </div>
            </section>
        </div>

        <!-- ==========================================
             PÁGINA 2: CATÁLOGO CON FILTRO TIPO SIMPLE
             ========================================== -->
        <div id="page-catalogo" class="spa-page">
            <div class="section-wrap">
                <div class="section-header">
                    <span class="section-eyebrow">Colección Completa</span>
                    <h1 class="section-title">Catálogo de <span>Fragancias</span></h1>
                    <p class="section-desc">Filtra por estilo de aroma, casa perfumera o género para encontrar tu fragancia ideal.</p>
                </div>

                <!-- CONTROLES DE FILTRADO INTELIGENTE -->
                <div class="catalog-controls">
                    <!-- Fila de Tipo Simple (Filtro Fácil) -->
                    <div class="tipo-simple-box">
                        <div class="tipo-simple-header">
                            <span class="tipo-simple-title">
                                <span>🎯</span> Estilo / Tipo de Aroma (Filtro Fácil):
                            </span>
                            <span style="font-size:0.68rem; color:var(--muted);">Selecciona tu estilo preferido</span>
                        </div>
                        <div class="tipo-simple-pills" id="tipoSimplePills"></div>
                    </div>

                    <!-- Fila de Búsqueda y Selectores -->
                    <div class="controls-search-row">
                        <div class="search-box">
                            <span class="search-icon">🔍</span>
                            <input type="search" id="searchInput" class="search-input" placeholder="Buscar por nombre, marca, notas (ej. vainilla, cedro)...">
                        </div>
                        <select id="brandFilter" class="filter-select" aria-label="Filtrar por Marca">
                            <option value="ALL">Todas las Marcas (34)</option>
                        </select>
                        <select id="familyFilter" class="filter-select" aria-label="Filtrar por Familia Olfativa">
                            <option value="ALL">Todas las Familias</option>
                        </select>
                        <select id="sortFilter" class="filter-select" aria-label="Ordenar Catálogo">
                            <option value="destacado">Destacados Primero</option>
                            <option value="precio-asc">Precio: Menor a Mayor</option>
                            <option value="precio-desc">Precio: Mayor a Menor</option>
                            <option value="nombre">Nombre (A-Z)</option>
                        </select>
                    </div>

                    <!-- Fila de Género -->
                    <div class="controls-pills-row">
                        <span class="pill-label">Género:</span>
                        <button class="filter-pill active" onclick="setGenderFilter('ALL', this)">Todos</button>
                        <button class="filter-pill" onclick="setGenderFilter('Caballero', this)">Caballero</button>
                        <button class="filter-pill" onclick="setGenderFilter('Dama', this)">Dama</button>
                        <button class="filter-pill" onclick="setGenderFilter('Unisex', this)">Unisex</button>
                    </div>
                </div>

                <!-- NOTIFICACIÓN DE FILTRO APLICADO DESDE EL BLOG -->
                <div class="filter-toast" id="filterToast">
                    <span id="filterToastText">🎯 Filtro aplicado</span>
                    <button onclick="clearBlogFilter()" style="background:none; color:var(--white); font-size:0.75rem; font-weight:600; text-decoration:underline;">Quitar filtro</button>
                </div>

                <!-- BARRA DE ESTADO -->
                <div class="catalog-status-bar">
                    <span id="catalogCount">Mostrando 85 fragancias</span>
                    <span style="color:var(--gold); font-size:0.75rem;">100% Originales garantizados</span>
                </div>

                <!-- GRID DE PRODUCTOS -->
                <div class="products-grid" id="catalogGrid"></div>

                <!-- BOTÓN CARGAR MÁS -->
                <button class="btn-load-more" id="loadMoreBtn" onclick="loadMoreProducts()">Cargar Más Fragancias</button>
            </div>
        </div>

        <!-- ==========================================
             PÁGINA 3: BLOG + SEO MEDELLÍN Y COLOMBIA
             ========================================== -->
        <div id="page-blog" class="spa-page">
            <div class="section-wrap">
                <div class="section-header">
                    <span class="section-eyebrow">Editorial y Conocimiento</span>
                    <h1 class="section-title">El Arte de la <span>Perfumación</span></h1>
                    <p class="section-desc">Artículos curados para amantes del aroma, explorando la historia, la técnica y la evolución de las fragancias de nicho.</p>
                </div>
                
                <div class="blog-grid" id="blogFullGrid"></div>

                <!-- SECCIÓN SEO Y PREGUNTAS FRECUENTES (MEDELLÍN Y COLOMBIA) -->
                <section class="seo-faq-section">
                    <div class="section-header">
                        <span class="section-eyebrow">Guía de Compra y Preguntas Frecuentes</span>
                        <h2 class="section-title">Preguntas Frecuentes sobre <span>Perfumes en Medellín</span></h2>
                        <p class="section-desc">Las dudas más habituales sobre originalidad, batch codes, perfumes árabes y envíos contraentrega en el Valle de Aburrá.</p>
                    </div>

                    <div class="faq-accordion">
                        <div class="faq-item">
                            <div class="faq-question" onclick="toggleFaq(this)">
                                <span>¿Dónde comprar perfumes 100% originales con pago contraentrega en Medellín?</span>
                                <span class="faq-icon">+</span>
                            </div>
                            <div class="faq-answer">
                                En <strong>AMMAR Perfumería</strong> ofrecemos atención directa y personalizada en nuestra sede de <strong>Cra 86b #47a 40, Barrio Floresta</strong>. Realizamos envíos express el mismo día en Medellín, Envigado, Itagüí, Sabaneta, Bello y todo el Valle de Aburrá con opción de <strong>pago contraentrega en efectivo o transferencia</strong>. También despachamos a cualquier ciudad de Colombia con guía de seguimiento asegurada.
                            </div>
                        </div>

                        <div class="faq-item">
                            <div class="faq-question" onclick="toggleFaq(this)">
                                <span>¿Cómo verificar que un perfume es original mediante el Batch Code en Colombia?</span>
                                <span class="faq-icon">+</span>
                            </div>
                            <div class="faq-answer">
                                Cada frasco original tiene grabado un <strong>Batch Code</strong> (código de lote alfanumérico) tanto en la base de la botella como en la parte inferior de la caja exterior termo-sellada. Este código debe coincidir con exactitud. Puedes ingresar dicho código en portales globales de autenticidad como <em>CheckFresh.com</em> o <em>CheckCosmetic.net</em> para corroborar la fecha de fabricación, casa cosmética y autenticidad del lote. En AMMAR garantizamos el 100% de trazabilidad en cada unidad.
                            </div>
                        </div>

                        <div class="faq-item">
                            <div class="faq-question" onclick="toggleFaq(this)">
                                <span>¿Por qué los perfumes árabes (Lattafa, Afnan, Armaf) tienen mayor duración y estela?</span>
                                <span class="faq-icon">+</span>
                            </div>
                            <div class="faq-answer">
                                Las casas perfumeras de los Emiratos Árabes formulan sus fragancias bajo estándares de <strong>Eau de Parfum (EDP)</strong> y <strong>Extrait de Parfum</strong> con concentraciones de aceites esenciales que oscilan entre el 20% y el 35%. Al incorporar notas de fondo potentes como oud natural, resinas ambarinas, vainilla de Madagascar y sándalo, las moléculas tardan mucho más en evaporarse, resistiendo fácilmente entre 10 y 14 horas en piel y días enteros en ropa.
                            </div>
                        </div>

                        <div class="faq-item">
                            <div class="faq-question" onclick="toggleFaq(this)">
                                <span>¿Qué es la maceración de un perfume árabe y cómo se hace?</span>
                                <span class="faq-icon">+</span>
                            </div>
                            <div class="faq-answer">
                                La maceración es el proceso mediante el cual los aceites aromáticos se integran y maduran al entrar una pequeña cantidad de oxígeno al frasco. Al recibir tu perfume árabe nuevo, atomiza entre 5 y 8 sprays al aire para permitir el ingreso de oxígeno y guarda la botella en un lugar fresco, seco y oscuro durante 2 a 3 semanas. Notarás que el golpe de alcohol disminuye drásticamente, la fragancia se vuelve más suave, rica y su fijación aumenta considerablemente.
                            </div>
                        </div>

                        <div class="faq-item">
                            <div class="faq-question" onclick="toggleFaq(this)">
                                <span>¿Cuáles son los mejores perfumes para el clima de Medellín?</span>
                                <span class="faq-icon">+</span>
                            </div>
                            <div class="faq-answer">
                                Para el clima primaveral y cálido de Medellín durante el día, las mejores opciones son fragancias de la categoría <strong>Frescas y Cítricas</strong> o <strong>Aromáticas Acuáticas</strong> (como <em>Afnan 9am Dive</em> o <em>Light Blue</em>), las cuales proyectan frescura limpia sin saturar el ambiente. Para las noches frescas en sectores como El Poblado, Laureles o Envigado, brillan las fragancias <strong>Dulces / Frutales</strong> y <strong>Especiadas</strong> como <em>Afnan 9pm</em> o <em>Lattafa Asad</em>.
                            </div>
                        </div>
                    </div>

                    <!-- Nube de Etiquetas de Búsqueda SEO -->
                    <div class="seo-tags-box">
                        <div class="seo-tags-title">Búsquedas Frecuentes en Medellín y Colombia</div>
                        <div class="seo-tags-cloud">
                            <span class="seo-tag-pill">Perfumes Originales Medellín</span>
                            <span class="seo-tag-pill">Perfumes Árabes Contraentrega</span>
                            <span class="seo-tag-pill">Comprar Lattafa Colombia</span>
                            <span class="seo-tag-pill">Afnan 9pm Medellín</span>
                            <span class="seo-tag-pill">Lattafa Asad Original</span>
                            <span class="seo-tag-pill">Perfumería Barrio Floresta Cra 86b</span>
                            <span class="seo-tag-pill">Batch Code Verificado</span>
                            <span class="seo-tag-pill">Perfumes de Nicho Medellín</span>
                            <span class="seo-tag-pill">Extrait de Parfum Colombia</span>
                        </div>
                    </div>
                </section>
            </div>
        </div>

        <!-- ==========================================
             PÁGINA 4: CONTACTO
             ========================================== -->
        <div id="page-contacto" class="spa-page">
            <div class="section-wrap">
                <div class="section-header">
                    <span class="section-eyebrow">Atención al Cliente</span>
                    <h1 class="section-title">Contáctanos <span>Directamente</span></h1>
                    <p class="section-desc">Estamos listos para asesorarte en la búsqueda de tu fragancia ideal o resolver dudas sobre disponibilidad y envíos.</p>
                </div>

                <div class="contact-layout">
                    <div class="contact-info-pane">
                        <h2 class="contact-title">AMMAR Perfumería</h2>
                        <p class="contact-subtitle">By Karen Rico · Tienda Boutique en Medellín</p>

                        <div class="contact-items">
                            <div class="contact-item">
                                <div class="contact-icon-box">📍</div>
                                <div>
                                    <div class="contact-item-title">Sede Física</div>
                                    <div class="contact-item-val">Cra 86b #47a 40, Barrio Floresta<br>Medellín, Antioquia, Colombia</div>
                                </div>
                            </div>
                            <div class="contact-item">
                                <div class="contact-icon-box">💬</div>
                                <div>
                                    <div class="contact-item-title">WhatsApp de Ventas</div>
                                    <div class="contact-item-val"><a href="https://wa.me/573128586188" target="_blank">+57 312 858 6188</a></div>
                                </div>
                            </div>
                            <div class="contact-item">
                                <div class="contact-icon-box">📦</div>
                                <div>
                                    <div class="contact-item-title">Cobertura de Envíos</div>
                                    <div class="contact-item-val">Envíos a todo el territorio nacional (Interrapidísimo / Servientrega). Pago contraentrega en Medellín.</div>
                                </div>
                            </div>
                            <div class="contact-item">
                                <div class="contact-icon-box">🕐</div>
                                <div>
                                    <div class="contact-item-title">Jornada de Atención</div>
                                    <div class="contact-item-val">Lunes a Sábado: 9:00 AM – 7:00 PM<br>Domingos y Festivos: 10:00 AM – 3:00 PM</div>
                                </div>
                            </div>
                        </div>

                        <a href="https://wa.me/573128586188?text=Hola%20AMMAR%20Perfumer%C3%ADa%2C%20quisiera%20hacer%20una%20consulta%20directa" target="_blank" rel="noopener noreferrer" class="btn-white" style="width: 100%; justify-content: center;">
                            Iniciar Chat en WhatsApp
                        </a>
                    </div>

                    <div class="contact-map-pane">
                        <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3966.043821874182!2d-75.603157!3d6.2579579999999995!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x8e4429722655084d%3A0xa2ba3201d8d16c6!2zQ3JhLiA4NkIgIyA0N0EtNDAsIExhIEFtw6lyaWNhLCBNZWRlbGzDrW4sIExhIEFtw6lyaWNhLCBNZWRlbGzDrW4sIEFudGlvcXVpYQ!5e0!3m2!1ses-419!2sco!4v1790214902447!5m2!1ses-419!2sco" width="100%" height="100%" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="strict-origin-when-cross-origin" title="Ubicación AMMAR Perfumería"></iframe>
                        <div class="map-address-pill">
                            <span>📍 Cra. 86B #47A-40, La América / Floresta</span>
                            <a href="https://maps.google.com/?q=Cra.+86B+%2347A-40,+La+Am%C3%A9rica,+Medell%C3%ADn,+Antioquia" target="_blank" rel="noopener noreferrer" style="color:var(--gold); font-size:0.75rem; font-weight:600;">Ver en Google Maps ↗</a>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <!-- MODAL DETALLE DE PRODUCTO -->
    <div class="modal-overlay" id="productModal" role="dialog" aria-modal="true">
        <div class="modal-container">
            <button class="modal-close-btn" onclick="closeProductModal()" aria-label="Cerrar modal">✕</button>
            
            <div class="modal-img-col">
                <img id="mImg" src="" alt="Perfume AMMAR">
            </div>

            <div class="modal-info-col">
                <div class="modal-brand" id="mBrand">MARCA</div>
                <h2 class="modal-name" id="mName">Nombre de la Fragancia</h2>
                <div class="modal-price" id="mPrice">$0 <span>COP</span></div>
                
                <div class="modal-tags">
                    <span class="modal-tag tag-simple-modal" id="mTipoSimple">Estilo</span>
                    <span class="modal-tag" id="mGender">Género</span>
                    <span class="modal-tag" id="mFamily">Familia</span>
                </div>

                <div class="modal-pitch" id="mPitch">
                    "Inspiración olfativa..."
                </div>

                <div class="modal-spec-block">
                    <div class="modal-spec-label">Notas Principales</div>
                    <div class="modal-spec-text" id="mNotes">-</div>
                </div>

                <div class="modal-spec-block">
                    <div class="modal-spec-label">Estilo Olfativo</div>
                    <div class="modal-spec-text" id="mStyleDesc">-</div>
                </div>

                <a href="#" target="_blank" rel="noopener noreferrer" class="btn-modal-wa" id="mWaBtn">
                    <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M.057 24l1.687-6.163c-1.041-1.804-1.588-3.849-1.587-5.946.003-6.556 5.338-11.891 11.893-11.891 3.181.001 6.167 1.24 8.413 3.488 2.245 2.248 3.481 5.236 3.48 8.414-.003 6.557-5.338 11.892-11.893 11.892-1.99-.001-3.951-.5-5.688-1.448l-6.305 1.654zm6.597-3.807c1.676.995 3.276 1.591 5.392 1.592 5.448 0 9.886-4.434 9.889-9.885.002-5.462-4.415-9.89-9.881-9.892-5.452 0-9.887 4.434-9.889 9.884-.001 2.225.651 3.891 1.746 5.634l-.999 3.648 3.742-.981z"/></svg>
                    Pedir esta Fragancia por WhatsApp
                </a>
            </div>
        </div>
    </div>

    <!-- MODAL LECTOR DE BLOG EDITORIAL CON FRAGANCIAS RELACIONADAS -->
    <div class="blog-modal-overlay" id="blogModal" role="dialog" aria-modal="true">
        <div class="blog-modal-content">
            <button class="modal-close-btn" onclick="closeBlogModal()" aria-label="Cerrar artículo">✕</button>
            <img id="bmImg" class="blog-modal-hero-img" src="" alt="Artículo Blog">
            <div class="blog-modal-body">
                <span class="blog-card-tag" id="bmTag">CATEGORÍA</span>
                <span style="font-size:0.75rem; color:var(--dim); margin-left:12px;" id="bmDate">FECHA</span>
                <h1 style="font-family:var(--font-heading); color:var(--white); font-size:2rem; margin:16px 0 20px; line-height:1.25;" id="bmTitle">Título del Artículo</h1>
                
                <div id="bmContent"></div>

                <!-- Mini vitrina de perfumes recomendados en este blog -->
                <div class="blog-showcase-box" id="bmShowcaseBox">
                    <div class="blog-showcase-title">
                        <span>✨</span> Fragancias Mencionadas en esta Guía:
                    </div>
                    <div class="blog-mini-grid" id="bmMiniGrid"></div>
                </div>

                <div style="margin-top:32px; padding-top:24px; border-top:1px solid var(--border-hairline); display:flex; gap:16px; flex-wrap:wrap;">
                    <button class="btn-white" id="bmRelatedBtn" onclick="applyBlogRelatedFilter()">
                        Ver Fragancias Relacionadas en el Catálogo →
                    </button>
                    <a href="https://wa.me/573128586188?text=Hola%20AMMAR%2C%20estuve%20leyendo%20su%20art%C3%ADculo%20del%20blog%20y%20quisiera%20asesor%C3%ADa" target="_blank" rel="noopener noreferrer" class="btn-outline-gold">
                        Consultar por WhatsApp
                    </a>
                </div>
            </div>
        </div>
    </div>

    <!-- BOTÓN FLOTANTE WA -->
    <a href="https://wa.me/573128586188?text=Hola%20AMMAR%20Perfumer%C3%ADa%2C%20quisiera%20asesor%C3%ADa%20personalizada" target="_blank" rel="noopener noreferrer" class="floating-wa" aria-label="Hablar por WhatsApp">
        <svg width="30" height="30" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
    </a>

    <!-- FOOTER -->
    <footer class="site-footer">
        <div class="footer-grid">
            <div class="footer-brand">
                <h3>AMMAR<span>.</span></h3>
                <p>By Karen Rico · Alta perfumería en Medellín. Curaduría de fragancias exclusivas, diseñadas para dejar una firma inolvidable.</p>
            </div>
            <div class="footer-col">
                <h4>Navegación</h4>
                <ul>
                    <li><a href="#" onclick="navigateTo('home')">Inicio</a></li>
                    <li><a href="#" onclick="navigateTo('catalogo')">Catálogo Completo</a></li>
                    <li><a href="#" onclick="navigateTo('blog')">Artículos y Guías</a></li>
                    <li><a href="#" onclick="navigateTo('contacto')">Contacto y Tienda</a></li>
                </ul>
            </div>
            <div class="footer-col">
                <h4>Atención</h4>
                <ul>
                    <li><a href="https://wa.me/573128586188" target="_blank">+57 312 858 6188</a></li>
                    <li><a href="#" onclick="navigateTo('contacto')">Cra 86b #47a 40</a></li>
                    <li><a href="#" onclick="navigateTo('contacto')">Barrio Floresta, Medellín</a></li>
                    <li><a href="#" onclick="navigateTo('contacto')">Lun - Sáb: 9am - 7pm</a></li>
                </ul>
            </div>
            <div class="footer-col">
                <h4>Garantía AMMAR</h4>
                <p style="font-size:0.8rem; color:var(--muted); line-height:1.7;">
                    Todos los productos despachados cuentan con verificación de lote, batch code de origen y empaque termo-sellado.
                </p>
            </div>
        </div>

        <div class="footer-bottom">
            <span>© 2026 AMMAR Perfumería By Karen Rico. Todos los derechos reservados.</span>
            <span>Medellín, Colombia</span>
        </div>
    </footer>

    <!-- JAVASCRIPT MASTER APP -->
    <script>
        const PRODUCTS = {DATA_JSON};
        const BRANDS = {json.dumps(marcas_unicas, ensure_ascii=False)};
        const FAMILIES = {json.dumps(familias_unicas, ensure_ascii=False)};
        const TIPOS_SIMPLES = {json.dumps(tipos_simples_unicos, ensure_ascii=False)};
        const BLOG_ARTICLES = {BLOGS_JSON};

        const formatCOP = (num) => '$' + Number(num).toLocaleString('es-CO');

        // Estado del Catálogo
        let currentGender = 'ALL';
        let currentTipoSimple = 'ALL';
        let currentActiveBlogArticle = null;
        let itemsPerPage = 16;
        let displayedCount = 16;
        let filteredProducts = [];

        // ── NAVEGACIÓN SPA ───────────────────────────────────────────
        function navigateTo(pageId) {{
            document.querySelectorAll('.spa-page').forEach(p => p.classList.remove('active'));
            document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));

            const targetPage = document.getElementById('page-' + pageId);
            const targetNav = document.getElementById('nav-' + pageId);

            if (targetPage) targetPage.classList.add('active');
            if (targetNav) targetNav.classList.add('active');

            window.scrollTo({{ top: 0, behavior: 'smooth' }});

            if (pageId === 'catalogo') {{
                renderCatalog();
            }}
        }}

        function toggleDrawer() {{
            const drawer = document.getElementById('mobileDrawer');
            drawer.classList.toggle('open');
        }}

        function toggleFaq(el) {{
            const item = el.parentElement;
            item.classList.toggle('active');
        }}

        // ── INICIALIZACIÓN ────────────────────────────────────────────
        document.addEventListener('DOMContentLoaded', () => {{
            initMarquee();
            initHomeShowcase();
            initFeaturedHome();
            initBlog();
            initFilters();
            applyCatalogFilters();
        }});

        function initMarquee() {{
            const track = document.getElementById('marqueeTrack');
            const items = BRANDS.concat(BRANDS);
            track.innerHTML = items.map(b => `
                <div class="marquee-item">
                    <span>${{b}}</span>
                    <span class="marquee-dot"></span>
                </div>
            `).join('');
        }}

        function initHomeShowcase() {{
            const destacados = PRODUCTS.filter(p => p.destacado);
            if (destacados.length > 0) {{
                const sample = destacados[0];
                document.getElementById('heroImg').src = sample.imagen;
                document.getElementById('heroImg').alt = sample.nombre;
                document.getElementById('heroBrand').textContent = sample.marca;
                document.getElementById('heroName').textContent = sample.nombre;
                document.getElementById('heroPrice').textContent = formatCOP(sample.precio) + ' COP';
                document.getElementById('heroFeaturedCard').querySelector('button').setAttribute('onclick', `openProductModal(${{sample.id}})`);
            }}
        }}

        function initFeaturedHome() {{
            const grid = document.getElementById('homeFeaturedGrid');
            const destacados = PRODUCTS.filter(p => p.destacado).slice(0, 8);
            grid.innerHTML = destacados.map(p => createProductCardHTML(p)).join('');
        }}

        function initBlog() {{
            const homeGrid = document.getElementById('homeBlogGrid');
            const fullGrid = document.getElementById('blogFullGrid');

            const articlesHTML = BLOG_ARTICLES.map(b => `
                <article class="blog-card">
                    <div class="blog-card-img" onclick="openBlogArticle(${{b.id}})" style="cursor:pointer;">
                        <img src="${{b.image}}" alt="${{b.title}}" loading="lazy" onerror="this.src='blogs/images/blog1_guia.jpg'">
                        <span class="blog-card-tag">${{b.tag}}</span>
                    </div>
                    <div class="blog-card-body">
                        <span class="blog-card-date">${{b.date}}</span>
                        <h3 class="blog-card-title">${{b.title}}</h3>
                        <p class="blog-card-excerpt">${{b.excerpt}}</p>
                        <span onclick="openBlogArticle(${{b.id}})" class="blog-card-link">
                            Leer Artículo Completo →
                        </span>
                    </div>
                </article>
            `).join('');

            if (homeGrid) homeGrid.innerHTML = BLOG_ARTICLES.slice(0, 3).map(b => `
                <article class="blog-card">
                    <div class="blog-card-img" onclick="openBlogArticle(${{b.id}})" style="cursor:pointer;">
                        <img src="${{b.image}}" alt="${{b.title}}" loading="lazy" onerror="this.src='blogs/images/blog1_guia.jpg'">
                        <span class="blog-card-tag">${{b.tag}}</span>
                    </div>
                    <div class="blog-card-body">
                        <span class="blog-card-date">${{b.date}}</span>
                        <h3 class="blog-card-title">${{b.title}}</h3>
                        <p class="blog-card-excerpt">${{b.excerpt}}</p>
                        <span onclick="openBlogArticle(${{b.id}})" class="blog-card-link">
                            Leer en el Blog →
                        </span>
                    </div>
                </article>
            `).join('');

            if (fullGrid) fullGrid.innerHTML = articlesHTML;
        }}

        // ── LECTOR MODAL DE BLOG CON VINCULACIÓN AL CATÁLOGO ──────────
        function openBlogArticle(id) {{
            const article = BLOG_ARTICLES.find(a => a.id === id);
            if (!article) return;

            currentActiveBlogArticle = article;

            document.getElementById('bmImg').src = article.image;
            document.getElementById('bmTag').textContent = article.tag;
            document.getElementById('bmDate').textContent = article.date;
            document.getElementById('bmTitle').textContent = article.title;
            document.getElementById('bmContent').innerHTML = article.content;

            // Renderizar mini-vitrina de fragancias relacionadas en este blog
            const miniGrid = document.getElementById('bmMiniGrid');
            const relatedProducts = PRODUCTS.filter(p => {{
                return article.relatedProductNames.some(name => p.nombre.toLowerCase().includes(name.toLowerCase()));
            }}).slice(0, 4);

            if (relatedProducts.length > 0) {{
                document.getElementById('bmShowcaseBox').style.display = 'block';
                miniGrid.innerHTML = relatedProducts.map(p => `
                    <div class="blog-mini-card">
                        <img src="${{p.imagen}}" alt="${{p.nombre}}" loading="lazy">
                        <div class="bm-brand">${{p.marca}}</div>
                        <div class="bm-name">${{p.nombre}}</div>
                        <div class="bm-price">${{formatCOP(p.precio)}} COP</div>
                        <button class="bm-btn-wa" onclick="closeBlogModal(); openProductModal(${{p.id}})">Ver Ficha</button>
                    </div>
                `).join('');
            }} else {{
                document.getElementById('bmShowcaseBox').style.display = 'none';
            }}

            document.getElementById('bmRelatedBtn').textContent = `Ver ${{article.filterLabel}} en el Catálogo →`;

            document.getElementById('blogModal').classList.add('open');
            document.body.style.overflow = 'hidden';
        }}

        function closeBlogModal() {{
            document.getElementById('blogModal').classList.remove('open');
            document.body.style.overflow = '';
        }}

        // Redirige al Catálogo aplicando el filtro exacto del Blog
        function applyBlogRelatedFilter() {{
            if (!currentActiveBlogArticle) return;
            const article = currentActiveBlogArticle;
            closeBlogModal();
            navigateTo('catalogo');

            // Resetear búsqueda general
            document.getElementById('searchInput').value = '';
            document.getElementById('brandFilter').value = 'ALL';
            document.getElementById('familyFilter').value = 'ALL';
            currentGender = 'ALL';
            document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
            document.querySelector('.filter-pill').classList.add('active');

            if (article.filterType === 'tipo_simple') {{
                currentTipoSimple = article.filterValue;
                // Marcar pill activo
                document.querySelectorAll('.pill-tipo').forEach(btn => {{
                    if (btn.textContent.includes(article.filterValue)) {{
                        btn.classList.add('active');
                    }} else {{
                        btn.classList.remove('active');
                    }}
                }});
            }} else if (article.filterType === 'marca') {{
                currentTipoSimple = 'ALL';
                document.querySelectorAll('.pill-tipo').forEach(p => p.classList.remove('active'));
                document.querySelector('.pill-tipo').classList.add('active');
                document.getElementById('brandFilter').value = article.filterValue;
            }} else if (article.filterType === 'destacados') {{
                currentTipoSimple = 'ALL';
                document.querySelectorAll('.pill-tipo').forEach(p => p.classList.remove('active'));
                document.querySelector('.pill-tipo').classList.add('active');
                document.getElementById('sortFilter').value = 'destacado';
            }}

            // Mostrar toast de filtro activo
            const toast = document.getElementById('filterToast');
            document.getElementById('filterToastText').textContent = `🎯 Mostrando selección recomendada para: "${{article.title}}" (${{article.filterLabel}})`;
            toast.classList.add('active');

            displayedCount = itemsPerPage;
            applyCatalogFilters();

            // Scroll suave hacia los productos
            setTimeout(() => {{
                document.getElementById('catalogGrid').scrollIntoView({{ behavior: 'smooth', block: 'start' }});
            }}, 200);
        }}

        function clearBlogFilter() {{
            document.getElementById('filterToast').classList.remove('active');
            currentTipoSimple = 'ALL';
            document.getElementById('brandFilter').value = 'ALL';
            document.getElementById('familyFilter').value = 'ALL';
            document.getElementById('searchInput').value = '';
            document.querySelectorAll('.pill-tipo').forEach(p => p.classList.remove('active'));
            document.querySelector('.pill-tipo').classList.add('active');
            displayedCount = itemsPerPage;
            applyCatalogFilters();
        }}

        document.getElementById('blogModal').addEventListener('click', (e) => {{
            if (e.target.id === 'blogModal') closeBlogModal();
        }});

        // ── FILTROS Y CONTROLES DEL CATÁLOGO ──────────────────────────
        function initFilters() {{
            const brandSelect = document.getElementById('brandFilter');
            const familySelect = document.getElementById('familyFilter');
            const tipoPillsContainer = document.getElementById('tipoSimplePills');

            let pillsHtml = `<button class="pill-tipo active" onclick="setTipoSimpleFilter('ALL', this)">✨ Todas las Fragancias</button>`;
            
            const iconMap = {{
                'Amaderadas y Fuertes': '🌲',
                'Dulces / Frutales': '🍓',
                'Especiadas / Orientales': '🌶️',
                'Florales': '🌸',
                'Frescas y Cítricas': '🌿',
                'Frescas / Aromáticas': '🍃',
                'Otras Fragancias': '💎'
            }};

            TIPOS_SIMPLES.forEach(t => {{
                const icon = iconMap[t] || '🏷️';
                pillsHtml += `<button class="pill-tipo" onclick="setTipoSimpleFilter('${{t}}', this)">${{icon}} ${{t}}</button>`;
            }});
            tipoPillsContainer.innerHTML = pillsHtml;

            BRANDS.forEach(b => {{
                const opt = document.createElement('option');
                opt.value = b;
                opt.textContent = b;
                brandSelect.appendChild(opt);
            }});

            FAMILIES.forEach(f => {{
                const opt = document.createElement('option');
                opt.value = f;
                opt.textContent = f;
                familySelect.appendChild(opt);
            }});

            document.getElementById('searchInput').addEventListener('input', () => {{
                displayedCount = itemsPerPage;
                applyCatalogFilters();
            }});

            brandSelect.addEventListener('change', () => {{
                displayedCount = itemsPerPage;
                applyCatalogFilters();
            }});

            familySelect.addEventListener('change', () => {{
                displayedCount = itemsPerPage;
                applyCatalogFilters();
            }});

            document.getElementById('sortFilter').addEventListener('change', () => {{
                applyCatalogFilters();
            }});
        }}

        function setTipoSimpleFilter(tipo, btn) {{
            currentTipoSimple = tipo;
            document.querySelectorAll('.pill-tipo').forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById('filterToast').classList.remove('active');
            displayedCount = itemsPerPage;
            applyCatalogFilters();
        }}

        function setGenderFilter(gender, btn) {{
            currentGender = gender;
            document.querySelectorAll('.filter-pill').forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            displayedCount = itemsPerPage;
            applyCatalogFilters();
        }}

        function applyCatalogFilters() {{
            const query = (document.getElementById('searchInput').value || '').toLowerCase().trim();
            const brand = document.getElementById('brandFilter').value;
            const family = document.getElementById('familyFilter').value;
            const sort = document.getElementById('sortFilter').value;

            filteredProducts = PRODUCTS.filter(p => {{
                if (currentTipoSimple !== 'ALL' && p.tipo_simple !== currentTipoSimple) return false;
                if (currentGender !== 'ALL' && p.genero !== currentGender) return false;
                if (brand !== 'ALL' && p.marca !== brand) return false;
                if (family !== 'ALL' && p.familia !== family) return false;
                if (query) {{
                    const haystack = (p.nombre + ' ' + p.marca + ' ' + p.tipo_simple + ' ' + p.notas + ' ' + p.familia + ' ' + p.pitch).toLowerCase();
                    if (!haystack.includes(query)) return false;
                }}
                return true;
            }});

            if (sort === 'precio-asc') {{
                filteredProducts.sort((a, b) => a.precio - b.precio);
            }} else if (sort === 'precio-desc') {{
                filteredProducts.sort((a, b) => b.precio - a.precio);
            }} else if (sort === 'nombre') {{
                filteredProducts.sort((a, b) => a.nombre.localeCompare(b.nombre));
            }} else {{
                filteredProducts.sort((a, b) => (b.destacado ? 1 : 0) - (a.destacado ? 1 : 0));
            }}

            renderCatalog();
        }}

        function renderCatalog() {{
            const grid = document.getElementById('catalogGrid');
            const countLabel = document.getElementById('catalogCount');
            const loadBtn = document.getElementById('loadMoreBtn');

            const total = filteredProducts.length;
            countLabel.textContent = `Mostrando ${{Math.min(displayedCount, total)}} de ${{total}} fragancias encontradas`;

            if (total === 0) {{
                grid.innerHTML = `
                    <div style="grid-column: 1/-1; text-align: center; padding: 60px 20px; color: var(--muted);">
                        <div style="font-size: 2.5rem; margin-bottom: 14px;">🔍</div>
                        <h3 style="font-family: var(--font-heading); color: var(--white); font-size: 1.4rem; margin-bottom: 8px;">No encontramos coincidencias</h3>
                        <p style="font-size: 0.85rem;">Prueba eligiendo otro estilo o borrando el término de búsqueda.</p>
                    </div>
                `;
                loadBtn.style.display = 'none';
                return;
            }}

            const visibleItems = filteredProducts.slice(0, displayedCount);
            grid.innerHTML = visibleItems.map(p => createProductCardHTML(p)).join('');

            loadBtn.style.display = displayedCount < total ? 'block' : 'none';
        }}

        function loadMoreProducts() {{
            displayedCount += itemsPerPage;
            renderCatalog();
        }}

        // ── TARJETA DE PRODUCTO ───────────────────────────────────────
        function createProductCardHTML(p) {{
            const msgWa = encodeURIComponent(`Hola AMMAR Perfumería, deseo comprar la fragancia ${{p.nombre}} de ${{p.marca}} (${{formatCOP(p.precio)}} COP - ${{p.tipo_simple}}). ¿Tienen entrega inmediata?`);
            return `
                <div class="product-card">
                    <div class="card-img-wrap" onclick="openProductModal(${{p.id}})">
                        ${{p.destacado ? '<span class="badge-top">Top Selección</span>' : ''}}
                        <span class="badge-tipo-card">${{p.tipo_simple}}</span>
                        <img src="${{p.imagen}}" alt="${{p.nombre}}" loading="lazy" onerror="this.src='https://images.unsplash.com/photo-1541643600914-78b084683702?w=400&q=80'">
                    </div>
                    <div class="card-body">
                        <div class="card-brand">${{p.marca}}</div>
                        <h3 class="card-name">${{p.nombre}}</h3>
                        <div class="card-price">${{formatCOP(p.precio)}} <span>COP</span></div>
                        
                        <div class="card-tags">
                            <span class="card-tag tag-simple">${{p.tipo_simple}}</span>
                            <span class="card-tag">${{p.genero}}</span>
                            <span class="card-tag">${{p.familia}}</span>
                        </div>

                        <p class="card-pitch">${{p.pitch || 'Aroma distinguido de alta concentración.'}}</p>

                        <div class="card-actions">
                            <button class="btn-card-detail" onclick="openProductModal(${{p.id}})">Ver Detalle</button>
                            <a href="https://wa.me/573128586188?text=${{msgWa}}" target="_blank" rel="noopener noreferrer" class="btn-card-wa" aria-label="Comprar ${{p.nombre}} por WhatsApp">
                                <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>
                            </a>
                        </div>
                    </div>
                </div>
            `;
        }}

        // ── MODAL DETALLE DE PRODUCTO ─────────────────────────────────
        function openProductModal(id) {{
            const p = PRODUCTS.find(item => item.id === id);
            if (!p) return;

            document.getElementById('mImg').src = p.imagen;
            document.getElementById('mImg').alt = p.nombre;
            document.getElementById('mBrand').textContent = p.marca;
            document.getElementById('mName').textContent = p.nombre;
            document.getElementById('mPrice').innerHTML = `${{formatCOP(p.precio)}} <span>COP</span>`;
            document.getElementById('mTipoSimple').textContent = p.tipo_simple;
            document.getElementById('mGender').textContent = p.genero;
            document.getElementById('mFamily').textContent = p.familia;
            document.getElementById('mPitch').textContent = `"${{p.pitch}}"`;
            document.getElementById('mNotes').textContent = p.notas;
            document.getElementById('mStyleDesc').textContent = `Categoría olfativa simplificada: ${{p.tipo_simple}}. Diseñado para adaptarse con facilidad a tu estilo diario.`;

            const msgWa = encodeURIComponent(`¡Hola AMMAR Perfumería! Estoy viendo ${{p.nombre}} de ${{p.marca}} (${{p.tipo_simple}}) en la tienda virtual (${{formatCOP(p.precio)}} COP) y deseo coordinar la compra y entrega. ¿Tienen disponibilidad?`);
            document.getElementById('mWaBtn').href = `https://wa.me/573128586188?text=${{msgWa}}`;

            document.getElementById('productModal').classList.add('open');
            document.body.style.overflow = 'hidden';
        }}

        function closeProductModal() {{
            document.getElementById('productModal').classList.remove('open');
            document.body.style.overflow = '';
        }}

        document.getElementById('productModal').addEventListener('click', (e) => {{
            if (e.target.id === 'productModal') closeProductModal();
        }});

        document.addEventListener('keydown', (e) => {{
            if (e.key === 'Escape') {{
                closeProductModal();
                closeBlogModal();
            }}
        }});
    </script>
</body>
</html>
"""

with open(OUT_HTML, "w", encoding="utf-8") as f:
    f.write(HTML_CONTENT)

print(f"Éxito: Generado index.html con {len(products)} perfumes, Blog interactivo con vitrina y SEO Medellín/Colombia.")

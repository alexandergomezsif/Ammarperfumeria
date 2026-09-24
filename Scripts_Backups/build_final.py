import pandas as pd
import json
import random
import os
import re
import math

random.seed(99)

df = pd.read_excel("Base_Datos_Perfumes.xlsx")
df["Marca"] = df["Marca"].ffill()

productos = []

# Build image map
image_map = {}
for root, dirs, files in os.walk("fotos"):
    for file in files:
        if file.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
            name = os.path.splitext(file)[0].lower()
            name = re.sub(r"[^a-z0-9]", "", name)
            rel_path = os.path.join(root, file).replace("\\", "/")
            image_map[name] = rel_path

for idx, row in df.iterrows():
    marca = str(row["Marca"]).strip() if pd.notna(row["Marca"]) else "Desconocida"
    ref = str(row["Referencia"]).strip() if pd.notna(row["Referencia"]) else "Sin Nombre"
    genero = str(row["Genero"]).strip() if pd.notna(row["Genero"]) else ""
    familia = str(row["Familia_Olfativa"]).strip() if pd.notna(row["Familia_Olfativa"]) else ""
    notas = str(row["Notas_Principales"]).strip() if pd.notna(row["Notas_Principales"]) else ""
    topologia = str(row["Topologia Y Arquitectura olfativa"]).strip() if pd.notna(row["Topologia Y Arquitectura olfativa"]) else ""
    pitch = str(row["Inspiracion_Ventas"]).strip() if pd.notna(row["Inspiracion_Ventas"]) else ""
    
    price = random.randint(160, 250) * 1000
    
    ref_norm = re.sub(r"[^a-z0-9]", "", ref.lower())
    image = image_map.get(ref_norm, "https://via.placeholder.com/400x500?text=No+Image")
    
    productos.append({
        "id": idx,
        "marca": marca,
        "nombre": ref,
        "genero": genero,
        "familia": familia,
        "notas": notas,
        "topologia": topologia,
        "pitch": pitch,
        "precio": price,
        "imagen": image,
        "destacado": random.random() < 0.12
    })

destacados = [p for p in productos if p["destacado"]]
if len(destacados) < 5:
    for p in productos:
        if not p["destacado"]:
            p["destacado"] = True
            destacados.append(p)
            if len(destacados) >= 5: break
elif len(destacados) > 5:
    for p in destacados[5:]:
        p["destacado"] = False

data_json = json.dumps(productos, ensure_ascii=False)

html_content = f"""<!DOCTYPE html>
<html lang="es" class="dark">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>AMMAR Perfumería | By Karen Rico</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;600;700&family=Inter:wght@300;400;500;600&family=Playfair+Display:ital,wght@0,400;0,600;1,400&display=swap" rel="stylesheet">
    <script src="https://cdn.tailwindcss.com"></script>
    <script>
        tailwind.config = {{
            darkMode: 'class',
            theme: {{
                extend: {{
                    colors: {{
                        dark: '#050505',
                        light: '#fafafa',
                        gold: '#d4af37',
                        cardDark: '#0d0d0d',
                        cardDarkHover: '#161616'
                    }},
                    fontFamily: {{
                        cinzel: ['Cinzel', 'serif'],
                        inter: ['Inter', 'sans-serif'],
                        playfair: ['Playfair Display', 'serif']
                    }}
                }}
            }}
        }}
    </script>
    <style>
        body {{
            background-color: theme('colors.dark');
            color: theme('colors.light');
            font-family: theme('fontFamily.inter');
            transition: background-color 0.3s, color 0.3s;
        }}
        html:not(.dark) body {{
            background-color: theme('colors.light');
            color: theme('colors.dark');
        }}
        html:not(.dark) .card {{
            background-color: #fff;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        }}
        html:not(.dark) .card:hover {{
            background-color: #f3f4f6;
        }}
        .card {{
            background-color: theme('colors.cardDark');
            border: 1px solid rgba(255,255,255,0.08);
            transition: all 0.3s ease;
        }}
        .card:hover {{
            background-color: theme('colors.cardDarkHover');
        }}
        html:not(.dark) .card {{ border-color: rgba(0,0,0,0.08); }}
        
        .page {{ display: none; }}
        .page.active {{ display: block; }}
        .nav-link.active {{ border-bottom: 2px solid theme('colors.gold'); }}
        
        /* Modal scrollbar */
        ::-webkit-scrollbar {{ width: 8px; }}
        ::-webkit-scrollbar-track {{ background: transparent; }}
        ::-webkit-scrollbar-thumb {{ background: rgba(212, 175, 55, 0.3); border-radius: 4px; }}
        ::-webkit-scrollbar-thumb:hover {{ background: rgba(212, 175, 55, 0.6); }}
    </style>
</head>
<body class="antialiased min-h-screen flex flex-col">

    <!-- Navbar -->
    <nav class="sticky top-0 z-50 bg-dark/90 backdrop-blur-md border-b border-white/10 transition-colors duration-300">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex justify-between h-20 items-center">
                <div class="flex-shrink-0 flex items-center cursor-pointer" onclick="navigate('home')">
                    <div class="text-center">
                        <h1 class="font-cinzel text-2xl font-bold tracking-widest text-gold">AMMAR</h1>
                        <p class="font-playfair text-xs italic opacity-70">By Karen Rico</p>
                    </div>
                </div>
                <div class="hidden md:flex space-x-8 items-center">
                    <a href="#" onclick="navigate('home')" class="nav-link nav-home active text-sm tracking-widest uppercase hover:text-gold transition px-2 py-1">Inicio</a>
                    <a href="#" onclick="navigate('catalogo')" class="nav-link nav-catalogo text-sm tracking-widest uppercase hover:text-gold transition px-2 py-1">Catálogo</a>
                    <a href="#" onclick="navigate('blog')" class="nav-link nav-blog text-sm tracking-widest uppercase hover:text-gold transition px-2 py-1">Blog</a>
                    <a href="#" onclick="navigate('contacto')" class="nav-link nav-contacto text-sm tracking-widest uppercase hover:text-gold transition px-2 py-1">Contacto</a>
                    <button onclick="toggleTheme()" class="p-2 rounded-full hover:bg-white/5 transition">
                        <svg class="w-5 h-5 text-gold" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>
                    </button>
                </div>
                <div class="md:hidden flex items-center">
                    <button onclick="toggleTheme()" class="p-2 mr-2 rounded-full hover:bg-white/5 transition">
                        <svg class="w-5 h-5 text-gold" fill="none" stroke="currentColor" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>
                    </button>
                    <button onclick="toggleMenu()" class="text-gray-300 hover:text-white">
                        <svg class="h-6 w-6" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h16"/>
                        </svg>
                    </button>
                </div>
            </div>
        </div>
    </nav>

    <!-- Mobile Menu -->
    <div id="mobileMenu" class="fixed inset-0 z-40 bg-dark/95 backdrop-blur-lg transform translate-x-full transition-transform duration-300 pt-24 px-8">
        <button onclick="toggleMenu()" class="absolute top-6 right-6 text-gray-300 hover:text-white">
            <svg class="h-8 w-8" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"/>
            </svg>
        </button>
        <div class="flex flex-col space-y-6 text-center">
            <a href="#" onclick="navigate('home'); toggleMenu()" class="text-2xl font-cinzel text-white hover:text-gold">Inicio</a>
            <a href="#" onclick="navigate('catalogo'); toggleMenu()" class="text-2xl font-cinzel text-white hover:text-gold">Catálogo</a>
            <a href="#" onclick="navigate('blog'); toggleMenu()" class="text-2xl font-cinzel text-white hover:text-gold">Blog</a>
            <a href="#" onclick="navigate('contacto'); toggleMenu()" class="text-2xl font-cinzel text-white hover:text-gold">Contacto</a>
        </div>
    </div>

    <!-- Main Content -->
    <main class="flex-grow">
        <!-- Home Page -->
        <div id="home" class="page active">
            <!-- Hero -->
            <div class="relative h-[80vh] flex items-center justify-center overflow-hidden">
                <div class="absolute inset-0 bg-[url('https://images.unsplash.com/photo-1615397323281-a587428169cd?ixlib=rb-4.0.3&auto=format&fit=crop&w=1920&q=80')] bg-cover bg-center opacity-30"></div>
                <div class="absolute inset-0 bg-gradient-to-t from-dark to-transparent"></div>
                <div class="relative z-10 text-center px-4">
                    <h2 class="font-cinzel text-5xl md:text-7xl font-bold mb-4 text-white">La Esencia de la <span class="text-gold">Elegancia</span></h2>
                    <p class="font-playfair text-xl md:text-2xl text-gray-300 mb-8 max-w-2xl mx-auto">Descubre fragancias excepcionales que definen tu presencia y dejan una huella imborrable.</p>
                    <button onclick="navigate('catalogo')" class="bg-gold text-dark font-semibold px-8 py-3 uppercase tracking-widest hover:bg-white transition duration-300">Explorar Catálogo</button>
                </div>
            </div>

            <!-- Guarantees -->
            <div class="py-16 border-b border-white/5">
                <div class="max-w-7xl mx-auto px-4 grid grid-cols-1 md:grid-cols-4 gap-8 text-center">
                    <div>
                        <div class="text-gold text-4xl mb-4">✨</div>
                        <h3 class="font-cinzel font-bold mb-2">Autenticidad 100%</h3>
                        <p class="text-sm opacity-70">Garantizamos la originalidad de todos nuestros productos.</p>
                    </div>
                    <div>
                        <div class="text-gold text-4xl mb-4">🚚</div>
                        <h3 class="font-cinzel font-bold mb-2">Envío Nacional</h3>
                        <p class="text-sm opacity-70">Envíos seguros a toda Colombia.</p>
                    </div>
                    <div>
                        <div class="text-gold text-4xl mb-4">💎</div>
                        <h3 class="font-cinzel font-bold mb-2">Marcas Exclusivas</h3>
                        <p class="text-sm opacity-70">Selección premium de las mejores casas perfumeras.</p>
                    </div>
                    <div>
                        <div class="text-gold text-4xl mb-4">🔒</div>
                        <h3 class="font-cinzel font-bold mb-2">Pago Seguro</h3>
                        <p class="text-sm opacity-70">Múltiples medios de pago 100% confiables.</p>
                    </div>
                </div>
            </div>

            <!-- Featured Products -->
            <div class="py-20 max-w-7xl mx-auto px-4">
                <div class="text-center mb-16">
                    <h2 class="font-cinzel text-3xl md:text-4xl font-bold mb-4">Selección <span class="text-gold">Destacada</span></h2>
                    <p class="font-playfair italic opacity-70">Nuestras recomendaciones exclusivas</p>
                </div>
                <div class="grid grid-cols-1 md:grid-cols-5 gap-6" id="featuredGrid">
                    <!-- Featured products injected via JS -->
                </div>
                <div class="text-center mt-12">
                    <button onclick="navigate('catalogo')" class="border border-gold text-gold hover:bg-gold hover:text-dark px-8 py-3 uppercase tracking-widest transition duration-300">Ver Todos</button>
                </div>
            </div>

            <!-- Contact Strip with Map -->
            <div class="py-20 bg-cardDark border-t border-white/5">
                <div class="max-w-7xl mx-auto px-4 grid grid-cols-1 md:grid-cols-2 gap-12 items-center">
                    <div>
                        <h2 class="font-cinzel text-3xl font-bold mb-6">Visítanos en <span class="text-gold">Medellín</span></h2>
                        <p class="mb-4 flex items-center"><span class="text-gold mr-3">📍</span> Cra 86 #47a-40, Barrio Floresta, Medellín, Antioquia</p>
                        <p class="mb-4 flex items-center"><span class="text-gold mr-3">🕒</span> Lunes a Sábado: 9am - 7pm<br>Domingos: 10am - 3pm</p>
                        <p class="mb-8 flex items-center"><span class="text-gold mr-3">📱</span> +57 312 858 6188</p>
                        <a href="https://wa.me/573128586188" target="_blank" class="inline-flex items-center bg-green-600 text-white px-6 py-3 rounded-full hover:bg-green-700 transition">
                            <svg class="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51a12.8 12.8 0 0 0-.57-.01c-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
                            Escríbenos por WhatsApp
                        </a>
                    </div>
                    <div class="h-64 md:h-96 rounded-lg overflow-hidden border border-white/10">
                        <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3966.02706399103!2d-75.60251702422718!3d6.260173626162235!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x8e44299b4d4dc0e7%3A0xc3f5fb4cfb159b95!2sCra.%2086%20%2347a-40%2C%20La%20Floresta%2C%20Medell%C3%ADn%2C%20La%20Am%C3%A9rica%2C%20Medell%C3%ADn%2C%20Antioquia!5e0!3m2!1ses!2sco!4v1716301234567!5m2!1ses!2sco" width="100%" height="100%" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>
                    </div>
                </div>
            </div>
        </div>

        <!-- Catalogo Page -->
        <div id="catalogo" class="page py-12">
            <div class="max-w-7xl mx-auto px-4">
                <h2 class="font-cinzel text-4xl font-bold mb-8 text-center">Nuestro <span class="text-gold">Catálogo</span></h2>
                
                <!-- Filters -->
                <div class="bg-cardDark p-6 rounded-lg mb-8 border border-white/5">
                    <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
                        <input type="text" id="searchInput" placeholder="Buscar por nombre o marca..." class="bg-dark border border-white/10 p-3 rounded text-light focus:border-gold focus:outline-none w-full">
                        <select id="brandFilter" class="bg-dark border border-white/10 p-3 rounded text-light focus:border-gold focus:outline-none w-full">
                            <option value="">Todas las Marcas</option>
                        </select>
                        <select id="genderFilter" class="bg-dark border border-white/10 p-3 rounded text-light focus:border-gold focus:outline-none w-full">
                            <option value="">Cualquier Género</option>
                        </select>
                        <select id="familyFilter" class="bg-dark border border-white/10 p-3 rounded text-light focus:border-gold focus:outline-none w-full">
                            <option value="">Cualquier Familia Olfativa</option>
                        </select>
                    </div>
                </div>

                <!-- Grid -->
                <div class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-6" id="catalogGrid">
                    <!-- Products injected via JS -->
                </div>
                
                <!-- Load More -->
                <div class="text-center mt-12">
                    <button id="loadMoreBtn" class="bg-gold text-dark font-semibold px-8 py-3 uppercase tracking-widest hover:bg-white transition duration-300 hidden">Cargar Más</button>
                </div>
            </div>
        </div>

        <!-- Blog Page -->
        <div id="blog" class="page py-12">
            <div class="max-w-7xl mx-auto px-4">
                <h2 class="font-cinzel text-4xl font-bold mb-12 text-center">Blog <span class="text-gold">Olfativo</span></h2>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-10">
                    <!-- Article 1 -->
                    <div class="card rounded-xl overflow-hidden cursor-pointer group">
                        <div class="h-64 overflow-hidden relative">
                            <img src="https://images.unsplash.com/photo-1592945403244-b3fbafd7f539?ixlib=rb-4.0.3&auto=format&fit=crop&w=800&q=80" alt="Blog 1" class="w-full h-full object-cover group-hover:scale-105 transition duration-500">
                            <span class="absolute top-4 left-4 bg-gold text-dark text-xs font-bold px-3 py-1 uppercase tracking-wider">Guía</span>
                        </div>
                        <div class="p-8">
                            <p class="text-sm opacity-60 mb-2">15 Septiembre, 2026</p>
                            <h3 class="font-cinzel text-2xl font-bold mb-3 group-hover:text-gold transition">Cómo elegir tu fragancia firma</h3>
                            <p class="opacity-80 font-playfair line-clamp-3">Encontrar el perfume que te defina no es tarea fácil. Requiere entender tu química corporal, tu personalidad y las familias olfativas...</p>
                        </div>
                    </div>
                    <!-- Article 2 -->
                    <div class="card rounded-xl overflow-hidden cursor-pointer group">
                        <div class="h-64 overflow-hidden relative">
                            <img src="https://images.unsplash.com/photo-1596462502278-27bfdc403348?ixlib=rb-4.0.3&auto=format&fit=crop&w=800&q=80" alt="Blog 2" class="w-full h-full object-cover group-hover:scale-105 transition duration-500">
                            <span class="absolute top-4 left-4 bg-gold text-dark text-xs font-bold px-3 py-1 uppercase tracking-wider">Tendencias</span>
                        </div>
                        <div class="p-8">
                            <p class="text-sm opacity-60 mb-2">2 Agosto, 2026</p>
                            <h3 class="font-cinzel text-2xl font-bold mb-3 group-hover:text-gold transition">Las notas de madera en otoño</h3>
                            <p class="opacity-80 font-playfair line-clamp-3">El sándalo, el cedro y el oud toman protagonismo esta temporada. Descubre por qué las fragancias amaderadas son perfectas para el clima frío.</p>
                        </div>
                    </div>
                    <!-- Article 3 -->
                    <div class="card rounded-xl overflow-hidden cursor-pointer group">
                        <div class="h-64 overflow-hidden relative">
                            <img src="https://images.unsplash.com/photo-1583445013765-46c20c4a6772?ixlib=rb-4.0.3&auto=format&fit=crop&w=800&q=80" alt="Blog 3" class="w-full h-full object-cover group-hover:scale-105 transition duration-500">
                            <span class="absolute top-4 left-4 bg-gold text-dark text-xs font-bold px-3 py-1 uppercase tracking-wider">Historia</span>
                        </div>
                        <div class="p-8">
                            <p class="text-sm opacity-60 mb-2">10 Julio, 2026</p>
                            <h3 class="font-cinzel text-2xl font-bold mb-3 group-hover:text-gold transition">El origen del perfume en la antigua Francia</h3>
                            <p class="opacity-80 font-playfair line-clamp-3">Un viaje a través del tiempo hacia Grasse, la cuna mundial de la perfumería, y cómo revolucionó la industria para siempre.</p>
                        </div>
                    </div>
                    <!-- Article 4 -->
                    <div class="card rounded-xl overflow-hidden cursor-pointer group">
                        <div class="h-64 overflow-hidden relative">
                            <img src="https://images.unsplash.com/photo-1557170334-a9632e77c6e4?ixlib=rb-4.0.3&auto=format&fit=crop&w=800&q=80" alt="Blog 4" class="w-full h-full object-cover group-hover:scale-105 transition duration-500">
                            <span class="absolute top-4 left-4 bg-gold text-dark text-xs font-bold px-3 py-1 uppercase tracking-wider">Tips</span>
                        </div>
                        <div class="p-8">
                            <p class="text-sm opacity-60 mb-2">22 Junio, 2026</p>
                            <h3 class="font-cinzel text-2xl font-bold mb-3 group-hover:text-gold transition">Secretos para que tu perfume dure todo el día</h3>
                            <p class="opacity-80 font-playfair line-clamp-3">¿Sientes que tu fragancia desaparece rápido? Aplica estos sencillos trucos en los puntos de pulso y maximiza la fijación.</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Contacto Page -->
        <div id="contacto" class="page py-12">
            <div class="max-w-7xl mx-auto px-4">
                <h2 class="font-cinzel text-4xl font-bold mb-12 text-center">Contáctanos</h2>
                <div class="grid grid-cols-1 md:grid-cols-2 gap-12">
                    <div class="bg-cardDark p-8 rounded-xl border border-white/5 shadow-xl">
                        <h3 class="font-cinzel text-2xl font-bold mb-6 text-gold">Información de Contacto</h3>
                        <div class="space-y-6">
                            <div class="flex items-start">
                                <span class="text-2xl mr-4">📍</span>
                                <div>
                                    <h4 class="font-bold mb-1">Dirección Física</h4>
                                    <p class="opacity-80">Cra 86 #47a-40, Barrio Floresta<br>Medellín, Antioquia, Colombia</p>
                                </div>
                            </div>
                            <div class="flex items-start">
                                <span class="text-2xl mr-4">📱</span>
                                <div>
                                    <h4 class="font-bold mb-1">WhatsApp</h4>
                                    <p class="opacity-80">+57 312 858 6188</p>
                                    <a href="https://wa.me/573128586188" target="_blank" class="text-gold text-sm hover:underline mt-1 inline-block">Enviar mensaje</a>
                                </div>
                            </div>
                            <div class="flex items-start">
                                <span class="text-2xl mr-4">🕒</span>
                                <div>
                                    <h4 class="font-bold mb-1">Horario de Atención</h4>
                                    <p class="opacity-80">Lunes a Sábado: 9:00 AM - 7:00 PM<br>Domingos: 10:00 AM - 3:00 PM</p>
                                </div>
                            </div>
                            <div class="flex items-start">
                                <span class="text-2xl mr-4">🚚</span>
                                <div>
                                    <h4 class="font-bold mb-1">Envíos</h4>
                                    <p class="opacity-80">Realizamos envíos a nivel nacional. Entrega en Medellín en menos de 24 horas.</p>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="h-[500px] rounded-xl overflow-hidden border border-white/10 shadow-xl">
                        <iframe src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3966.02706399103!2d-75.60251702422718!3d6.260173626162235!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x8e44299b4d4dc0e7%3A0xc3f5fb4cfb159b95!2sCra.%2086%20%2347a-40%2C%20La%20Floresta%2C%20Medell%C3%ADn%2C%20La%20Am%C3%A9rica%2C%20Medell%C3%ADn%2C%20Antioquia!5e0!3m2!1ses!2sco!4v1716301234567!5m2!1ses!2sco" width="100%" height="100%" style="border:0;" allowfullscreen="" loading="lazy" referrerpolicy="no-referrer-when-downgrade"></iframe>
                    </div>
                </div>
            </div>
        </div>
    </main>

    <!-- Footer -->
    <footer class="bg-cardDark border-t border-white/5 py-12 mt-auto">
        <div class="max-w-7xl mx-auto px-4 text-center">
            <h2 class="font-cinzel text-2xl font-bold tracking-widest text-gold mb-2">AMMAR</h2>
            <p class="font-playfair italic opacity-70 mb-8">By Karen Rico</p>
            <div class="flex justify-center space-x-6 mb-8">
                <a href="#" class="text-gray-400 hover:text-gold transition">Instagram</a>
                <a href="#" class="text-gray-400 hover:text-gold transition">Facebook</a>
                <a href="#" class="text-gray-400 hover:text-gold transition">TikTok</a>
            </div>
            <p class="text-sm opacity-50">&copy; 2026 AMMAR Perfumería. Todos los derechos reservados.</p>
        </div>
    </footer>

    <!-- WhatsApp Floating Button -->
    <a href="https://wa.me/573128586188" target="_blank" class="fixed bottom-6 right-6 bg-green-500 text-white p-4 rounded-full shadow-2xl hover:bg-green-600 hover:scale-110 transition duration-300 z-50 flex items-center justify-center">
        <svg class="w-8 h-8" fill="currentColor" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51a12.8 12.8 0 0 0-.57-.01c-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 0 1-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 0 1-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 0 1 2.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0 0 12.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 0 0 5.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 0 0-3.48-8.413Z"/></svg>
    </a>

    <!-- Product Modal -->
    <div id="productModal" class="fixed inset-0 z-[60] bg-black/80 backdrop-blur-sm hidden items-center justify-center p-4 opacity-0 transition-opacity duration-300">
        <div class="bg-cardDark border border-white/10 rounded-2xl max-w-4xl w-full max-h-[90vh] overflow-y-auto flex flex-col md:flex-row relative transform scale-95 transition-transform duration-300 shadow-2xl">
            <button onclick="closeModal()" class="absolute top-4 right-4 z-10 bg-black/50 text-white rounded-full p-2 hover:bg-gold transition">
                <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
            </button>
            <div class="md:w-1/2 h-80 md:h-auto bg-dark p-6 flex items-center justify-center">
                <img id="modalImg" src="" alt="Product" class="max-h-full max-w-full object-contain drop-shadow-2xl">
            </div>
            <div class="md:w-1/2 p-8 flex flex-col">
                <p id="modalBrand" class="text-gold font-cinzel tracking-widest text-sm mb-1 uppercase"></p>
                <h3 id="modalName" class="font-cinzel text-3xl font-bold mb-2"></h3>
                <p id="modalPrice" class="text-2xl font-semibold mb-4 text-light/90"></p>
                
                <div class="flex flex-wrap gap-2 mb-6">
                    <span id="modalGender" class="text-xs border border-white/20 px-3 py-1 rounded-full opacity-80"></span>
                    <span id="modalFamily" class="text-xs border border-white/20 px-3 py-1 rounded-full opacity-80"></span>
                </div>
                
                <div class="space-y-4 mb-8 flex-grow">
                    <div>
                        <h4 class="text-gold text-sm font-bold uppercase tracking-wider mb-1">Inspiración</h4>
                        <p id="modalPitch" class="text-sm opacity-80 font-playfair italic"></p>
                    </div>
                    <div>
                        <h4 class="text-gold text-sm font-bold uppercase tracking-wider mb-1">Notas Principales</h4>
                        <p id="modalNotes" class="text-sm opacity-80"></p>
                    </div>
                    <div>
                        <h4 class="text-gold text-sm font-bold uppercase tracking-wider mb-1">Topología Olfativa</h4>
                        <p id="modalTopo" class="text-sm opacity-80"></p>
                    </div>
                </div>
                
                <a id="modalWhatsApp" href="#" target="_blank" class="w-full bg-gold text-dark text-center font-bold uppercase tracking-widest py-4 rounded hover:bg-white transition duration-300 flex justify-center items-center">
                    <svg class="w-5 h-5 mr-2" fill="currentColor" viewBox="0 0 24 24"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51a12.8 12.8 0 0 0-.57-.01c-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347"/></svg>
                    Lo Quiero
                </a>
            </div>
        </div>
    </div>

    <!-- Application Logic -->
    <script>
        const DATA = {data_json};
        
        let filteredData = [...DATA];
        let currentPage = 1;
        const ITEMS_PER_PAGE = 18;
        
        // Navigation
        function navigate(pageId) {{
            document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
            document.getElementById(pageId).classList.add('active');
            
            document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
            const link = document.querySelector(`.nav-${{pageId}}`);
            if(link) link.classList.add('active');
            
            window.scrollTo(0,0);
        }}
        
        function toggleMenu() {{
            const menu = document.getElementById('mobileMenu');
            if(menu.classList.contains('translate-x-full')) {{
                menu.classList.remove('translate-x-full');
            }} else {{
                menu.classList.add('translate-x-full');
            }}
        }}

        // Theme Toggle
        function toggleTheme() {{
            const html = document.documentElement;
            if (html.classList.contains('dark')) {{
                html.classList.remove('dark');
                localStorage.setItem('theme', 'light');
            }} else {{
                html.classList.add('dark');
                localStorage.setItem('theme', 'dark');
            }}
        }}

        // Format Currency
        function formatPrice(price) {{
            return new Intl.NumberFormat('es-CO', {{ style: 'currency', currency: 'COP', maximumFractionDigits: 0 }}).format(price);
        }}

        // Initialize App
        function init() {{
            // Theme init
            if (localStorage.getItem('theme') === 'light') {{
                document.documentElement.classList.remove('dark');
            }}
            
            // Populate filters
            const brands = [...new Set(DATA.map(p => p.marca))].filter(Boolean).sort();
            const genders = [...new Set(DATA.map(p => p.genero))].filter(Boolean).sort();
            const families = [...new Set(DATA.map(p => p.familia))].filter(Boolean).sort();
            
            const brandFilter = document.getElementById('brandFilter');
            brands.forEach(b => brandFilter.innerHTML += `<option value="${{b}}">${{b}}</option>`);
            
            const genderFilter = document.getElementById('genderFilter');
            genders.forEach(g => genderFilter.innerHTML += `<option value="${{g}}">${{g}}</option>`);
            
            const familyFilter = document.getElementById('familyFilter');
            families.forEach(f => familyFilter.innerHTML += `<option value="${{f}}">${{f}}</option>`);
            
            // Event listeners
            document.getElementById('searchInput').addEventListener('input', applyFilters);
            document.getElementById('brandFilter').addEventListener('change', applyFilters);
            document.getElementById('genderFilter').addEventListener('change', applyFilters);
            document.getElementById('familyFilter').addEventListener('change', applyFilters);
            document.getElementById('loadMoreBtn').addEventListener('click', loadMore);
            
            // Render Featured
            const featuredData = DATA.filter(p => p.destacado).slice(0, 5);
            const featuredHtml = featuredData.map(p => createProductCard(p)).join('');
            document.getElementById('featuredGrid').innerHTML = featuredHtml;
            
            // Initial Render Catalog
            renderCatalog();
        }}
        
        function createProductCard(product) {{
            return `
                <div class="card rounded-xl overflow-hidden flex flex-col h-full group">
                    <div class="relative h-64 bg-white/5 p-4 flex items-center justify-center overflow-hidden cursor-pointer" onclick="openModal(${{product.id}})">
                        <img src="${{product.imagen}}" alt="${{product.nombre}}" class="max-h-full object-contain group-hover:scale-110 transition duration-500 drop-shadow-xl">
                        ${{product.destacado ? '<span class="absolute top-2 left-2 bg-gold text-dark text-xs font-bold px-2 py-1 uppercase">Top</span>' : ''}}
                    </div>
                    <div class="p-5 flex flex-col flex-grow">
                        <p class="text-gold font-cinzel text-xs uppercase tracking-widest mb-1">${{product.marca}}</p>
                        <h3 class="font-bold text-lg mb-2 line-clamp-1">${{product.nombre}}</h3>
                        <p class="text-xl font-semibold mb-3 text-light/90">${{formatPrice(product.precio)}}</p>
                        <p class="text-sm opacity-60 mb-4 line-clamp-2 italic font-playfair flex-grow">${{product.pitch || 'Una fragancia inolvidable que define tu estilo.'}}</p>
                        <button onclick="openModal(${{product.id}})" class="w-full border border-white/20 py-2 uppercase text-xs tracking-widest hover:border-gold hover:text-gold transition">Ver Detalle</button>
                    </div>
                </div>
            `;
        }}
        
        function applyFilters() {{
            const search = document.getElementById('searchInput').value.toLowerCase();
            const brand = document.getElementById('brandFilter').value;
            const gender = document.getElementById('genderFilter').value;
            const family = document.getElementById('familyFilter').value;
            
            filteredData = DATA.filter(p => {{
                const matchSearch = p.nombre.toLowerCase().includes(search) || p.marca.toLowerCase().includes(search);
                const matchBrand = !brand || p.marca === brand;
                const matchGender = !gender || p.genero === gender;
                const matchFamily = !family || p.familia === family;
                return matchSearch && matchBrand && matchGender && matchFamily;
            }});
            
            currentPage = 1;
            renderCatalog();
        }}
        
        function renderCatalog() {{
            const grid = document.getElementById('catalogGrid');
            const items = filteredData.slice(0, currentPage * ITEMS_PER_PAGE);
            
            if(items.length === 0) {{
                grid.innerHTML = '<div class="col-span-full text-center py-12 opacity-50">No se encontraron productos con estos filtros.</div>';
            }} else {{
                grid.innerHTML = items.map(p => createProductCard(p)).join('');
            }}
            
            const btn = document.getElementById('loadMoreBtn');
            if(currentPage * ITEMS_PER_PAGE < filteredData.length) {{
                btn.classList.remove('hidden');
            }} else {{
                btn.classList.add('hidden');
            }}
        }}
        
        function loadMore() {{
            currentPage++;
            renderCatalog();
        }}
        
        function openModal(id) {{
            const p = DATA.find(x => x.id === id);
            if(!p) return;
            
            document.getElementById('modalImg').src = p.imagen;
            document.getElementById('modalBrand').textContent = p.marca;
            document.getElementById('modalName').textContent = p.nombre;
            document.getElementById('modalPrice').textContent = formatPrice(p.precio);
            
            document.getElementById('modalGender').textContent = p.genero || 'Unisex';
            document.getElementById('modalFamily').textContent = p.familia || 'Perfume';
            
            document.getElementById('modalPitch').textContent = p.pitch || 'Exclusiva fragancia de nuestra colección.';
            document.getElementById('modalNotes').textContent = p.notas || 'Consultar empaque.';
            document.getElementById('modalTopo').textContent = p.topologia || 'Evolución clásica.';
            
            const msg = encodeURIComponent(`¡Hola! Me interesa la fragancia ${{p.nombre}} de ${{p.marca}} por ${{formatPrice(p.precio)}}. ¿Tienen disponibilidad?`);
            document.getElementById('modalWhatsApp').href = `https://wa.me/573128586188?text=${{msg}}`;
            
            const modal = document.getElementById('productModal');
            modal.classList.remove('hidden');
            modal.classList.add('flex');
            // small delay for transition
            setTimeout(() => {{
                modal.classList.remove('opacity-0');
                modal.querySelector('div').classList.remove('scale-95');
            }}, 10);
        }}
        
        function closeModal() {{
            const modal = document.getElementById('productModal');
            modal.classList.add('opacity-0');
            modal.querySelector('div').classList.add('scale-95');
            setTimeout(() => {{
                modal.classList.add('hidden');
                modal.classList.remove('flex');
            }}, 300);
        }}
        
        // Close modal on outside click
        document.getElementById('productModal').addEventListener('click', function(e) {{
            if(e.target === this) closeModal();
        }});
        
        document.addEventListener('DOMContentLoaded', init);
    </script>
</body>
</html>
"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html_content)

print(f"Generated index.html with {len(productos)} products.")

$fotosDir = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Create new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Sheets.Item(1)
$worksheet.Name = "Catalogo"

# Headers
$worksheet.Cells.Item(1, 1).Value2 = "Marca"
$worksheet.Cells.Item(1, 2).Value2 = "Referencia"
$worksheet.Cells.Item(1, 3).Value2 = "Genero"
$worksheet.Cells.Item(1, 4).Value2 = "Familia_Olfativa"
$worksheet.Cells.Item(1, 5).Value2 = "Notas_Principales"
$worksheet.Cells.Item(1, 6).Value2 = "Inspiracion_Ventas"
$worksheet.Cells.Item(1, 7).Value2 = "Archivo_Foto"

$headerRange = $worksheet.Range("A1", "G1")
$headerRange.Font.Bold = $true

$files = Get-ChildItem -Path $fotosDir -Recurse -File -Filter "*.jpg"
$row = 2

foreach ($file in $files) {
    $marca = $file.Directory.Name
    $referencia = $file.BaseName -replace "_", " "
    $archivo = $file.Name
    
    $gen = "Unisex"
    $fam = "Por definir"
    $not = "Por definir"
    $tips = "Aroma característico de la casa."
    
    switch -Regex ($referencia) {
        "(?i)9am Dive" { 
            $gen = "Caballero"; $fam = "Aromática Acuática"; 
            $not = "Limón, menta, grosella negra, pimienta rosa, manzana, cedro, incienso, jengibre.";
            $tips = "Fresco y versátil (diurno). Estilo similar a Bleu de Chanel / YSL Y." 
        }
        "(?i)9pm Rebel" { 
            $gen = "Caballero"; $fam = "Aromática Frutal"; 
            $not = "Piña, manzana Granny Smith, mandarina, musgo de roble, vainilla, caramelo.";
            $tips = "Afrutado y tropical (piña) manteniendo el dulzor original." 
        }
        "(?i)9pm" { 
            $gen = "Caballero"; $fam = "Ámbar Vainilla"; 
            $not = "Manzana, canela, lavanda, bergamota, vainilla, haba tonka, ámbar.";
            $tips = "Nocturno, dulce y especiado. Alternativa directa a Ultra Male de JPG." 
        }
        "(?i)Cloud 2\.0" { 
            $gen = "Dama"; $fam = "Ámbar Vainilla"; 
            $not = "Pera, lavanda, praliné, crema de coco, maderas rubias, ambroxan.";
            $tips = "Mayor fijación y profundidad (más maderas/ambroxan) que el original." 
        }
        "(?i)Cloud Pink" { 
            $gen = "Dama"; $fam = "Ámbar Vainilla"; 
            $not = "Pitahaya, frutos rojos, agua de coco, orquídea de vainilla, praliné.";
            $tips = "Versión más frutal y tropical con pitahaya (fruta del dragón)." 
        }
        "(?i)Cloud" { 
            $gen = "Dama"; $fam = "Ámbar Vainilla"; 
            $not = "Lavanda, pera, bergamota, crema batida, coco, praliné.";
            $tips = "Dulzor cremoso popular. Comparte ADN con Baccarat Rouge 540." 
        }
        "(?i)Mod Blush" { 
            $gen = "Dama"; $fam = "Ámbar Amaderada"; 
            $not = "Maracuyá, bergamota, frambuesa, pimienta rosa, rosa, ambrox.";
            $tips = "Vibrante, fresca y afrutada con un secado limpio." 
        }
        "(?i)R\.?E\.?M" { 
            $gen = "Dama"; $fam = "Ámbar Vainilla"; 
            $not = "Caramelo, sal, higo, membrillo, lavanda, malvavisco, haba tonka.";
            $tips = "Inusual y relajante (lavanda dulce-salada con caramelo)." 
        }
        "(?i)Thank U Next" { 
            $gen = "Dama"; $fam = "Floral Frutal Gourmand"; 
            $not = "Pera, frambuesa, coco, rosa rosada, azúcar de macarrón.";
            $tips = "Juvenil y veraniego. Predominan notas de coco y repostería." 
        }
        "(?i)Club de Nuit Intense" { 
            $gen = "Caballero"; $fam = "Chipre Frutal"; 
            $not = "Limón, piña, bergamota, grosella negra, manzana, abedul, jazmín, ámbar gris.";
            $tips = "Clon legendario de Creed Aventus. Frutal y ahumado." 
        }
        "(?i)Club de Nuit Woman" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Naranja, bergamota, toronja, durazno, rosa, jazmín, lichi, pachulí.";
            $tips = "Alternativa precisa y de altísimo rendimiento a Coco Mademoiselle." 
        }
        "(?i)Club de Nuit Untold" { 
            $gen = "Unisex"; $fam = "Ámbar Amaderada"; 
            $not = "Azafrán, jazmín, amberwood, ámbar gris, cedro.";
            $tips = "Clon de altísima calidad de Baccarat Rouge 540." 
        }
        "(?i)Lionheart" { 
            $gen = "Caballero"; $fam = "Oriental Fougère"; 
            $not = "Menta, lavanda, vainilla, benjuí, miel, haba tonka y tabaco.";
            $tips = "Perfume oriental fougère, imponente y elegante." 
        }
        "(?i)Odyssey Mandarin Sky Elixir" { 
            $gen = "Caballero"; $fam = "Oriental Amaderada"; 
            $not = "Mandarina, naranja, lavanda, cardamomo, pimienta negra, caramelo, pachulí.";
            $tips = "Versión más especiada y profunda de Mandarin Sky." 
        }
        "(?i)Odyssey Mandarin Sky" { 
            $gen = "Caballero"; $fam = "Ámbar Amaderada"; 
            $not = "Mandarina, naranja, azafrán, salvia, caramelo, haba tonka, ambroxan.";
            $tips = "Contraste cítrico-dulce. Alternativa a Scandal Pour Homme de JPG." 
        }
        "(?i)Acqua di Gio Profondo" { 
            $gen = "Caballero"; $fam = "Aromática Acuática"; 
            $not = "Notas marinas, mandarina verde, bergamota, romero, lavanda, ciprés, notas minerales.";
            $tips = "El clásico renovado. Hiper-marino, mineral y aromático." 
        }
        "(?i)S[iì] Passione" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Pera, grosella negra, pimienta rosa, toronja, piña, rosa, jazmín, vainilla.";
            $tips = "Femenino, vibrante y frutal. Mucho más dulce que el 'Si' clásico." 
        }
        "(?i)King" { 
            $gen = "Caballero"; $fam = "Aromática Frutal"; 
            $not = "Naranja, bergamota, limón, notas frutales, vainilla, almizcle blanco y ámbar.";
            $tips = "Clon frutal dulce con ADN de Erba Pura (Bharara)." 
        }
        "(?i)Niche Femme" { 
            $gen = "Dama"; $fam = "Floral"; 
            $not = "Frutas, rosa turca, incienso, sándalo y vainilla.";
            $tips = "Perfil oriental-floral de alta concentración (aceites)." 
        }
        "(?i)212 VIP Black" { 
            $gen = "Caballero"; $fam = "Ámbar Fougère"; 
            $not = "Absenta, anís, hinojo, lavanda, cáscara de naranja, café, cuero, vainilla.";
            $tips = "Nocturno y de fiesta. Salida alcohólica (absenta) y dulce." 
        }
        "(?i)Very Good Girl" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Lichi, grosella roja, rosa, vainilla y vetiver.";
            $tips = "Rosa tropical y frutal (lichi). Menos oscuro que el original." 
        }
        "(?i)Good Girl" { 
            $gen = "Dama"; $fam = "Ámbar Floral"; 
            $not = "Almendra, café, limón, nardo, jazmín, cacao, vainilla, praliné, haba tonka.";
            $tips = "Sensual y nocturno. Icónico envase de tacón. Cacao y café dulce." 
        }
        "(?i)Coco Mademoiselle" { 
            $gen = "Dama"; $fam = "Ámbar Amaderada"; 
            $not = "Naranja, mandarina, bergamota, rosa, jazmín, pachulí, vetiver, vainilla.";
            $tips = "Estándar de elegancia. Cítricos, rosa y pachulí." 
        }
        "(?i)Aventus" { 
            $gen = "Caballero"; $fam = "Chipre Frutal"; 
            $not = "Grosella negra, bergamota, manzana, piña, rosa, abedul, jazmín, pachulí, ámbar gris.";
            $tips = "Pionero frutal-ahumado. Una de las fragancias más clonadas de la historia." 
        }
        "(?i)Silver Mountain Water" { 
            $gen = "Unisex"; $fam = "Aromática"; 
            $not = "Bergamota, mandarina, té verde, grosella negra, almizcle, sándalo.";
            $tips = "Fresco, metálico y verde. Inspirado en los Alpes suizos." 
        }
        "(?i)Jadore" { 
            $gen = "Dama"; $fam = "Floral Acuática"; 
            $not = "Limón, mandarina, romero, nenúfar, jazmín, magnolia, almizcle.";
            $tips = "Elegancia floral clásica con jazmín y notas afrutadas (pera)." 
        }
        "(?i)Miss Dior" { 
            $gen = "Dama"; $fam = "Floral Ámbar"; 
            $not = "Iris, peonía, lirio de los valles, rosa, durazno, vainilla, haba tonka.";
            $tips = "Romántico, empolvado y dulce (peonía, rosa, vainilla)." 
        }
        "(?i)Sauvage Parfum" { 
            $gen = "Caballero"; $fam = "Ámbar Fougère"; 
            $not = "Bergamota, mandarina, sándalo, incienso de olíbano, haba tonka, vainilla.";
            $tips = "El rey de los cumplidos en versión densa y amaderada." 
        }
        "(?i)Sauvage" { 
            $gen = "Caballero"; $fam = "Aromática Fougère"; 
            $not = "Bergamota, pimienta, lavanda, geranio, pachulí, ambroxan, cedro.";
            $tips = "Fresco, especiado, con gran proyección." 
        }
        "(?i)Light Blue" { 
            $gen = "Dama"; $fam = "Floral Frutal / Amaderada"; 
            $not = "Limón siciliano, manzana verde, cedro, campanilla, jazmín, bambú.";
            $tips = "Cítrico icónico. Referencia absoluta para el clima cálido/verano." 
        }
        "(?i)Boss Bottled Parfum" { 
            $gen = "Caballero"; $fam = "Amaderada Especiada"; 
            $not = "Incienso, mandarina, higo, iris, cedro, cuero.";
            $tips = "Versión oscura, profunda y formal del clásico Boss Bottled." 
        }
        "(?i)Boss Bottled" { 
            $gen = "Caballero"; $fam = "Amaderada Especiada"; 
            $not = "Manzana, ciruela, bergamota, limón, canela, clavo, caoba, vainilla, vetiver.";
            $tips = "Clásico versátil. El aroma por excelencia para la oficina (manzana/canela)." 
        }
        "(?i)Il Femme" { 
            $gen = "Dama"; $fam = "Aromática / Oriental Amaderada"; 
            $not = "Vainilla, rosa, notas polvorientas, maracuyá, nuez moscada, canela, cedro.";
            $tips = "Nicho colombiano. Perfil oriental lujoso de gran proyección." 
        }
        "(?i)Il Roso" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Bergamota, ruibarbo, rosa, jazmín, nuez moscada, ámbar, incienso, vainilla.";
            $tips = "Protagonismo absoluto de la rosa combinada con maderas ahumadas." 
        }
        "(?i)Oud For Greatness" { 
            $gen = "Unisex"; $fam = "Amaderada Ámbar"; 
            $not = "Oud, azafrán, nuez moscada, lavanda, pachulí y almizcle.";
            $tips = "Nicho de alto impacto. Oscuro, amaderado (oud/azafrán) y opulento." 
        }
        "(?i)Le Beau" { 
            $gen = "Caballero"; $fam = "Amaderada Aromática"; 
            $not = "Bergamota, madera de coco y haba tonka.";
            $tips = "Tropical y seductor. Ideal clima cálido (coco, tonka, bergamota)." 
        }
        "(?i)Scandal" { 
            $gen = "Dama"; $fam = "Chipre Floral Gourmand"; 
            $not = "Naranja sanguina, miel, gardenia, flor de azahar, jazmín, cera de abeja, caramelo.";
            $tips = "Extremadamente dulce (miel/caramelo), denso y nocturno." 
        }
        "(?i)L12\.?12 Blanc" { 
            $gen = "Caballero"; $fam = "Amaderada Aromática"; 
            $not = "Toronja, romero, cardamomo, nardo, cedro, ante y vetiver.";
            $tips = "Fresco y deportivo. Huele a algodón blanco recién lavado." 
        }
        "(?i)Lacoste Rouge" { 
            $gen = "Caballero"; $fam = "Amaderada Especiada"; 
            $not = "Manzana roja, pino, cardamomo, pimienta, canela, pachulí, madera.";
            $tips = "Energético y especiado. Manzana roja sobre maderas picantes." 
        }
        "(?i)La vie est belle" { 
            $gen = "Dama"; $fam = "Floral Gourmand"; 
            $not = "Grosella negra, pera, iris, jazmín, flor de azahar, praliné, vainilla, pachulí.";
            $tips = "Súper ventas dulzón (praliné/vainilla). Alta duración, empolvado." 
        }
        "(?i)Afeef" { 
            $gen = "Dama"; $fam = "Floral Amaderada"; 
            $not = "Durazno, pimienta rosa, bergamota, nardo, flor de azahar, jazmín, praliné.";
            $tips = "Uso diario, maderas limpias y frutas en la salida." 
        }
        "(?i)Asad Bourbon" { 
            $gen = "Caballero"; $fam = "Oriental Especiada"; 
            $not = "Lavanda, mirabel, pimienta rosa, cacao, nuez moscada, vainilla bourbon.";
            $tips = "Novedad dulce/especiada con potente vainilla bourbon." 
        }
        "(?i)Asad" { 
            $gen = "Caballero"; $fam = "Ámbar"; 
            $not = "Pimienta negra, piña, tabaco, pachulí, café, vainilla, ámbar, maderas secas.";
            $tips = "Alternativa económica a Dior Sauvage Elixir. Potente y especiado." 
        }
        "(?i)Badee Al Oud Amethyst" { 
            $gen = "Unisex"; $fam = "Floral Ámbar"; 
            $not = "Pimienta rosa, bergamota, rosa turca, rosa búlgara, jazmín, oud, ámbar y vainilla.";
            $tips = "Clon de Atomic Rose (Initio). Rosa dominante, dulce y amaderado." 
        }
        "(?i)Badee Al Oud Honor" { 
            $gen = "Unisex"; $fam = "Oriental Gourmand / Especiada"; 
            $not = "Piña brûlée, cúrcuma, canela, pimienta negra, benjuí, vainilla, cashmeran.";
            $tips = "Gourmand tropical único (acorde de piña caramelizada/creme brulee)." 
        }
        "(?i)Badee Al Oud Oud for Glory" { 
            $gen = "Unisex"; $fam = "Ámbar Amaderada"; 
            $not = "Azafrán, nuez moscada, lavanda, oud, pachulí y almizcle.";
            $tips = "Clon exacto de Oud for Greatness. Opulento, azafrán y oud." 
        }
        "(?i)Badee Al Oud Sublime" { 
            $gen = "Unisex"; $fam = "Amaderada Aromática"; 
            $not = "Manzana, lichi, rosa, ciruela, jazmín, musgo, vainilla y pachulí.";
            $tips = "Clon de Eden Juicy Apple (Kayali). Afrutado, jugoso (manzana/ciruela)." 
        }
        "(?i)Eclaire" { 
            $gen = "Unisex"; $fam = "Floral Frutal Gourmand"; 
            $not = "Caramelo, leche, azúcar, miel, flores blancas, vainilla, praliné y almizcle.";
            $tips = "Clon exacto de Bianco Latte. Extremadamente lactónico, caramelo y miel." 
        }
        "(?i)His Confession" { 
            $gen = "Caballero"; $fam = "Oriental Amaderada"; 
            $not = "Mandarina, canela, lavanda, iris, benjuí, vainilla, haba tonka, incienso.";
            $tips = "Elegante. Se posiciona como Armani Code Absolu / YSL Tuxedo." 
        }
        "(?i)Khamrah Dukhan" { 
            $gen = "Unisex"; $fam = "Oriental"; 
            $not = "Especias, mandarina, incienso, ládano, flor de azahar, praliné, tabaco, ámbar.";
            $tips = "Notas ahumadas y de incienso sobre la base dulce original." 
        }
        "(?i)Khamrah Qahwa" { 
            $gen = "Unisex"; $fam = "Aromática Especiada"; 
            $not = "Jengibre, canela, cardamomo, praliné, café arábica, haba tonka, vainilla.";
            $tips = "Café tostado que equilibra el dulzor. Flanker muy exitoso." 
        }
        "(?i)Khamrah" { 
            $gen = "Unisex"; $fam = "Oriental Especiada"; 
            $not = "Canela, nuez moscada, dátiles, praliné, nardo, vainilla, haba tonka, ámbar.";
            $tips = "Inspirado en Angels' Share (Kilian). Dátiles dulces y especias cálidas." 
        }
        "(?i)Yara Candy" { 
            $gen = "Dama"; $fam = "Frutal Gourmand"; 
            $not = "Mandarina verde, grosella negra, caramelo de fresa, gardenia, jarabe de vainilla.";
            $tips = "Emula dulces frutales y confitería (estilo Oriana PDM)." 
        }
        "(?i)Yara Tous" { 
            $gen = "Dama"; $fam = "Floral Tropical Gourmand"; 
            $not = "Mango, coco, maracuyá, jazmín, flor de azahar, vainilla, almizcle.";
            $tips = "El Yara del verano. Coco y mango, vibra de protector solar de lujo." 
        }
        "(?i)Yara" { 
            $gen = "Dama"; $fam = "Ámbar Vainilla"; 
            $not = "Mandarina, heliotropo, frutas tropicales, acorde gourmand, vainilla, sándalo.";
            $tips = "Súper viral. Huele a batido de fresa, dulce, atalcado y cremoso." 
        }
        "(?i)Santal 33" { 
            $gen = "Unisex"; $fam = "Amaderada Aromática"; 
            $not = "Cardamomo, violeta, iris, papiro, sándalo, cedro, cuero, ámbar y almizcle.";
            $tips = "Fragancia nicho de estatus. Seca, amaderada (sándalo), atalcada/cuero." 
        }
        "(?i)Arabians Tonka" { 
            $gen = "Unisex"; $fam = "Oriental Amaderada"; 
            $not = "Bergamota, azafrán, rosa, oud, ámbar, musgo de roble, haba tonka, azúcar moreno.";
            $tips = "'Beast mode' (rendimiento nuclear). Azúcar quemada, rosa y oud intenso." 
        }
        "(?i)Starry Nights" { 
            $gen = "Unisex"; $fam = "Oriental Floral"; 
            $not = "Bergamota, limón, rosa, jazmín, pachulí, ámbar, vainilla, almizcle blanco.";
            $tips = "Frutas frescas en la salida y un secado denso de rosa y pachulí." 
        }
        "(?i)Toy 2 Pearl" { 
            $gen = "Unisex"; $fam = "Cítrica Floral Amaderada"; 
            $not = "Limón, orégano, jazmín, arena, fresia, vetiver, ciprés y almizcle.";
            $tips = "Mineral y salado. Simula arena mezclada con cítricos italianos." 
        }
        "(?i)Toy 2" { 
            $gen = "Dama"; $fam = "Floral Amaderada Almizclada"; 
            $not = "Manzana, mandarina, magnolia, grosella blanca, peonía, jazmín, sándalo.";
            $tips = "Aroma a limpio supremo. Gel de ducha de lujo o sábanas limpias." 
        }
        "(?i)Toy Boy 2" { 
            $gen = "Caballero"; $fam = "Amaderada Especiada"; 
            $not = "Nuez moscada, jengibre, Akigalawood, café, mirra y vetiver.";
            $tips = "Versión alternativa y especiada de Toy Boy." 
        }
        "(?i)Amber Rouge" { 
            $gen = "Unisex"; $fam = "Amaderada Especiada"; 
            $not = "Azafrán, jazmín, amberwood, ámbar gris, cedro y resina de abeto.";
            $tips = "Clon de altísima calidad de Baccarat Rouge 540 Extrait." 
        }
        "(?i)Royal Amber" { 
            $gen = "Unisex"; $fam = "Oriental Amaderada"; 
            $not = "Canela, naranja, bálsamo de Gurjun, orquídea negra, pachulí y sándalo.";
            $tips = "Clon premium de Erba Pura (Xerjoff). Explosión frutal y almizcle." 
        }
        "(?i)Velvet Gold" { 
            $gen = "Dama"; $fam = "Oriental Suave / Gourmand Almizclada"; 
            $not = "Bergamota, caramelo, violeta, pimienta rosa, notas empolvadas, pachulí, rosa, vainilla.";
            $tips = "Alternativa a Gentle Fluidity Gold (MFK). Dulce, atalcada y metálica." 
        }
        "(?i)1 Million Lucky" { 
            $gen = "Caballero"; $fam = "Amaderada"; 
            $not = "Ciruela, toronja, avellana, miel, cedro, flor de azahar, jazmín, pachulí.";
            $tips = "Muy codiciado. Destaca el acorde único de avellana con miel." 
        }
        "(?i)Fame Blooming Pink" { 
            $gen = "Dama"; $fam = "Almizclada Floral Amaderada"; 
            $not = "Mango, bergamota, jazmín, incienso de olíbano, vainilla y sándalo.";
            $tips = "ADN tropical dominado por el mango jugoso, con jazmín e incienso." 
        }
        "(?i)Fame" { 
            $gen = "Dama"; $fam = "Almizclada Floral Amaderada"; 
            $not = "Mango, bergamota, jazmín, incienso de olíbano, vainilla y sándalo.";
            $tips = "ADN tropical dominado por el mango jugoso, con jazmín e incienso." 
        }
        "(?i)Invictus" { 
            $gen = "Caballero"; $fam = "Amaderada Acuática"; 
            $not = "Toronja, notas marinas, mandarina, jazmín, ámbar gris, madera de guayaco, pachulí.";
            $tips = "Fundador de lo dulce/acuático. Sintético, deportivo, gran rendimiento." 
        }
        "(?i)Phantom" { 
            $gen = "Caballero"; $fam = "Aromática Fougère"; 
            $not = "Limón, lavanda, manzana, pachulí, tierra, vainilla, vetiver.";
            $tips = "Lavanda tradicional mezclada con tecnología y dulzor denso (terroso)." 
        }
        "(?i)Delina Exclusif" { 
            $gen = "Dama"; $fam = "Ámbar Floral"; 
            $not = "Bergamota, pera, lichi, rosa damascena, incienso, vetiver, vainilla, almizcle.";
            $tips = "Versión más densa, empolvada y dulce de la icónica rosa de Delina." 
        }
        "(?i)Kalan" { 
            $gen = "Caballero"; $fam = "Oriental Especiada"; 
            $not = "Naranja sanguina, pimienta negra, especias, lavanda, ajenjo, flor de azahar, cedro.";
            $tips = "Nicho polarizante. Terroso, seco, punzante (naranja sanguina y pimienta)." 
        }
        "(?i)Marshmallow Blush" { 
            $gen = "Dama"; $fam = "Floral Gourmand"; 
            $not = "Bergamota, toronja, frutas, coco, nardo, malvavisco, vainilla, almizcle.";
            $tips = "Extremadamente dulce, centrado exclusivamente en malvavisco/azúcar." 
        }
        "(?i)Rose Rush" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Pétalos de rosa, lichi, neroli, rosa de mayo, peonía, papaya, almizcle blanco.";
            $tips = "Accesible, fresca, acuática y centrada en una rosa juvenil." 
        }
        "(?i)360" { 
            $gen = "Dama"; $fam = "Floral"; 
            $not = "Melón, lirio, osmanto, mandarina, rosa, nenúfar, salvia, sándalo, vainilla.";
            $tips = "Clásico de los 90. Perfil altamente pulcro, limpio y floral/acuático." 
        }
        "(?i)Paradoxe Intense" { 
            $gen = "Dama"; $fam = "Floral Ámbar Amaderada"; 
            $not = "Neroli, Ambrofix y acorde de musgo.";
            $tips = "Sobredosis de flores blancas (neroli) con musgo para profundidad." 
        }
        "(?i)Now Women" { 
            $gen = "Dama"; $fam = "Floral Aldehídica"; 
            $not = "Té blanco, maracuyá, bergamota, aldehídos, flor de tiaré, jazmín, almizcle.";
            $tips = "Alternativa económica a Burberry Her. Fresa sintética atalcada." 
        }
        "(?i)Tobacco Vanille" { 
            $gen = "Unisex"; $fam = "Ámbar Especiada"; 
            $not = "Hoja de tabaco, especias, vainilla, cacao, haba tonka, frutos secos y maderas.";
            $tips = "Obra maestra. Picante del tabaco húmedo y dulzor resinoso (vainilla)." 
        }
        "(?i)Tommy" { 
            $gen = "Caballero"; $fam = "Cítrica Aromática"; 
            $not = "Bergamota, menta, lavanda, manzana, arándano, rosa, jazmín, algodón, maderas.";
            $tips = "Icono fresco y casual americano de los 90. Menta y manzana." 
        }
        "(?i)Uomo Born In Roma" { 
            $gen = "Caballero"; $fam = "Amaderada Aromática / Mineral"; 
            $not = "Hojas de violeta, sal, jengibre, salvia, maderas y vetiver.";
            $tips = "Secado moderno y versátil. Juega con acordes minerales y dulzor fresco." 
        }
        "(?i)Spicebomb" { 
            $gen = "Caballero"; $fam = "Amaderada Especiada"; 
            $not = "Pimienta rosa, bergamota, toronja, canela, azafrán, cuero, tabaco y vetiver.";
            $tips = "'Bomba de especias'. Perfecto para el frío, picante y cálido a la vez." 
        }
        "(?i)Libre" { 
            $gen = "Dama"; $fam = "Ámbar Fougère"; 
            $not = "Lavanda, mandarina, bergamota, flor de azahar, jazmín, vainilla, haba tonka.";
            $tips = "Toma la lavanda masculina y la feminiza con flor de azahar y vainilla." 
        }
        "(?i)Ajwad" { 
            $gen = "Unisex"; $fam = "Amaderada Ámbar"; 
            $not = "Frutas, canela, bergamota, rosa, jazmín, ámbar, vainilla, almizcle, cedro, oud.";
            $tips = "Rendimiento extremo. Dulzor frutal metálico y rosa avainillada." 
        }
        "(?i)Victoria" { 
            $gen = "Dama"; $fam = "Floral Frutal"; 
            $not = "Flores silvestres, vainilla, sándalo.";
            $tips = "Aroma dulce y versátil femenino de Lattafa." 
        }
        "(?i)Ombre Nomade" { 
            $gen = "Unisex"; $fam = "Oriental Amaderada"; 
            $not = "Oud, rosa, incienso, frambuesa, azafrán, ámbar, benjuí, abedul y geranio.";
            $tips = "Opulencia pura. Cuero oscuro, oud potente y frambuesa de altísima estela." 
        }
        "(?i)Odyssey Collection" {
            $gen = "Unisex"; $fam = "Colección";
            $not = "Depende de la fragancia específica (Spectra, Mandarin, Mega, etc).";
            $tips = "Línea completa de Armaf con enfoques frescos, cítricos y frutales."
        }
    }
    
    $worksheet.Cells.Item($row, 1).Value2 = $marca
    $worksheet.Cells.Item($row, 2).Value2 = $referencia
    $worksheet.Cells.Item($row, 3).Value2 = $gen
    $worksheet.Cells.Item($row, 4).Value2 = $fam
    $worksheet.Cells.Item($row, 5).Value2 = $not
    $worksheet.Cells.Item($row, 6).Value2 = $tips
    $worksheet.Cells.Item($row, 7).Value2 = $archivo
    
    $row++
}

$worksheet.Columns.AutoFit() | Out-Null
if (Test-Path $xlsx_path) { Remove-Item $xlsx_path }
$workbook.SaveAs($xlsx_path)
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

Write-Host "Base de datos reconstruida de 0."

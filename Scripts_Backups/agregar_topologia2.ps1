$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.Cells.SpecialCells(11).Row

for ($row = 2; $row -le $lastRow; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $top = ""
    
    switch -Regex ($ref) {
        "(?i)^9am Dive$" { $top = "Clasificación: Inspiración / Híbrido`nDetalle: Toma el framework aromático-acuático de Bleu de Chanel y fusiona variables de YSL Y. No es una réplica exacta, sino una iteración optimizada." }
        "(?i)^9pm$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Ingeniería inversa directa y fiel a la fórmula de Ultra Male de Jean Paul Gaultier." }
        "(?i)^9pm Rebel$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Bifurcación (fork) del código base de 9pm. Añade variables frutales (piña) para acercarse a perfiles tipo Aventus o Phantom." }
        "(?i)^Cloud$" { $top = "Clasificación: Pilar / Inspiración`nDetalle: Es el código fuente original de su propia línea, pero su arquitectura técnica está fuertemente inspirada en el perfil dulce/medicinal de Baccarat Rouge 540 (MFK)." }
        "(?i)^Cloud 2\.0 Intense$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Optimización de concentración del Cloud original; inyecta mayor carga de ambroxan y maderas." }
        "(?i)^Cloud Pink$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Derivación del Cloud original. Sustituye la nota de lavanda por fruta del dragón." }
        "(?i)^Mod Blush$" { $top = "Clasificación: Pilar`nDetalle: Código base original e independiente dentro de la marca." }
        "(?i)^R\.?E\.?M(_Alt)?$" { $top = "Clasificación: Pilar`nDetalle: Código base original centrado en el acorde lavanda-sal-caramelo." }
        "(?i)^Thank U Next$" { $top = "Clasificación: Pilar`nDetalle: Código base original centrado en repostería y coco." }
        "(?i)^Club de Nuit Intense$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica técnica de Creed Aventus." }
        "(?i)^Club de Nuit Woman$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica técnica de alto rendimiento de Coco Mademoiselle de Chanel." }
        "(?i)^Club de Nuit Untold$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica técnica de Baccarat Rouge 540." }
        "(?i)^Lionheart(_Alt)?$" { $top = "Clasificación: Equivalencia`nDetalle: Contratipo genérico de bajo coste comercializado por marcas de catálogo (ej. Dorall Collection), usualmente sin arquitectura propia." }
        "(?i)^Odyssey Mandarin Sky$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Ingeniería inversa de Scandal Pour Homme de Jean Paul Gaultier." }
        "(?i)^Odyssey Mandarin Sky Elixir$" { $top = "Clasificación: No verificable`nDetalle: Ausencia de registros oficiales en la base de datos global de Armaf para la terminación 'Elixir' en esta línea específica." }
        "(?i)^Odyssey Collection$" { $top = "Clasificación: Framework / Línea base`nDetalle: No es un perfume individual, sino el ecosistema donde se alojan múltiples derivaciones y dupes de la marca." }
        "(?i)^Yum Yum$" { $top = "Clasificación: Dupe / Inspiración`nDetalle: Si es de marca árabe, es una réplica técnica diseñada para emular a Yum Pistachio Gelato | 33 de Kayali." }
        "(?i)^Acqua di Gio Profondo$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Actualización marina e inyección de notas minerales (Aquozone) al código pilar del Acqua di Gio de 1996." }
        "(?i)^S[iì] Passione$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Derivación frutal y vibrante del pilar original Armani Si." }
        "(?i)^King$" { $top = "Clasificación: Ambigüedad de nodo`nDetalle: Si es Bharara King, es una Inspiración/Dupe que utiliza la estructura base de Erba Pura de Xerjoff." }
        "(?i)^Niche Femme$" { $top = "Clasificación: Inspiración`nDetalle: Perfil construido sobre frameworks de perfumería oriental nicho, sin apuntar a una réplica 1:1 específica." }
        "(?i)^Scarlet$" { $top = "Clasificación: Inspiración / Dupe`nDetalle: Si refiere a Bharara Scarlet, es una iteración que clona el código frutal-almizclado de referencias nicho como Erba Pura o Kirke." }
        "(?i)^212 VIP Black$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Bifurcación nocturna (absenta y vainilla) del pilar 212 VIP Men." }
        "(?i)^Good Girl$" { $top = "Clasificación: Pilar`nDetalle: Código fuente original (envase tacón) y base de toda su línea de flankers." }
        "(?i)^Very Good Girl$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Deriva de Good Girl, pero modifica su núcleo hacia notas de lichi y rosa (arquitectura inspirada por Delina de PDM)." }
        "(?i)^Coco Mademoiselle$" { $top = "Clasificación: Pilar (Evolutivo)`nDetalle: Nació como derivación de Coco, pero por su impacto en la industria opera actualmente como un código fuente independiente." }
        "(?i)^Aventus$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: El estándar de oro de la industria. Código base (piña ahumada) que originó la mayor cantidad de clones y dupes en el mercado moderno." }
        "(?i)^Silver Mountain Water$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Arquitectura original de Creed; fresco, metálico y verde." }
        "(?i)^Jadore Marina$" { $top = "Clasificación: Error de sintaxis (Null)`nDetalle: Nomenclatura inexistente. Posible confusión con flankers oficiales como J'adore In Joy." }
        "(?i)^Miss Dior$" { $top = "Clasificación: Pilar`nDetalle: Código base con múltiples reformulaciones oficiales a lo largo de las décadas." }
        "(?i)^Sauvage Parfum$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Iteración de mayor concentración de aceites y resinas (olíbano) basada en el pilar Sauvage EDT." }
        "(?i)^Sauvage$" { $top = "Clasificación: Pilar`nDetalle: Framework comercial moderno." }
        "(?i)^Light Blue$" { $top = "Clasificación: Pilar`nDetalle: Estructura cítrica fundacional." }
        "(?i)^Boss Bottled Parfum$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Densificación del pilar original mediante notas de cuero e incienso." }
        "(?i)^Boss Bottled$" { $top = "Clasificación: Pilar`nDetalle: Diseño clásico centrado en manzana y canela." }
        "(?i)^Il Femme$" { $top = "Clasificación: Pilar (Nicho Latino)`nDetalle: Opera con patente propia, fuertemente inspirado en protocolos olfativos de medio oriente." }
        "(?i)^Il Roso$" { $top = "Clasificación: Pilar (Nicho Latino)`nDetalle: Estructura independiente dentro de la marca centrada en la rosa turca." }
        "(?i)^Oud For Greatness(_Alt)?$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Código fuente de azafrán, oud y pachulí. Es el objetivo principal de múltiples casas clonadoras árabes." }
        "(?i)^Le Beau$" { $top = "Clasificación: Flanker (Derivación autónoma)`nDetalle: Se ramifica de la familia Le Male, pero establece un entorno completamente nuevo basado en el coco." }
        "(?i)^Scandal$" { $top = "Clasificación: Pilar`nDetalle: Código base original (perfil gourmand extremo)." }
        "(?i)^L12\.?12 Blanc$" { $top = "Clasificación: Pilar`nDetalle: Código base fundacional de la línea L1212 (algodón limpio)." }
        "(?i)^Lacoste Rouge$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Iteración especiada (manzana roja, té rooibos) del marco L1212." }
        "(?i)^La vie est belle$" { $top = "Clasificación: Pilar`nDetalle: Estándar moderno de la perfumería floral-gourmand." }
        "(?i)^Afeef$" { $top = "Clasificación: Inspiración`nDetalle: Usa arquitecturas de diseñador occidental (fresco amaderado) sin ser un clon 1:1." }
        "(?i)^Asad Bourbon$" { $top = "Clasificación: No verificable / Error de sintaxis`nDetalle: La base de datos de Lattafa solo registra Asad y el flanker Asad Zanzibar." }
        "(?i)^Asad$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Ingeniería inversa de alta precisión orientada a replicar Sauvage Elixir (Dior)." }
        "(?i)^Badee Al Oud Amethyst$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica técnica de Atomic Rose (Initio Parfums Privés)." }
        "(?i)^Badee Al Oud Honor and Glory White$" { $top = "Clasificación: Inspiración`nDetalle: Toma elementos del marco de TriBeCa de Bond No. 9 y lo fusiona con un acorde de piña acaramelada. No es un clon 1:1." }
        "(?i)^Badee Al Oud Oud for Glory Black$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica técnica y directa de Oud for Greatness (Initio)." }
        "(?i)^Badee Al Oud Sublime$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica orientada a copiar Eden Juicy Apple | 01 de Kayali." }
        "(?i)^Eclaire$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Copia exacta de la formulación hiper-lactónica de Bianco Latte (Giardini Di Toscana)." }
        "(?i)^His Confession$" { $top = "Clasificación: Dupe / Híbrido`nDetalle: Ingeniería inversa que busca revivir la arquitectura de descontinuados como Armani Code Absolu y YSL Tuxedo." }
        "(?i)^Khamrah Dukhan$" { $top = "Clasificación: No verificable / Error de sintaxis`nDetalle: Sólo existen registros oficiales de Khamrah y Khamrah Qahwa." }
        "(?i)^Khamrah Qahwa$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Modifica el código original de Khamrah añadiendo un nodo de café tostado amargo." }
        "(?i)^Khamrah$" { $top = "Clasificación: Inspiración`nDetalle: Basado en el marco de Angels' Share de Kilian, pero modifica variables clave (elimina el coñac, inyecta dátiles). No es un Dupe 1:1." }
        "(?i)^Yara Candy$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Derivación del Yara original; iteración dirigida a emular dulces de confitería." }
        "(?i)^Yara Tous$" { $top = "Clasificación: Flanker (Derivación) / Inspiración`nDetalle: Derivación tropical; toma inspiraciones estructurales de Alien Goddess de Mugler y Fame de Paco Rabanne." }
        "(?i)^Yara$" { $top = "Clasificación: Pilar / Inspiración`nDetalle: Pilar comercial de Lattafa, pero su estructura toma prestados protocolos del Pepe Jeans for Her original." }
        "(?i)^Santal 33$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Código maestro original (sándalo, papiro, cuero). Extremadamente imitado." }
        "(?i)^Arabians Tonka$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Código fuente de rendimiento extremo (azúcar, oud, haba tonka)." }
        "(?i)^Starry Nights$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Estructura independiente dentro del ecosistema Montale." }
        "(?i)^Toy 2 Pearl$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Modificación del framework original introduciendo notas minerales y saladas." }
        "(?i)^Toy 2$" { $top = "Clasificación: Pilar`nDetalle: Código base independiente centrado en manzana y almizcle blanco." }
        "(?i)^Toy Boy 2$" { $top = "Clasificación: No verificable / Error de sintaxis`nDetalle: Moschino solo ha desplegado el pilar original Toy Boy en esta sub-línea." }
        "(?i)^Amber Rouge$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica en hardware premium del Baccarat Rouge 540 Extrait." }
        "(?i)^Royal Amber$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Clon de alto rendimiento de Erba Pura de Xerjoff." }
        "(?i)^Velvet Gold$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Réplica de Gentle Fluidity Gold de Maison Francis Kurkdjian (MFK)." }
        "(?i)^1 Million Lucky$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Alteración agresiva del 1 Million original, introduciendo variables de ciruela y avellana." }
        "(?i)^Fame( Blooming Pink)?$" { $top = "Clasificación: Flanker / Edición Especial`nDetalle: Modificación estética/comercial del pilar Fame." }
        "(?i)^Invictus$" { $top = "Clasificación: Pilar`nDetalle: Arquitectura original que definió el género 'dulce/acuático' moderno." }
        "(?i)^Phantom$" { $top = "Clasificación: Pilar`nDetalle: Código original basado en la integración de cítricos, lavanda y vainilla terrosa." }
        "(?i)^Delina Exclusif$" { $top = "Clasificación: Flanker (Derivación Nicho)`nDetalle: Densificación (oud e incienso) del pilar original Delina." }
        "(?i)^Kalan$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Estructura independiente de especias rojas y naranja sanguina." }
        "(?i)^Marshmallow Blush$" { $top = "Clasificación: Equivalencia / Body Mist`nDetalle: Formulación ultra simple (bruma corporal genérica) sin pirámide olfativa escalonada." }
        "(?i)^Rose Rush$" { $top = "Clasificación: Flanker / Pilar`nDetalle: Parte de la línea 'Rush', operando con código propio." }
        "(?i)^360$" { $top = "Clasificación: Pilar`nDetalle: Clásico original de los 90." }
        "(?i)^Paradoxe Intense$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Incremento de variables florales (neroli) y musgo frente al Paradoxe base." }
        "(?i)^Now Women$" { $top = "Clasificación: Dupe (Clon exacto)`nDetalle: Ingeniería inversa que emula el despliegue de Burberry Her EDP." }
        "(?i)^Tobacco Vanille$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Obra maestra original de la línea Private Blend." }
        "(?i)^Tommy$" { $top = "Clasificación: Pilar`nDetalle: Código fuente fundacional de la casa." }
        "(?i)^Uomo Born In Roma( Set)?$" { $top = "Clasificación: Flanker (Derivación)`nDetalle: Derivación modernizada (notas minerales/dulces) del Valentino Uomo original." }
        "(?i)^Spicebomb$" { $top = "Clasificación: Pilar`nDetalle: Código original de especias y tabaco." }
        "(?i)^Libre$" { $top = "Clasificación: Pilar`nDetalle: Arquitectura original que redefinió la lavanda en el entorno femenino comercial." }
        "(?i)^Ajwad$" { $top = "Clasificación: Inspiración / Híbrido`nDetalle: Toma el núcleo floral-dulce de Roses Vanille (Mancera) y le inyecta una sobredosis de fruta metálica." }
        "(?i)^Victoria$" { $top = "Clasificación: No verificable`nDetalle: Sintaxis ambigua. Requiere el nombre de la casa matriz." }
        "(?i)^Ombre Nomade$" { $top = "Clasificación: Pilar (Nicho)`nDetalle: Fórmula original y exclusiva de la casa (Oud, rosa, incienso)." }
        "(?i)^Art of Universe$" { $top = "Clasificación: No verificable`nDetalle: Carente de documentación técnica pública en el ecosistema perfumista oficial." }
    }
    
    if ($top) {
        $worksheet.Cells.Item($row, 6).Value2 = $top
    }
}

$worksheet.Columns.Item(6).ColumnWidth = 60
$worksheet.Columns.Item(6).WrapText = $true

$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Topologia agregada exitosamente."

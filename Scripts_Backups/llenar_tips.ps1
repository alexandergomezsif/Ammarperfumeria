$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)

$lastRow = $worksheet.Cells.SpecialCells(11).Row

$worksheet.Cells.Item(1, 8).Value2 = "Tips_Ventas"

for ($row = 2; $row -le $lastRow; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    $tips = ""
    
    switch -Regex ($ref) {
        "(?i)9am Dive" { $tips = "Fresco y versátil (diurno). Estilo similar a Bleu de Chanel / YSL Y." }
        "(?i)9pm Rebel" { $tips = "Afrutado y tropical (piña) manteniendo el dulzor original." }
        "(?i)9pm" { $tips = "Nocturno, dulce y especiado. Alternativa directa a Ultra Male de JPG." }
        "(?i)Cloud Pink" { $tips = "Versión más frutal y tropical con pitahaya (fruta del dragón)." }
        "(?i)Cloud 2.0" { $tips = "Mayor fijación y profundidad (más maderas/ambroxan) que el original." }
        "(?i)Cloud" { $tips = "Dulzor cremoso popular. Comparte ADN con Baccarat Rouge 540." }
        "(?i)Mod Blush" { $tips = "Vibrante, fresca y afrutada con un secado limpio." }
        "(?i)R\.E\.M" { $tips = "Inusual y relajante (lavanda dulce-salada con caramelo)." }
        "(?i)Thank U Next" { $tips = "Juvenil y veraniego. Predominan notas de coco y repostería." }
        "(?i)Club de Nuit Intense" { $tips = "Clon legendario de Creed Aventus. Frutal y ahumado." }
        "(?i)Club de Nuit Woman" { $tips = "Alternativa precisa y de altísimo rendimiento a Coco Mademoiselle." }
        "(?i)Club de Nuit Untold" { $tips = "Clon de altísima calidad de Baccarat Rouge 540." }
        "(?i)Lionheart" { $tips = "Perfume oriental fougère (menta, lavanda, miel, tabaco)." }
        "(?i)Odyssey Mandarin" { $tips = "Contraste cítrico-dulce. Alternativa a Scandal Pour Homme de JPG." }
        "(?i)Acqua di Gio Profondo" { $tips = "El clásico renovado. Hiper-marino, mineral y aromático." }
        "(?i)Si Passione" { $tips = "Femenino, vibrante y frutal. Mucho más dulce que el 'Si' clásico." }
        "(?i)King" { $tips = "Clon frutal dulce con ADN de Erba Pura (si es Bharara)." }
        "(?i)Niche Femme" { $tips = "Perfil oriental-floral de alta concentración (aceites)." }
        "(?i)212 VIP Black" { $tips = "Nocturno y de fiesta. Salida alcohólica (absenta) y dulce." }
        "(?i)Very Good Girl" { $tips = "Rosa tropical y frutal (lichi). Menos oscuro que el original." }
        "(?i)Good Girl" { $tips = "Sensual y nocturno. Icónico envase de tacón. Cacao y café dulce." }
        "(?i)Coco Mademoiselle" { $tips = "Estándar de elegancia. Cítricos, rosa y pachulí." }
        "(?i)Aventus" { $tips = "Pionero frutal-ahumado. Una de las fragancias más clonadas de la historia." }
        "(?i)Silver Mountain" { $tips = "Fresco, metálico y verde. Inspirado en los Alpes suizos." }
        "(?i)Jadore" { $tips = "Elegancia floral clásica con jazmín y notas afrutadas (pera)." }
        "(?i)Miss Dior" { $tips = "Romántico, empolvado y dulce (peonía, rosa, vainilla)." }
        "(?i)Sauvage" { $tips = "El rey de los cumplidos. Fresco, especiado, con gran proyección." }
        "(?i)Light Blue" { $tips = "Cítrico icónico. Referencia absoluta para el clima cálido/verano." }
        "(?i)Boss Bottled" { $tips = "Clásico versátil. El aroma por excelencia para la oficina (manzana/canela)." }
        "(?i)Il Femme" { $tips = "Nicho colombiano. Perfil oriental lujoso (rosa/nuez moscada) de gran proyección." }
        "(?i)Il Roso" { $tips = "Protagonismo absoluto de la rosa combinada con maderas ahumadas." }
        "(?i)Oud For Greatness" { $tips = "Nicho de alto impacto. Oscuro, amaderado (oud/azafrán) y opulento." }
        "(?i)Le Beau" { $tips = "Tropical y seductor. Ideal clima cálido (coco, tonka, bergamota)." }
        "(?i)Scandal" { $tips = "Extremadamente dulce (miel/caramelo), denso y nocturno." }
        "(?i)L1212 Blanc" { $tips = "Fresco y deportivo. Huele a algodón blanco recién lavado." }
        "(?i)Lacoste Rouge" { $tips = "Energético y especiado. Manzana roja sobre maderas picantes." }
        "(?i)La vie est belle" { $tips = "Súper ventas dulzón (praliné/vainilla). Alta duración, empolvado." }
        "(?i)Afeef" { $tips = "Uso diario, maderas limpias y frutas en la salida." }
        "(?i)Asad Bourbon" { $tips = "Novedad dulce/especiada con potente vainilla bourbon." }
        "(?i)Asad" { $tips = "La mejor alternativa económica a Dior Sauvage Elixir. Potente y especiado." }
        "(?i)Badee Al Oud Amethyst" { $tips = "Clon de Atomic Rose (Initio). Rosa dominante, dulce y amaderado." }
        "(?i)Badee Al Oud Honor" { $tips = "Gourmand tropical único (acorde de piña caramelizada/creme brulee)." }
        "(?i)Badee Al Oud Oud for Glory" { $tips = "Clon exacto de Oud for Greatness. Opulento, azafrán y oud." }
        "(?i)Badee Al Oud Pink|Badee Al Oud Sublime" { $tips = "Clon de Eden Juicy Apple (Kayali). Afrutado, jugoso (manzana/ciruela)." }
        "(?i)Eclaire" { $tips = "Clon exacto de Bianco Latte. Extremadamente lactónico, caramelo y miel." }
        "(?i)His Confession" { $tips = "Elegante. Se posiciona como Armani Code Absolu / YSL Tuxedo." }
        "(?i)Khamrah Qahwa" { $tips = "Café tostado que equilibra el dulzor. Flanker muy exitoso." }
        "(?i)Khamrah Dukhan" { $tips = "Notas ahumadas y de incienso sobre la base dulce original." }
        "(?i)Khamrah" { $tips = "Inspirado en Angels' Share (Kilian). Dátiles dulces y especias cálidas." }
        "(?i)Yara Candy" { $tips = "Emula dulces frutales y confitería (estilo Oriana PDM)." }
        "(?i)Yara Tous" { $tips = "El Yara del verano. Coco y mango, vibra de protector solar de lujo." }
        "(?i)Yara" { $tips = "Súper viral. Huele a batido de fresa, dulce, atalcado y cremoso." }
        "(?i)Ajwad" { $tips = "Rendimiento extremo. Dulzor frutal metálico y rosa avainillada." }
        "(?i)Santal 33" { $tips = "Fragancia nicho de estatus. Seca, amaderada (sándalo), atalcada/cuero." }
        "(?i)Arabians Tonka" { $tips = "'Beast mode' (rendimiento nuclear). Azúcar quemada, rosa y oud intenso." }
        "(?i)Starry Nights" { $tips = "Frutas frescas en la salida y un secado denso de rosa y pachulí." }
        "(?i)Toy 2 Pearl" { $tips = "Mineral y salado. Simula arena mezclada con cítricos italianos." }
        "(?i)Toy 2" { $tips = "Aroma a limpio supremo. Gel de ducha de lujo o sábanas limpias." }
        "(?i)Toy Boy" { $tips = "Rosa picante (pimienta rosa) y maderas oscuras. Muy polarizante." }
        "(?i)Amber Rouge" { $tips = "Clon de altísima calidad de Baccarat Rouge 540 Extrait." }
        "(?i)Royal Amber" { $tips = "Clon premium de Erba Pura (Xerjoff). Explosión frutal y almizcle." }
        "(?i)Velvet Gold" { $tips = "Alternativa a Gentle Fluidity Gold (MFK). Dulce, atalcada y metálica." }
        "(?i)1 Million Lucky" { $tips = "Muy codiciado. Destaca el acorde único de avellana con miel." }
        "(?i)Fame" { $tips = "ADN tropical dominado por el mango jugoso, con jazmín e incienso." }
        "(?i)Invictus" { $tips = "Fundador de lo dulce/acuático. Sintético, deportivo, gran rendimiento." }
        "(?i)Phantom" { $tips = "Lavanda tradicional mezclada con tecnología y dulzor denso (terroso)." }
        "(?i)Delina" { $tips = "Versión más densa, empolvada y dulce de la icónica rosa de Delina." }
        "(?i)Kalan" { $tips = "Nicho polarizante. Terroso, seco, punzante (naranja sanguina y pimienta)." }
        "(?i)Rose Rush" { $tips = "Accesible, fresca, acuática y centrada en una rosa juvenil." }
        "(?i)360" { $tips = "Clásico de los 90. Perfil altamente pulcro, limpio y floral/acuático." }
        "(?i)Paradoxe" { $tips = "Sobredosis de flores blancas (neroli) con musgo para profundidad." }
        "(?i)Now Women" { $tips = "Alternativa económica a Burberry Her. Fresa sintética atalcada." }
        "(?i)Tobacco Vanille" { $tips = "Obra maestra. Picante del tabaco húmedo y dulzor resinoso (vainilla)." }
        "(?i)Tommy" { $tips = "Icono fresco y casual americano de los 90. Menta y manzana." }
        "(?i)Born In Roma" { $tips = "Secado moderno y versátil. Juega con acordes minerales y dulzor fresco." }
        "(?i)Spicebomb" { $tips = "'Bomba de especias'. Perfecto para el frío, picante y cálido a la vez." }
        "(?i)Libre" { $tips = "Toma la lavanda masculina y la feminiza con flor de azahar y vainilla." }
        "(?i)Marshmallow Blush" { $tips = "Extremadamente dulce, centrado exclusivamente en malvavisco/azúcar." }
        default { $tips = "Aroma único de la casa. Consultar perfil detallado." }
    }
    
    $worksheet.Cells.Item($row, 8).Value2 = $tips
}

$worksheet.Columns.AutoFit() | Out-Null

$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel llenado con Tips de Ventas."

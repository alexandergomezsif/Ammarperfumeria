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
    
    $familia = ""
    $notas = ""
    
    switch -Regex ($ref) {
        "(?i)Asad Bourbon" { $familia = "Amaderada/Especiada"; $notas = "Vainilla, Bourbon, Ámbar" }
        "(?i)Asad" { $familia = "Ámbar/Especiada"; $notas = "Pimienta negra, Vainilla, Maderas" }
        "(?i)Khamrah Qahwa" { $familia = "Gourmand"; $notas = "Café, Praliné, Vainilla, Canela" }
        "(?i)Khamrah Dukhan" { $familia = "Ambar/Especiada"; $notas = "Incienso, Dátiles, Vainilla oscura" }
        "(?i)Khamrah" { $familia = "Dulce/Gourmand"; $notas = "Dátiles, Praliné, Vainilla" }
        "(?i)Yara Candy" { $familia = "Floral/Frutal Dulce"; $notas = "Caramelo, Frutas dulces, Vainilla" }
        "(?i)Yara Tous" { $familia = "Floral/Frutal"; $notas = "Mango, Coco, Jazmín" }
        "(?i)Yara" { $familia = "Gourmand"; $notas = "Orquídea, Heliotropo, Frutas tropicales" }
        "(?i)Badee Al Oud Honor" { $familia = "Dulce/Especiada"; $notas = "Piña, Crème Brûlée, Especias" }
        "(?i)Badee Al Oud Amethyst" { $familia = "Ámbar/Floral"; $notas = "Rosa, Oud, Vainilla" }
        "(?i)Badee Al Oud Sublime" { $familia = "Frutal/Dulce"; $notas = "Manzana, Rosa, Ciruela" }
        "(?i)Badee Al Oud Oud for Glory" { $familia = "Amaderada/Oud"; $notas = "Oud, Azafrán, Nuez moscada" }
        "(?i)Badee Al Oud Pink" { $familia = "Floral/Dulce"; $notas = "Ruibarbo, Rosa, Vainilla" }
        "(?i)Eclaire" { $familia = "Gourmand"; $notas = "Leche, Azúcar, Caramelo, Miel" }
        "(?i)Oud For Greatness" { $familia = "Amaderada/Especiada"; $notas = "Oud, Azafrán, Nuez moscada" }
        "(?i)Starry Nights" { $familia = "Oriental/Floral"; $notas = "Rosa, Pachulí, Ámbar" }
        "(?i)Arabians Tonka" { $familia = "Ámbar/Amaderada"; $notas = "Haba Tonka, Azafrán, Oud, Azúcar" }
        "(?i)Club de Nuit Intense" { $familia = "Amaderada/Especiada"; $notas = "Limón, Piña, Abedul, Almizcle" }
        "(?i)Club de Nuit Untold" { $familia = "Ámbar/Amaderada"; $notas = "Azafrán, Jazmín, Ámbar gris" }
        "(?i)Lionheart" { $familia = "Fougère/Aromática"; $notas = "Bergamota, Lavanda, Especias" }
        "(?i)Acqua di Gio" { $familia = "Acuática/Aromática"; $notas = "Notas marinas, Bergamota, Romero" }
        "(?i)Si Passione" { $familia = "Floral/Frutal"; $notas = "Pera, Rosa, Vainilla" }
        "(?i)212 VIP Black" { $familia = "Aromática/Fougère"; $notas = "Absenta, Lavanda, Vainilla" }
        "(?i)Good Girl" { $familia = "Ámbar/Floral"; $notas = "Almendra, Nardo, Haba Tonka, Cacao" }
        "(?i)Coco Mademoiselle" { $familia = "Ambar/Floral"; $notas = "Naranja, Rosa, Pachulí" }
        "(?i)Aventus" { $familia = "Frutal/Amaderada"; $notas = "Piña, Abedul, Almizcle" }
        "(?i)Silver Mountain" { $familia = "Aromática"; $notas = "Té verde, Grosellas negras, Almizcle" }
        "(?i)Jadore" { $familia = "Floral"; $notas = "Jazmín, Rosa, Pera, Melón" }
        "(?i)Miss Dior" { $familia = "Floral"; $notas = "Rosa, Peonía, Almizcle blanco" }
        "(?i)Sauvage" { $familia = "Aromática/Fougère"; $notas = "Bergamota, Pimienta, Ambroxan" }
        "(?i)Light Blue" { $familia = "Cítrica/Fresca"; $notas = "Limón, Manzana, Bambú" }
        "(?i)Boss Bottled" { $familia = "Amaderada/Especiada"; $notas = "Manzana, Canela, Maderas" }
        "(?i)Il Femme" { $familia = "Ámbar/Floral"; $notas = "Rosa, Pachulí, Notas empolvadas" }
        "(?i)Il Roso" { $familia = "Floral/Dulce"; $notas = "Rosal, Especias dulces" }
        "(?i)Le Beau" { $familia = "Amaderada/Aromática"; $notas = "Coco, Haba Tonka, Bergamota" }
        "(?i)Scandal" { $familia = "Chipre/Floral"; $notas = "Miel, Pachulí, Naranja sanguina" }
        "(?i)L1212 Blanc" { $familia = "Amaderada/Floral"; $notas = "Toronja, Romero, Nardos" }
        "(?i)La vie est belle" { $familia = "Floral/Frutal/Gourmand"; $notas = "Iris, Praliné, Pachulí" }
        "(?i)Santal 33" { $familia = "Amaderada/Aromática"; $notas = "Sándalo, Cuero, Papiro, Cardamomo" }
        "(?i)Ombre Nomade" { $familia = "Ámbar/Amaderada"; $notas = "Oud, Incienso, Rosa, Frambuesa" }
        "(?i)Toy 2" { $familia = "Floral/Amaderada"; $notas = "Manzana verde, Peonía, Almizcle" }
        "(?i)Toy Boy" { $familia = "Amaderada/Especiada"; $notas = "Pimienta rosa, Rosa, Vetiver" }
        "(?i)Amber Rouge" { $familia = "Ámbar/Amaderada"; $notas = "Azafrán, Jazmín, Madera de cedro" }
        "(?i)Royal Amber" { $familia = "Ámbar/Vainilla"; $notas = "Ámbar, Vainilla, Notas dulces" }
        "(?i)Velvet Gold" { $familia = "Floral/Oriental"; $notas = "Rosa, Violeta, Almizcle, Ámbar" }
        "(?i)1 Million" { $familia = "Amaderada/Especiada"; $notas = "Canela, Cuero, Ámbar, Ciruela" }
        "(?i)Invictus" { $familia = "Acuática/Amaderada"; $notas = "Notas marinas, Toronja, Hoja de laurel" }
        "(?i)Phantom" { $familia = "Amaderada/Aromática"; $notas = "Lavanda, Limón, Vainilla" }
        "(?i)Delina" { $familia = "Floral"; $notas = "Rosa turca, Lichi, Peonía, Vainilla" }
        "(?i)Kalan" { $familia = "Oriental/Especiada"; $notas = "Naranja sanguina, Pimienta, Lavanda" }
        "(?i)360" { $familia = "Floral"; $notas = "Melón, Lirio, Lavanda" }
        "(?i)Paradoxe" { $familia = "Ámbar/Floral"; $notas = "Pera, Neroli, Vainilla Bourbon" }
        "(?i)Tobacco Vanille" { $familia = "Ámbar/Especiada"; $notas = "Hojas de tabaco, Vainilla, Cacao" }
        "(?i)Tommy" { $familia = "Cítrica/Aromática"; $notas = "Menta, Bergamota, Manzana" }
        "(?i)Born In Roma" { $familia = "Ambar/Floral"; $notas = "Grosella negra, Jazmín, Vainilla" }
        "(?i)Spicebomb" { $familia = "Amaderada/Especiada"; $notas = "Pimienta rosa, Canela, Tabaco" }
        "(?i)Libre" { $familia = "Ámbar/Fougère"; $notas = "Lavanda, Flor de azahar, Vainilla" }
        "(?i)Odyssey Mandarin" { $familia = "Cítrica/Gourmand"; $notas = "Mandarina, Caramelo, Haba Tonka" }
        "(?i)Odyssey Collection" { $familia = "Aromática"; $notas = "Iris, Ámbar, Especias" }
        "(?i)Fame" { $familia = "Floral/Amaderada"; $notas = "Mango, Jazmín, Incienso" }
        "(?i)Marshmallow Blush" { $familia = "Gourmand"; $notas = "Malvavisco, Vainilla, Azúcar" }
        "(?i)King" { $familia = "Frutal/Dulce"; $notas = "Naranja, Frutos Rojos, Vainilla" }
        "(?i)Ajwad" { $familia = "Floral/Amaderada"; $notas = "Rosa, Jazmín, Vainilla, Cedro" }
        default { $familia = "Mix/Variado"; $notas = "Consultar según versión específica" }
    }
    
    $worksheet.Cells.Item($row, 3).Value2 = $familia
    $worksheet.Cells.Item($row, 4).Value2 = $notas
}

$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel llenado con familias y notas."

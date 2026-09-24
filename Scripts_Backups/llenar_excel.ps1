$xlsx_path = "C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx"
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$workbook = $excel.Workbooks.Open($xlsx_path)
$worksheet = $workbook.Sheets.Item(1)
$lastRow = $worksheet.Cells.SpecialCells(11).Row

$worksheet.Cells.Item(1, 9).Value2 = "Genero"

$dictFamilias = @{
    "9AM DIVE" = "Aromática Acuática";
    "9PM REBEL" = "Aromática Frutal";
    "9PM" = "Ámbar Vainilla";
    "CLOUD 2.0 INTENSE" = "Ámbar Vainilla";
    "CLOUD PINK" = "Ámbar Vainilla";
    "CLOUD" = "Ámbar Vainilla";
    "MOD BLUSH" = "Ámbar Amaderada";
    "R.E.M." = "Ámbar Vainilla";
    "THANK U, NEXT" = "Floral Frutal Gourmand";
    "CLUB DE NUIT WOMAN" = "Floral Frutal";
    "CLUB DE NUIT UNTOLD" = "Ámbar Amaderada";
    "CLUB DE NUIT INTENSE" = "Chipre Frutal";
    "LIONHEART" = "Oriental Fougère";
    "ODYSSEY MANDARIN SKY ELIXIR" = "Oriental Amaderada";
    "ODYSSEY MANDARIN SKY" = "Ámbar Amaderada";
    "YUM YUM" = "Frutal Floral Gourmand";
    "ACQUA DI GIO PROFONDO" = "Aromática Acuática";
    "SÌ PASSIONE" = "Floral Frutal";
    "KING" = "Aromática";
    "NICHE FEMME" = "Floral";
    "SCARLET" = "Floral Amaderada / Frutal Floral";
    "212 VIP BLACK" = "Ámbar Fougère";
    "VERY GOOD GIRL" = "Floral Frutal";
    "GOOD GIRL" = "Ámbar Floral";
    "COCO MADEMOISELLE" = "Ámbar Amaderada";
    "AVENTUS" = "Chipre Frutal";
    "SILVER MOUNTAIN WATER" = "Aromática";
    "J'ADORE MARINA" = "Floral Acuática";
    "MISS DIOR" = "Floral Ámbar";
    "LIGHT BLUE" = "Floral Frutal / Amaderada según versión";
    "BOSS BOTTLED PARFUM" = "Amaderada Especiada";
    "BOSS BOTTLED" = "Amaderada Especiada";
    "IL FEMME" = "Aromática / Oriental Amaderada";
    "IL ROSO" = "Floral Frutal";
    "OUD FOR GREATNESS" = "Amaderada Ámbar";
    "LE BEAU" = "Amaderada Aromática";
    "SCANDAL" = "Chipre Floral Gourmand";
    "L12.12 BLANC" = "Amaderada Aromática";
    "LACOSTE ROUGE" = "Amaderada Especiada";
    "LA VIE EST BELLE" = "Floral Gourmand";
    "AFEEF" = "Floral Amaderada";
    "ART OF UNIVERSE" = "Aromática Cítrica";
    "ASAD BOURBON" = "Oriental Especiada";
    "ASAD" = "Ámbar";
    "BADEE AL OUD AMETHYST" = "Floral Ámbar";
    "BADEE AL OUD HONOR AND GLORY WHITE" = "Oriental Gourmand / Especiada";
    "BADEE AL OUD OUD FOR GLORY BLACK" = "Ámbar Amaderada";
    "BADEE AL OUD SUBLIME" = "Amaderada Aromática";
    "ECLAIRE" = "Floral Frutal Gourmand";
    "HIS CONFESSION" = "Oriental Amaderada";
    "KHAMRAH DUKHAN" = "Oriental";
    "KHAMRAH QAHWA" = "Aromática Especiada";
    "KHAMRAH" = "Oriental Especiada";
    "YARA CANDY" = "Frutal Gourmand";
    "YARA TOUS" = "Floral Tropical Gourmand";
    "YARA" = "Ámbar Vainilla";
    "SANTAL 33" = "Amaderada Aromática";
    "ARABIANS TONKA" = "Oriental Amaderada";
    "STARRY NIGHTS" = "Oriental Floral";
    "TOY 2 PEARL" = "Cítrica Floral Amaderada";
    "TOY 2" = "Floral Amaderada Almizclada";
    "TOY BOY 2" = "Amaderada Especiada";
    "AMBER ROUGE" = "Amaderada Especiada";
    "ROYAL AMBER" = "Oriental Amaderada";
    "VELVET GOLD" = "Oriental Suave / Gourmand Almizclada";
    "1 MILLION LUCKY" = "Amaderada";
    "FAME BLOOMING PINK" = "Almizclada Floral Amaderada";
    "INVICTUS" = "Amaderada Acuática";
    "PHANTOM" = "Aromática Fougère";
    "DELINA EXCLUSIF" = "Ámbar Floral";
    "DELINA" = "Floral";
    "KALAN" = "Oriental Especiada";
    "MARSHMALLOW BLUSH" = "Floral Gourmand";
    "ROSE RUSH" = "Floral Frutal";
    "360" = "Floral";
    "PARADOXE INTENSE" = "Floral Ámbar Amaderada";
    "NOW WOMEN" = "Floral Aldehídica";
    "TOBACCO VANILLE" = "Ámbar Especiada";
    "TOMMY" = "Cítrica Aromática";
    "UOMO BORN IN ROMA" = "Amaderada Aromática / Mineral";
    "SPICEBOMB" = "Amaderada Especiada";
    "LIBRE" = "Ámbar Fougère";
    "AJWAD" = "Amaderada Ámbar";
    "VICTORIA" = "Floral Frutal";
    "OMBRE NOMADE" = "Oriental Amaderada";
    "SAUVAGE PARFUM" = "Ámbar Fougère"
}

$dictNotas = @{
    "9AM DIVE" = "limón, menta, grosella negra, pimienta rosa, manzana, cedro, incienso, jengibre, sándalo, pachulí y jazmín.";
    "9PM REBEL" = "piña, manzana Granny Smith, mandarina, musgo de roble, cedro, vainilla, maderas secas, ámbar gris, caramelo y almizcle.";
    "9PM" = "manzana, canela, lavanda, bergamota, flor de azahar, lirio de los valles, vainilla, haba tonka, ámbar y pachulí.";
    "CLOUD 2.0 INTENSE" = "pera, lavanda, bergamota, praliné, crema de coco, orquídea de vainilla, maderas rubias, ambroxan, cashmeran y almizcle.";
    "CLOUD PINK" = "pitahaya, frutos rojos, piña, agua de coco, orquídea de vainilla, ambreta, praliné, almizcle, maderas ambaradas y musgo.";
    "CLOUD" = "lavanda, pera, bergamota, crema batida, coco, praliné, orquídea de vainilla, almizcle y maderas.";
    "MOD BLUSH" = "maracuyá, bergamota, frambuesa, pimienta rosa, magnolia, rosa, pera, ambrox, Dreamwood y almizcles.";
    "R.E.M." = "caramelo, sal, higo, membrillo, lavanda, flor de pera, malvavisco, haba tonka, sándalo y almizcle.";
    "THANK U, NEXT" = "pera, frambuesa, coco, rosa rosada, azúcar de macarrón y almizcle.";
    "CLUB DE NUIT WOMAN" = "naranja, bergamota, toronja, durazno, rosa, jazmín, geranio, lichi, pachulí, almizcle, vainilla y vetiver.";
    "CLUB DE NUIT UNTOLD" = "azafrán, jazmín, amberwood, ámbar gris, cedro.";
    "CLUB DE NUIT INTENSE" = "limón, piña, bergamota, grosella negra, manzana, abedul, jazmín, rosa, almizcle, ámbar gris, pachulí, vainilla.";
    "LIONHEART" = "menta, lavanda, vainilla, benjuí, miel, haba tonka y tabaco.";
    "ODYSSEY MANDARIN SKY ELIXIR" = "mandarina, naranja, lavanda, cardamomo, pimienta negra, caramelo, haba tonka, pachulí, incienso, vainilla y vetiver.";
    "ODYSSEY MANDARIN SKY" = "mandarina, naranja, azafrán, salvia, caramelo, haba tonka, caléndula, ambroxan, cedro y vetiver.";
    "YUM YUM" = "frutos rojos, frutas dulces, notas florales, vainilla, almizcle y acordes gourmand.";
    "ACQUA DI GIO PROFONDO" = "notas marinas, mandarina verde, bergamota, romero, lavanda, ciprés, lentisco, notas minerales, almizcle, pachulí y ámbar.";
    "SÌ PASSIONE" = "pera, grosella negra, pimienta rosa, toronja, piña, rosa, jazmín, heliotropo, vainilla, cedro y pachulí.";
    "KING" = "naranja, bergamota, limón, notas frutales, vainilla, almizcle blanco y ámbar.";
    "NICHE FEMME" = "frutas, rosa turca, incienso, sándalo y vainilla.";
    "SCARLET" = "pomelo, limón, bergamota, cardamomo, manzana verde, jazmín, nardo, rosa, clavel, musgo de roble, ámbar gris, madera de cashmere y vainilla.";
    "212 VIP BLACK" = "absenta, anís, hinojo, lavanda, cáscara de naranja, café, cuero, vainilla y almizcle.";
    "VERY GOOD GIRL" = "lichi, grosella roja, rosa, vainilla y vetiver.";
    "GOOD GIRL" = "almendra, café, bergamota, limón, nardo, jazmín sambac, raíz de lirio, rosa de Bulgaria, haba tonka, cacao, vainilla, praliné, sándalo, almizcle, ámbar y cedro.";
    "COCO MADEMOISELLE" = "naranja, mandarina, bergamota, limón, rosa, jazmín, ylang-ylang, pachulí, vetiver, vainilla, haba tonka, almizcle blanco y opopónaco.";
    "AVENTUS" = "grosella negra, bergamota, manzana, piña, rosa, abedul, jazmín, pachulí, almizcle, musgo de roble, ámbar gris y vainilla.";
    "SILVER MOUNTAIN WATER" = "bergamota, mandarina, té verde, grosella negra, almizcle, petitgrain, sándalo y gálbano.";
    "J'ADORE MARINA" = "limón, mandarina, naranja sanguina, romero, nenúfar, jazmín, magnolia, almizcle y maderas.";
    "MISS DIOR" = "iris, peonía, lirio de los valles, rosa, albaricoque, durazno, vainilla, almizcle, haba tonka, sándalo y benjuí.";
    "LIGHT BLUE" = "limón siciliano, manzana verde, cedro, campanilla, jazmín, bambú, rosa, almizcle y ámbar.";
    "BOSS BOTTLED PARFUM" = "incienso, mandarina, higo, iris, cedro, cuero.";
    "BOSS BOTTLED" = "manzana, ciruela, bergamota, limón, canela, geranio, clavo, caoba, vainilla, sándalo, cedro, vetiver y olivo.";
    "IL FEMME" = "vainilla, rosa, notas polvorientas, maracuyá, nuez moscada, canela, cedro, maderas, sándalo, tabaco, almizcle y ládano.";
    "IL ROSO" = "bergamota, mandarina, ruibarbo, rosa, jazmín, nuez moscada, almizcle blanco, cedro, ámbar, incienso, vainilla y vetiver.";
    "OUD FOR GREATNESS" = "oud, azafrán, nuez moscada, lavanda, pachulí y almizcle.";
    "LE BEAU" = "bergamota, madera de coco y haba tonka.";
    "SCANDAL" = "naranja sanguina, mandarina, miel, gardenia, flor de azahar, jazmín, durazno, cera de abeja, caramelo, pachulí y regaliz.";
    "L12.12 BLANC" = "toronja, romero, cardamomo, hojas de cedro, ylang-ylang, nardo, cedro, ante y vetiver.";
    "LACOSTE ROUGE" = "manzana roja, pino, cardamomo, pimienta, canela, pachulí, vetiver y madera.";
    "LA VIE EST BELLE" = "grosella negra, pera, iris, jazmín, flor de azahar, praliné, vainilla, pachulí y haba tonka.";
    "AFEEF" = "durazno, pimienta rosa, bergamota, nardo, flor de azahar, jazmín, praliné, ámbar, sándalo y pachulí.";
    "ART OF UNIVERSE" = "mandarina, bergamota, jengibre, menta, pera, flor de azahar, almizcle, ámbar y cedro.";
    "ASAD BOURBON" = "lavanda, mirabel, pimienta rosa, cacao, nuez moscada, davana, vainilla bourbon, ámbar y vetiver.";
    "ASAD" = "pimienta negra, piña, tabaco, pachulí, café, iris, vainilla, ámbar, maderas secas, benjuí y ládano.";
    "BADEE AL OUD AMETHYST" = "pimienta rosa, bergamota, rosa turca, rosa búlgara, jazmín, oud, ámbar y vainilla.";
    "BADEE AL OUD HONOR AND GLORY WHITE" = "piña brûlée, cúrcuma, canela, pimienta negra, benjuí, vainilla, cashmeran, sándalo y musgo.";
    "BADEE AL OUD OUD FOR GLORY BLACK" = "azafrán, nuez moscada, lavanda, oud, pachulí y almizcle.";
    "BADEE AL OUD SUBLIME" = "manzana, lichi, rosa, ciruela, jazmín, musgo, vainilla y pachulí.";
    "ECLAIRE" = "caramelo, leche, azúcar, miel, flores blancas, vainilla, praliné y almizcle.";
    "HIS CONFESSION" = "mandarina, canela, lavanda, iris, benjuí, ciprés, Mahonial, vainilla, haba tonka, ámbar, incienso, cedro y pachulí.";
    "KHAMRAH DUKHAN" = "especias, pimienta de Jamaica, mandarina, incienso, ládano, flor de azahar, pachulí, praliné, tabaco, ámbar, haba tonka y benjuí.";
    "KHAMRAH QAHWA" = "jengibre, canela, cardamomo, praliné, frutas confitadas, flores blancas, café arábica, haba tonka, almizcle, benjuí y vainilla.";
    "KHAMRAH" = "canela, nuez moscada, bergamota, dátiles, praliné, nardo, vainilla, haba tonka, ámbar, mirra, benjuí y Akigalawood.";
    "YARA CANDY" = "mandarina verde, grosella negra, caramelo de fresa, gardenia, sándalo, jarabe de vainilla, almizcle y ámbar.";
    "YARA TOUS" = "mango, coco, maracuyá, jazmín, flor de azahar, heliotropo, vainilla, almizcle y cachemira.";
    "YARA" = "mandarina, heliotropo, orquídea, frutas tropicales, acorde gourmand, vainilla, sándalo y almizcle.";
    "SANTAL 33" = "cardamomo, violeta, iris, papiro, sándalo, cedro, cuero, ámbar y almizcle.";
    "ARABIANS TONKA" = "bergamota, azafrán, rosa, oud, ámbar, musgo de roble, haba tonka, azúcar moreno, cuero y almizcle blanco.";
    "STARRY NIGHTS" = "bergamota, limón, rosa, jazmín, pachulí, ámbar, vainilla, almizcle blanco y notas empolvadas.";
    "TOY 2 PEARL" = "limón, orégano, jazmín, arena, fresia, vetiver, ciprés y almizcle.";
    "TOY 2" = "manzana, mandarina, magnolia, grosella blanca, peonía, jazmín, almizcle, sándalo y madera de ámbar.";
    "TOY BOY 2" = "nuez moscada, jengibre, Akigalawood, café, mirra y vetiver.";
    "AMBER ROUGE" = "azafrán, jazmín, amberwood, ámbar gris, cedro y resina de abeto.";
    "ROYAL AMBER" = "canela, naranja, bálsamo de Gurjun, orquídea negra, pachulí y sándalo.";
    "VELVET GOLD" = "bergamota, caramelo, violeta, pimienta rosa, notas empolvadas, pachulí, rosa, vainilla, almizcle y acorde animal.";
    "1 MILLION LUCKY" = "ciruela, notas ozónicas, toronja, bergamota, avellana, miel, cedro, cashmeran, flor de azahar, jazmín, amberwood, pachulí, vetiver y musgo de roble.";
    "FAME BLOOMING PINK" = "mango, bergamota, jazmín, incienso de olíbano, vainilla y sándalo.";
    "INVICTUS" = "toronja, notas marinas, mandarina, hoja de laurel, jazmín, violeta, ámbar gris, madera de guayaco, musgo de roble, pachulí y ámbar.";
    "PHANTOM" = "limón, lavanda, manzana, pachulí, tierra, vainilla, lavanda ahumada y vetiver.";
    "DELINA EXCLUSIF" = "bergamota, pera, lichi, rosa damascena, incienso, vetiver, vainilla, almizcle, musgo y haba tonka.";
    "KALAN" = "naranja sanguina, pimienta negra, especias, lavanda, ajenjo, flor de azahar, haba tonka, cedro y musgo.";
    "MARSHMALLOW BLUSH" = "bergamota, toronja, frutas, jengibre, pimienta rosa, coco, nardo, ámbar, malvavisco, vainilla, maderas suaves y almizcle.";
    "ROSE RUSH" = "pétalos de rosa, lichi, neroli, rosa de mayo, peonía, papaya, almizcle blanco, ámbar y cedro.";
    "360" = "melón, lirio, osmanto, mandarina, rosa, nenúfar, lirio de los valles, lavanda, salvia, almizcle, sándalo, vetiver, ámbar y vainilla.";
    "PARADOXE INTENSE" = "neroli, Ambrofix y acorde de musgo.";
    "NOW WOMEN" = "té blanco, maracuyá, bergamota, aldehídos, flor de tiaré, jazmín, almizcle, palo de rosa y ámbar.";
    "TOBACCO VANILLE" = "hoja de tabaco, especias, vainilla, cacao, haba tonka, frutos secos y maderas.";
    "TOMMY" = "bergamota, menta, lavanda, manzana, arándano, rosa, jazmín, cactus, algodón, ámbar y maderas.";
    "UOMO BORN IN ROMA" = "hojas de violeta, sal, jengibre, salvia, maderas y vetiver.";
    "SPICEBOMB" = "pimienta rosa, bergamota, toronja, canela, azafrán, cuero, tabaco y vetiver.";
    "LIBRE" = "lavanda, mandarina, bergamota, flor de azahar, jazmín, orquídea, vainilla, haba tonka, ámbar gris, almizcle y cedro.";
    "AJWAD" = "frutas, canela, bergamota, rosa, jazmín, ámbar, vainilla, almizcle, cedro y oud.";
    "VICTORIA" = "flores silvestres, vainilla, sándalo.";
    "OMBRE NOMADE" = "oud, rosa, incienso, frambuesa, azafrán, ámbar, benjuí, abedul y geranio.";
    "SAUVAGE PARFUM" = "bergamota, mandarina, sándalo, incienso de olíbano, haba tonka, vainilla."
}

$dictGenero = @{
    "9AM DIVE" = "Caballero";
    "9PM" = "Caballero";
    "9PM REBEL" = "Caballero";
    "CLUB DE NUIT INTENSE" = "Caballero";
    "LIONHEART" = "Caballero";
    "ODYSSEY MANDARIN SKY" = "Caballero";
    "ODYSSEY MANDARIN SKY ELIXIR" = "Caballero";
    "ACQUA DI GIO PROFONDO" = "Caballero";
    "KING" = "Caballero";
    "212 VIP BLACK" = "Caballero";
    "AVENTUS" = "Caballero";
    "SAUVAGE PARFUM" = "Caballero";
    "BOSS BOTTLED" = "Caballero";
    "BOSS BOTTLED PARFUM" = "Caballero";
    "LE BEAU" = "Caballero";
    "L12.12 BLANC" = "Caballero";
    "LACOSTE ROUGE" = "Caballero";
    "ASAD" = "Caballero";
    "ASAD BOURBON" = "Caballero";
    "HIS CONFESSION" = "Caballero";
    "TOY BOY 2" = "Caballero";
    "1 MILLION LUCKY" = "Caballero";
    "INVICTUS" = "Caballero";
    "PHANTOM" = "Caballero";
    "KALAN" = "Caballero";
    "TOMMY" = "Caballero";
    "UOMO BORN IN ROMA" = "Caballero";
    "SPICEBOMB" = "Caballero";
    "CLOUD" = "Dama";
    "CLOUD 2.0 INTENSE" = "Dama";
    "CLOUD PINK" = "Dama";
    "MOD BLUSH" = "Dama";
    "R.E.M." = "Dama";
    "THANK U, NEXT" = "Dama";
    "CLUB DE NUIT WOMAN" = "Dama";
    "SÌ PASSIONE" = "Dama";
    "NICHE FEMME" = "Dama";
    "SCARLET" = "Dama";
    "GOOD GIRL" = "Dama";
    "VERY GOOD GIRL" = "Dama";
    "COCO MADEMOISELLE" = "Dama";
    "J'ADORE MARINA" = "Dama";
    "MISS DIOR" = "Dama";
    "LIGHT BLUE" = "Dama";
    "IL FEMME" = "Dama";
    "IL ROSO" = "Dama";
    "SCANDAL" = "Dama";
    "LA VIE EST BELLE" = "Dama";
    "AFEEF" = "Dama";
    "YARA" = "Dama";
    "YARA CANDY" = "Dama";
    "YARA TOUS" = "Dama";
    "TOY 2" = "Dama";
    "VELVET GOLD" = "Dama";
    "FAME BLOOMING PINK" = "Dama";
    "DELINA EXCLUSIF" = "Dama";
    "MARSHMALLOW BLUSH" = "Dama";
    "ROSE RUSH" = "Dama";
    "360" = "Dama";
    "PARADOXE INTENSE" = "Dama";
    "NOW WOMEN" = "Dama";
    "LIBRE" = "Dama";
    "VICTORIA" = "Dama";
    "CLUB DE NUIT UNTOLD" = "Unisex";
    "SILVER MOUNTAIN WATER" = "Unisex";
    "OUD FOR GREATNESS" = "Unisex";
    "BADEE AL OUD AMETHYST" = "Unisex";
    "BADEE AL OUD HONOR AND GLORY WHITE" = "Unisex";
    "BADEE AL OUD OUD FOR GLORY BLACK" = "Unisex";
    "BADEE AL OUD SUBLIME" = "Unisex";
    "ECLAIRE" = "Unisex";
    "KHAMRAH" = "Unisex";
    "KHAMRAH DUKHAN" = "Unisex";
    "KHAMRAH QAHWA" = "Unisex";
    "AJWAD" = "Unisex";
    "SANTAL 33" = "Unisex";
    "OMBRE NOMADE" = "Unisex";
    "ARABIANS TONKA" = "Unisex";
    "STARRY NIGHTS" = "Unisex";
    "TOY 2 PEARL" = "Unisex";
    "AMBER ROUGE" = "Unisex";
    "ROYAL AMBER" = "Unisex";
    "TOBACCO VANILLE" = "Unisex";
    "YUM YUM" = "Unisex";
    "ART OF UNIVERSE" = "Unisex"
}

for ($row = 2; $row -le $lastRow; $row++) {
    $ref = $worksheet.Cells.Item($row, 2).Value2
    if (-not $ref) { continue }
    
    if ($ref -eq "R.E.M. 2") {
        $worksheet.Cells.Item($row, 2).Value2 = "R.E.M."
        $worksheet.Cells.Item($row, 3).Value2 = "REM_Alt.jpg"
        $ref = "R.E.M."
    } elseif ($ref -eq "Lionheart 2") {
        $worksheet.Cells.Item($row, 2).Value2 = "Lionheart"
        $worksheet.Cells.Item($row, 3).Value2 = "Lionheart_Alt.jpg"
        $ref = "Lionheart"
    } elseif ($ref -eq "Sauvage 2") {
        $worksheet.Cells.Item($row, 2).Value2 = "Sauvage Parfum"
        $worksheet.Cells.Item($row, 3).Value2 = "Sauvage_Parfum.jpg"
        $ref = "Sauvage Parfum"
    } elseif ($ref -eq "Boss Bottled 2") {
        $worksheet.Cells.Item($row, 2).Value2 = "Boss Bottled Parfum"
        $worksheet.Cells.Item($row, 3).Value2 = "Boss_Bottled_Parfum.jpg"
        $ref = "Boss Bottled Parfum"
    } elseif ($ref -eq "Oud For Greatness Blue") {
        $worksheet.Cells.Item($row, 2).Value2 = "Oud For Greatness"
        $worksheet.Cells.Item($row, 3).Value2 = "Oud_For_Greatness_Alt.jpg"
        $ref = "Oud For Greatness"
    } elseif ($ref -eq "Badee Al Oud Pink") {
        $worksheet.Cells.Item($row, 2).Value2 = "ELIMINAR_DUPLICADO"
        continue
    } elseif ($ref -eq "Ombre Nomade 3") {
        $worksheet.Cells.Item($row, 2).Value2 = "Ombre Nomade"
        $worksheet.Cells.Item($row, 3).Value2 = "Ombre_Nomade.jpg"
        $ref = "Ombre Nomade"
    } elseif ($ref -eq "Born In Roma Set") {
        $worksheet.Cells.Item($row, 2).Value2 = "Uomo Born In Roma Set"
        $worksheet.Cells.Item($row, 3).Value2 = "Uomo_Born_In_Roma_Set.jpg"
        $ref = "Uomo Born In Roma Set"
        $worksheet.Cells.Item($row, 9).Value2 = "Caballero"
    } elseif ($ref -eq "Odyssey Collection") {
        $worksheet.Cells.Item($row, 6).Value2 = "COLECCIÓN"
        $worksheet.Cells.Item($row, 7).Value2 = "Depende de la fragancia"
        $worksheet.Cells.Item($row, 9).Value2 = "Unisex"
        continue
    }
    
    $key = $ref.ToUpper().Trim()
    
    $fam = ""
    $not = ""
    $gen = ""
    
    # Matching
    foreach ($k in $dictFamilias.Keys) {
        if ($key -eq $k) {
            $fam = $dictFamilias[$k]
            $not = $dictNotas[$k]
            break
        }
    }
    # If no exact match, try regex
    if ($fam -eq "") {
        foreach ($k in $dictFamilias.Keys) {
            if ($key -match [regex]::Escape($k)) {
                $fam = $dictFamilias[$k]
                $not = $dictNotas[$k]
                break
            }
        }
    }
    
    foreach ($k in $dictGenero.Keys) {
        if ($key -eq $k) {
            $gen = $dictGenero[$k]
            break
        }
    }
    if ($gen -eq "") {
        foreach ($k in $dictGenero.Keys) {
            if ($key -match [regex]::Escape($k)) {
                $gen = $dictGenero[$k]
                break
            }
        }
    }
    
    if ($fam) { $worksheet.Cells.Item($row, 6).Value2 = $fam }
    if ($not) { $worksheet.Cells.Item($row, 7).Value2 = $not }
    if ($gen) { $worksheet.Cells.Item($row, 9).Value2 = $gen }
}

$worksheet.Columns.AutoFit() | Out-Null
$workbook.Save()
$workbook.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Host "Excel actualizado con Familias, Notas exactas de ChatGPT y Géneros."

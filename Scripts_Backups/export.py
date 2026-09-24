import json
import os
import random
import win32com.client

xlsx_path = r'C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.xlsx'
fotos_dir = r'C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos'
out_js = r'C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\js\perfumes.js'

img_dict = {}
for root, dirs, files in os.walk(fotos_dir):
    for f in files:
        if f.lower().endswith('.jpg'):
            # e.g. "fotos/Lattafa/Asad.jpg"
            rel_path = f"fotos/{os.path.basename(root)}/{f}"
            ref_name = f[:-4].replace('_', ' ').upper().strip()
            img_dict[ref_name] = rel_path

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
wb = excel.Workbooks.Open(xlsx_path)
ws = wb.Sheets(1)
last_row = ws.UsedRange.Rows.Count

perfumes = []
last_marca = "Desconocido"

for row in range(2, last_row + 1):
    marca = ws.Cells(row, 1).Value
    if marca:
        last_marca = marca
    else:
        marca = last_marca
        
    ref = ws.Cells(row, 2).Value
    if not ref:
        continue
        
    gen = ws.Cells(row, 3).Value or ""
    fam = ws.Cells(row, 4).Value or ""
    notas = ws.Cells(row, 5).Value or ""
    top = ws.Cells(row, 6).Value or ""
    tips = ws.Cells(row, 7).Value or ""
    
    # Handle floats like 360.0
    if isinstance(ref, float):
        if ref.is_integer():
            ref = str(int(ref))
        else:
            ref = str(ref)
    else:
        ref = str(ref)
        
    ref_key = ref.upper().strip()
    
    img_path = "https://via.placeholder.com/300x400?text=Falta+Foto"
    if ref_key in img_dict:
        img_path = img_dict[ref_key]
    else:
        for k in img_dict:
            if ref_key in k:
                img_path = img_dict[k]
                break
                
    # Random price between 160000 and 250000, rounded to nearest 1000
    precio = random.randint(160, 250) * 1000
    
    perfumes.append({
        "Marca": str(marca),
        "Referencia": ref,
        "Genero": str(gen),
        "Familia": str(fam),
        "Notas": str(notas),
        "Topologia": str(top),
        "Tips": str(tips),
        "Imagen": img_path,
        "Precio": precio
    })

wb.Close(False)
excel.Quit()

with open(out_js, 'w', encoding='utf-8') as f:
    f.write("const perfumesData = " + json.dumps(perfumes, ensure_ascii=False, indent=2) + ";")

print(f"Exportados {len(perfumes)} perfumes exitosamente.")

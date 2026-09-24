import pandas as pd
import sqlite3
import random
import os

excel_path = 'Base_Datos_Perfumes.xlsx'
db_path = 'ammar.db'
fotos_dir = 'fotos'

img_dict = {}
for root, dirs, files in os.walk(fotos_dir):
    for f in files:
        if f.lower().endswith('.jpg'):
            rel_path = f"fotos/{os.path.basename(root)}/{f}"
            ref_name = f[:-4].replace('_', ' ').upper().strip()
            img_dict[ref_name] = rel_path

df = pd.read_excel(excel_path)
df = df.fillna('')

perfumes = []
last_marca = "Desconocido"

for index, row in df.iterrows():
    marca = str(row.iloc[0]).strip()
    if marca:
        last_marca = marca
    else:
        marca = last_marca
        
    ref = str(row.iloc[1]).strip()
    if not ref:
        continue
        
    gen = str(row.iloc[2]).strip()
    fam = str(row.iloc[3]).strip()
    notas = str(row.iloc[4]).strip()
    top = str(row.iloc[5]).strip()
    tips = str(row.iloc[6]).strip()
    
    if ref.endswith('.0'):
        ref = ref[:-2]
        
    ref_key = ref.upper()
    img_path = "https://via.placeholder.com/300x400?text=Falta+Foto"
    if ref_key in img_dict:
        img_path = img_dict[ref_key]
    else:
        for k in img_dict:
            if ref_key in k:
                img_path = img_dict[k]
                break
                
    precio = random.randint(160, 250) * 1000
    
    perfumes.append((marca, ref, gen, fam, notas, top, tips, img_path, precio))

conn = sqlite3.connect(db_path)
cursor = conn.cursor()
cursor.execute('''
CREATE TABLE IF NOT EXISTS perfumes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    marca TEXT,
    referencia TEXT,
    genero TEXT,
    familia TEXT,
    notas TEXT,
    topologia TEXT,
    tips TEXT,
    imagen TEXT,
    precio INTEGER
)
''')
cursor.execute('DELETE FROM perfumes')
cursor.executemany('''
INSERT INTO perfumes (marca, referencia, genero, familia, notas, topologia, tips, imagen, precio)
VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
''', perfumes)

conn.commit()
conn.close()
print(f"Migrados {len(perfumes)} perfumes a SQLite.")

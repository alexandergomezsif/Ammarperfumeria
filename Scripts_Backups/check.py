# -*- coding: utf-8 -*-
import pandas as pd
import os, random, json, sys
sys.stdout.reconfigure(encoding='utf-8')

EXCEL='Base_Datos_Perfumes.xlsx';FOTOS='fotos';OUT='index.html'
WP='+57 312 858 6188';WP_NUM='573128586188'

img_map={}
for root,_,files in os.walk(FOTOS):
    for f in files:
        if f.lower().endswith('.jpg'):
            img_map[f[:-4].replace('_',' ').upper().strip()]=f'fotos/{os.path.basename(root)}/{f}'

def find_img(ref):
    k=ref.upper()
    if k in img_map:return img_map[k]
    for mk,mv in img_map.items():
        if k in mk:return mv
    return 'https://images.unsplash.com/photo-1541643600914-78b084683702?w=400&q=80'

df=pd.read_excel(EXCEL).fillna('')
perfumes=[];last_marca='Desconocido';random.seed(99)

for _,row in df.iterrows():
    marca=str(row.iloc[0]).strip() or last_marca;last_marca=marca
    ref=str(row.iloc[1]).strip()
    if not ref:continue
    if ref.endswith('.0'):ref=ref[:-2]
    perfumes.append({'marca':marca,'ref':ref,'genero':str(row.iloc[2]).strip(),
        'familia':str(row.iloc[3]).strip(),'notas':str(row.iloc[4]).strip(),
        'top':str(row.iloc[5]).strip(),'tips':str(row.iloc[6]).strip(),
        'img':find_img(ref),'precio':random.randint(160,250)*1000,
        'destacado':random.random()<0.12})

DATA_JSON=json.dumps(perfumes,ensure_ascii=False)

with open('index_template.html','r',encoding='utf-8-sig') as f:
    template=f.read()

print('Template size:', len(template))
print('Has placeholder:', '__DATA_JSON__' in template)
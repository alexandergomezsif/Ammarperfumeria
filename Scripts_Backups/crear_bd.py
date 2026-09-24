import os
import csv

fotos_dir = r"C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\fotos"
csv_path = r"C:\Users\AXL\.gemini\antigravity\scratch\Ammar Perfumería\Base_Datos_Perfumes.csv"

# Columns for the database
headers = ["Marca", "Referencia", "Familia_Olfativa", "Notas_Principales", "Precio_Venta", "Stock", "Archivo_Foto"]

with open(csv_path, mode='w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f, delimiter=';') # Use semicolon for Spanish Excel
    writer.writerow(headers)
    
    for brand in os.listdir(fotos_dir):
        brand_path = os.path.join(fotos_dir, brand)
        if os.path.isdir(brand_path):
            for file in os.listdir(brand_path):
                if file.lower().endswith('.jpg'):
                    # Clean up the reference name
                    ref = file.replace('.jpg', '').replace('.JPG', '').replace('_', ' ')
                    
                    # If it's an unrenamed IMG or UUID, leave reference empty or mark it
                    if ref.startswith('IMG_') or len(ref) > 20 and '-' in ref:
                        ref_display = "PENDIENTE DE IDENTIFICAR"
                    else:
                        ref_display = ref

                    writer.writerow([brand, ref_display, "", "", "", "", file])

print("CSV creado exitosamente en: " + csv_path)

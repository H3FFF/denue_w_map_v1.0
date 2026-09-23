#!/usr/bin/env python
# coding: utf-8

# In[3]:


import pandas as pd
import json
import urllib.parse
from datetime import datetime
import os

# ============================================
# CONFIGURACIÓN DE RUTAS
# ============================================
# Si lo ejecutas en GitHub Actions, el archivo debe estar en la raíz o ajusta la ruta
ARCHIVO_EXCEL = "DENUE 2026.xlsx" 
ARCHIVO_JSON = "datos.json"

def generar_json():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f"❌ Error: No se encontró {ARCHIVO_EXCEL}")
        return

    try:
        df = pd.read_excel(ARCHIVO_EXCEL)
    except Exception as e:
        print(f"❌ Error al leer Excel: {e}")
        return

    if 'id' not in df.columns:
        print("❌ ERROR: No se encontró la columna 'id'")
        return

    # Limpieza básica
    df['id'] = df['id'].astype(str).str.strip()
    df["latitud"] = pd.to_numeric(df["latitud"], errors="coerce")
    df["longitud"] = pd.to_numeric(df["longitud"], errors="coerce")
    df = df.dropna(subset=["latitud", "longitud"])

    columnas_texto = ["nom_estab", "raz_social", "nom_vial", "numero_ext", 
                      "cod_postal", "telefono", "status"]
    for col in columnas_texto:
        if col in df.columns:
            df[col] = df[col].fillna("No disponible").astype(str).str.strip()

    if "status" in df.columns:
        df["status"] = df["status"].str.upper()

    # --- LÓGICA DE WHATSAPP ---
    def generar_link_whatsapp(row):
        numero = str(row.get('telefono', ''))
        nombre_estab = str(row.get('nom_estab', 'Estimado cliente')).strip()

        if not numero or numero in ["No disponible", "nan"]:
            return ""

        digits = ''.join(filter(str.isdigit, numero))

        if len(digits) == 12 and digits.startswith('52'): pass 
        elif len(digits) == 11 and (digits.startswith('01') or digits.startswith('1')):
            digits = '52' + digits[1:] 
        elif len(digits) == 10:
            digits = '52' + digits 
        else:
            return "" 

        mensaje = (
            f"Hola {nombre_estab},\n\n"
            "Te escribo porque creo que este programa de Amazon Hub Delivery Partner puede ser una excelente oportunidad para ti.\n\n"
            "Se trata de un esquema oficial de Amazon México en el que puedes generar ingresos extras entregando paquetes muy cerca de tu casa.\n\n"
            "¿Te interesa conocer más detalles? Responde a este mensaje."
        )
        mensaje_codificado = urllib.parse.quote(mensaje, safe='')
        return f"https://wa.me/{digits}?text={mensaje_codificado}"

    if "telefono" in df.columns:
        df['whatsapp_link'] = df.apply(generar_link_whatsapp, axis=1)

    # Detectar municipio
    municipio_col = None
    for col_alt in ['municipio', 'nom_mun', 'mun']:
        if col_alt in df.columns:
            municipio_col = col_alt
            break

    if municipio_col:
        df['municipio'] = df[municipio_col].astype(str).str.strip()

    # ============================================
    # GENERAR JSON
    # ============================================
    registros = []
    for _, row in df.iterrows():
        registro = {
            "id": str(row['id']),
            "latitud": float(row['latitud']),
            "longitud": float(row['longitud']),
            "nom_estab": row.get('nom_estab', ''),
            "raz_social": row.get('raz_social', ''),
            "nom_vial": row.get('nom_vial', ''),
            "numero_ext": row.get('numero_ext', ''),
            "cod_postal": row.get('cod_postal', ''),
            "municipio": row.get('municipio', 'No disponible'),
            "telefono": row.get('telefono', ''),
            "status": row.get('status', 'SIN STATUS'),
            "whatsapp_link": row.get('whatsapp_link', ''),
            "fecha_alta": str(row.get('fecha_alta', ''))
        }
        registros.append(registro)

    with open(ARCHIVO_JSON, 'w', encoding='utf-8') as f:
        json.dump({
            "ultima_actualizacion": datetime.now().isoformat(),
            "total_registros": len(registros),
            "datos": registros
        }, f, ensure_ascii=False, indent=2)

    print(f"✅ JSON generado exitosamente: {ARCHIVO_JSON} ({len(registros)} registros)")

if __name__ == "__main__":
    generar_json()


# In[2]:





# In[ ]:





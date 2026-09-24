#!/usr/bin/env python
# coding: utf-8
import pandas as pd
import json
import urllib.parse
from datetime import datetime
import os
import hashlib

# ============================================
# CONFIGURACIÓN DE COLORES
# ============================================
# Paleta de colores distintivos para Delegaciones y CPs
PALETA_COLORES = [
    "#FF5733", "#33FF57", "#3357FF", "#F333FF", "#FF33A8", 
    "#33FFF5", "#F5FF33", "#FF8C33", "#8C33FF", "#33FF8C",
    "#E74C3C", "#2ECC71", "#3498DB", "#9B59B6", "#F1C40F",
    "#1ABC9C", "#E67E22", "#34495E", "#7F8C8D", "#C0392B"
]

def obtener_color(nombre):
    """Genera un color consistente basado en el nombre del municipio o CP"""
    if not nombre or nombre == "No disponible":
        return "#95a5a6" # Gris por defecto
    hash_val = int(hashlib.md5(str(nombre).encode()).hexdigest(), 16)
    index = hash_val % len(PALETA_COLORES)
    return PALETA_COLORES[index]

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

    # Detectar municipio
    municipio_col = None
    for col_alt in ['municipio', 'nom_mun', 'mun']:
        if col_alt in df.columns:
            municipio_col = col_alt
            break

    if municipio_col:
        df['municipio'] = df[municipio_col].astype(str).str.strip()

    # ============================================
    # GENERAR JSON CON COLORES ASIGNADOS
    # ============================================
    registros = []
    for _, row in df.iterrows():
        mun = row.get('municipio', 'No disponible')
        cp = row.get('cod_postal', 'No disponible')
        
        registro = {
            "id": str(row['id']),
            "latitud": float(row['latitud']),
            "longitud": float(row['longitud']),
            "nom_estab": row.get('nom_estab', ''),
            "raz_social": row.get('raz_social', ''),
            "nom_vial": row.get('nom_vial', ''),
            "numero_ext": row.get('numero_ext', ''),
            "cod_postal": cp,
            "municipio": mun,
            "telefono": row.get('telefono', ''),
            "status": row.get('status', 'SIN STATUS'),
            # Asignamos colores únicos para filtrado visual
            "color_municipio": obtener_color(mun),
            "color_cp": obtener_color(cp),
            "whatsapp_link": "", # Puedes reactivar la lógica de WhatsApp aquí si gustas
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

# In[ ]:





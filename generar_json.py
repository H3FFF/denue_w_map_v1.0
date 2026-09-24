#!/usr/bin/env python
# coding: utf-8

import pandas as pd
import json
import urllib.parse
import urllib.request
from datetime import datetime
import os


ARCHIVO_EXCEL = "DENUE 2026.xlsx"
ARCHIVO_JSON  = "datos.json"
ARCHIVO_CAMBIOS = "cambios_status.json" 

APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxwIRiMnSRsLpsdCUP_48_3eqVRdFF5ve0Jn32502JCkSnEf43ho0KWdcza_n3p1s-6xw/exec"

def cargar_status_desde_sheets():
    """Obtiene el diccionario {id: status} desde Google Sheets vía Apps Script."""
    try:
        print(" Consultando status actual desde Google Sheets...")
        with urllib.request.urlopen(APPS_SCRIPT_URL, timeout=30) as resp:
            data = json.loads(resp.read().decode('utf-8'))
        
        if data.get("ok") and isinstance(data.get("data"), dict):
            status_dict = data["data"]
            print(f" {len(status_dict)} status cargados desde Sheets")
            return status_dict
        else:
            print("  Respuesta de Sheets no válida:", data)
            return {}
    except Exception as e:
        print(f" Error al consultar Sheets: {e}")
        return {}

def cargar_cambios_guardados():
    """Aplica el JSON de cambios exportado desde el navegador (si existe)."""
    if not os.path.exists(ARCHIVO_CAMBIOS):
        return {}
    try:
        with open(ARCHIVO_CAMBIOS, 'r', encoding='utf-8') as f:
            cambios = json.load(f)
        print(f" {len(cambios)} cambios de status encontrados en {ARCHIVO_CAMBIOS}")
        return cambios
    except Exception as e:
        print(f"  No se pudo leer {ARCHIVO_CAMBIOS}: {e}")
        return {}

def generar_json():
    if not os.path.exists(ARCHIVO_EXCEL):
        print(f" Error: No se encontró {ARCHIVO_EXCEL}")
        return

    try:
        df = pd.read_excel(ARCHIVO_EXCEL)
    except Exception as e:
        print(f" Error al leer Excel: {e}")
        return

    if 'id' not in df.columns:
        print(" ERROR: No se encontró la columna 'id'")
        return

    # ─── Limpieza básica ─────────────────────────────────────────────────────
    df['id'] = df['id'].astype(str).str.strip()

    columnas_mostrar = ["nom_estab", "raz_social", "nom_vial", "numero_ext",
                        "cod_postal", "fecha_alta", "telefono", "status"]

    df["latitud"]  = pd.to_numeric(df["latitud"],  errors="coerce")
    df["longitud"] = pd.to_numeric(df["longitud"], errors="coerce")
    df = df.dropna(subset=["latitud", "longitud"])

    for col in columnas_mostrar:
        if col in df.columns:
            df[col] = df[col].fillna("No disponible").astype(str)

    df["cod_postal"] = df["cod_postal"].astype(str).str.strip()
    if "telefono" in df.columns:
        df["telefono"] = df["telefono"].astype(str).str.strip()
    if "status" in df.columns:
        df["status"] = df["status"].astype(str).str.strip().str.upper()

    # ─── LÓGICA DE WHATSAPP ──────────────────────────────────────────────────
    def generar_link_whatsapp(row):
        numero       = str(row.get('telefono', ''))
        nombre_estab = str(row.get('nom_estab', 'Estimado cliente')).strip()

        if not numero or numero in ["No disponible", "nan"]:
            return ""

        digits = ''.join(filter(str.isdigit, numero))

        if len(digits) == 12 and digits.startswith('52'):
            pass
        elif len(digits) == 11 and (digits.startswith('01') or digits.startswith('1')):
            digits = '52' + digits[1:]
        elif len(digits) == 10:
            digits = '52' + digits
        else:
            return ""

        mensaje = (
            f"Hola {nombre_estab},\n\n"
            "Te escribo porque creo que este programa de Amazon Hub Delivery Partner puede ser una excelente oportunidad para ti.\n\n"
            "Se trata de un esquema oficial de Amazon México en el que puedes generar ingresos extras entregando paquetes muy cerca de tu casa (solo en un radio de aproximadamente 1 km). No necesitas vehículo especial ni hacer recorridos largos: trabajas en tu propia zona.\n\n"
            "¿Por qué vale la pena considerarlo?\n\n"
            "✅ Ingresos adicionales de forma sencilla\n"
            "✅ Flexibilidad de horarios\n"
            "✅ Forma parte del ecosistema oficial de Amazon Logistics\n"
            "✅ Recibes los paquetes directamente en tu negocio cada mañana\n\n"
            "Si te interesa conocer más detalles, puedes revisar la información oficial aquí:\n\n"
            "👉 Amazon Hub Delivery Partner – México\n"
            "https://logistics.amazon.com.mx/hubdelivery/marketing\n\n"
            "También te dejo un video corto que explica cómo funciona:\n"
            "▶️ https://youtu.be/UrKzMnWxKio?is=Hv_lLVemmjfjCBnB\n\n"
            "Si quieres que te explique más o tienes alguna duda, solo respóndeme a este mensaje y con gusto te ayudo.\n\n"
            "¡Saludos y que tengas un excelente día!"
        )
        mensaje_codificado = urllib.parse.quote(mensaje, safe='')
        return f"https://wa.me/{digits}?text={mensaje_codificado}"

    if "telefono" in df.columns:
        df['whatsapp_link'] = df.apply(generar_link_whatsapp, axis=1)

    municipio_col = None
    for col_alt in ['municipio', 'nom_mun', 'mun', 'municipio_nombre']:
        if col_alt in df.columns:
            municipio_col = col_alt
            break
    if municipio_col:
        df['municipio'] = df[municipio_col].astype(str).str.strip()

    status_sheets = cargar_status_desde_sheets()
    if status_sheets:
        actualizados_sheets = 0
        for idx, row in df.iterrows():
            id_str = str(row['id'])
            if id_str in status_sheets:
                nuevo = str(status_sheets[id_str]).strip().upper()
                if nuevo and nuevo != "NAN":
                    df.at[idx, 'status'] = nuevo
                    actualizados_sheets += 1
        print(f" {actualizados_sheets} registros actualizados con status de Sheets")


    cambios = cargar_cambios_guardados()
    if cambios:
        actualizados = 0
        for idx, row in df.iterrows():
            id_str = str(row['id'])
            if id_str in cambios:
                df.at[idx, 'status'] = cambios[id_str].get('nuevo', cambios[id_str])
                actualizados += 1
        print(f" {actualizados} registros actualizados con cambios locales")

        try:
            os.rename(ARCHIVO_CAMBIOS,
                      ARCHIVO_CAMBIOS.replace('.json', '_aplicado.json'))
            print(f"  {ARCHIVO_CAMBIOS} renombrado a *_aplicado.json")
        except Exception as e:
            print(f"  No se pudo renombrar {ARCHIVO_CAMBIOS}: {e}")

    try:
        df.to_excel(ARCHIVO_EXCEL, index=False)
        print(f" Excel actualizado: {ARCHIVO_EXCEL}")
    except Exception as e:
        print(f"  No se pudo actualizar el Excel: {e}")

   
    for _, row in df.iterrows():
        registro = {
            "id":            str(row['id']),
            "latitud":       float(row['latitud']),
            "longitud":      float(row['longitud']),
            "nom_estab":     row.get('nom_estab',   ''),
            "raz_social":    row.get('raz_social',  ''),
            "nom_vial":      row.get('nom_vial',    ''),
            "numero_ext":    row.get('numero_ext',  ''),
            "cod_postal":    row.get('cod_postal',  ''),
            "municipio":     row.get('municipio',   'No disponible'),
            "telefono":      row.get('telefono',    ''),
            "status":        row.get('status',      'SIN STATUS'),
            "whatsapp_link": row.get('whatsapp_link',''),
            "fecha_alta":    str(row.get('fecha_alta', ''))
        }
        registros.append(registro)

    with open(ARCHIVO_JSON, 'w', encoding='utf-8') as f:
        json.dump({
            "ultima_actualizacion": datetime.now().isoformat(),
            "total_registros":      len(registros),
            "datos":                registros
        }, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    generar_json()

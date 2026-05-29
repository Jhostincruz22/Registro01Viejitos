import os
import shutil
from datetime import datetime
import tkinter as tk

import pandas as pd

from .config import CARPETA_BACKUPS, EXCEL_CENTRAL


def formatear_fecha_evento(event):
    """Aplica máscara de fecha DD/MM/AAAA al teclear."""
    entry = event.widget
    texto = entry.get().replace("/", "")
    texto_filtrado = "".join([c for c in texto if c.isdigit()])

    if len(texto_filtrado) >= 4:
        texto_formateado = f"{texto_filtrado[:2]}/{texto_filtrado[2:4]}/{texto_filtrado[4:8]}"
    elif len(texto_filtrado) >= 2:
        texto_formateado = f"{texto_filtrado[:2]}/{texto_filtrado[2:]}"
    else:
        texto_formateado = texto_filtrado

    posicion_cursor = entry.index(tk.INSERT)
    entry.delete(0, tk.END)
    entry.insert(0, texto_formateado)

    if len(texto_formateado) in [3, 6] and event.keysym != "Backspace":
        entry.icursor(posicion_cursor + 1)


def parsear_fecha(texto_fecha):
    """Convierte un texto con fecha al objeto datetime."""
    if not texto_fecha:
        return None
    for formato in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto_fecha.strip(), formato)
        except ValueError:
            continue
    return None


def validar_coherencia_fechas(fecha_nacimiento, fecha_inscripcion):
    """Comprueba que la inscripción no sea anterior al nacimiento."""
    nac = parsear_fecha(fecha_nacimiento)
    insc = parsear_fecha(fecha_inscripcion)

    if not nac or not insc:
        return True

    return insc >= nac


def crear_backup_seguro(ruta_archivo=EXCEL_CENTRAL, carpeta_backups=CARPETA_BACKUPS):
    """Genera una copia de seguridad antes de cualquier escritura."""
    if not os.path.exists(ruta_archivo):
        return None

    os.makedirs(carpeta_backups, exist_ok=True)
    nombre_base, extension = os.path.splitext(os.path.basename(ruta_archivo))
    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = os.path.join(carpeta_backups, f"{nombre_base}_{marca_tiempo}{extension}")
    shutil.copy2(ruta_archivo, destino)
    return destino


def leer_archivo_central(ruta_archivo=EXCEL_CENTRAL):
    """Carga todas las hojas del archivo maestro a un diccionario."""
    if not os.path.exists(ruta_archivo):
        return {}

    diccionario = {}
    with pd.ExcelFile(ruta_archivo) as xls:
        for hoja in xls.sheet_names:
            diccionario[hoja] = pd.read_excel(xls, sheet_name=hoja)
    return diccionario


def guardar_archivo_central(diccionario_hojas, ruta_archivo=EXCEL_CENTRAL):
    """Reescribe el archivo maestro preservando todas las pestañas."""
    with pd.ExcelWriter(ruta_archivo, engine="openpyxl") as writer:
        for nombre_hoja, df in diccionario_hojas.items():
            df.to_excel(writer, sheet_name=nombre_hoja, index=False)

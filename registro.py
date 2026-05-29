import os
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
import pandas as pd

# Nombre único del archivo centralizado
EXCEL_CENTRAL = "Años No registrados a partir de 1901.xlsx"


def obtener_nombre_hoja(tomo_val=None, año_val=None):
    """Genera el nombre estandarizado de la pestaña (hoja) dentro del Excel"""
    tomo = tomo_val if tomo_val is not None else entry_tomo.get().strip()
    año = año_val if año_val is not None else entry_año.get().strip()
    
    tomo_limpio = "".join([c for c in tomo if c.isalnum()])
    año_limpio = "".join([c for c in año if c.isalnum()])
    
    if not tomo_limpio or not año_limpio:
        return None
    # Ejemplo de nombre de pestaña: "T1_1901" (Excel limita a 31 caracteres)
    return f"T{tomo_limpio}_{año_limpio}"


def escanear_historial_año(event=None):
    """Busca dentro del archivo central las pestañas que pertenezcan al año digitado"""
    año_actual = entry_año.get().strip()
    
    if len(año_actual) < 4:
        return

    registro_encontrado = False
    ultimo_tomo = ""
    ultimo_folio = "1"
    ultima_partida = "1"
    total_actas = 0

    if os.path.exists(EXCEL_CENTRAL):
        try:
            # Leer todas las hojas del Excel
            excel_file = pd.ExcelFile(EXCEL_CENTRAL)
            hojas_existentes = excel_file.sheet_names
            
            max_partida = -1
            hoja_mas_reciente = None
            
            # Buscar las pestañas que terminen con el año digitado (ej: _1901)
            for hoja in hojas_existentes:
                if hoja.endswith(f"_{año_actual}"):
                    try:
                        df = pd.read_excel(excel_file, sheet_name=hoja)
                        if not df.empty and "Partida" in df.columns:
                            ultima_fila = df.iloc[-1]
                            partida_act = int(ultima_fila.get("Partida", 0))
                            
                            if partida_act > max_partida:
                                max_partida = partida_act
                                hoja_mas_reciente = hoja
                                ultimo_tomo = str(ultima_fila.get("Tomo", ""))
                                ultimo_folio = str(ultima_fila.get("Folio", "1"))
                                ultima_partida = str(partida_act + 1)
                                total_actas = len(df)
                                registro_encontrado = True
                    except:
                        continue
        except Exception as e:
            print(f"Error al escanear el archivo: {e}")

    if registro_encontrado:
        # --- SE ENCONTRÓ TRABAJO PREVIO EN ESTE AÑO ---
        var_edit_tfp.set(0)
        
        entry_tomo.config(state="normal")
        entry_tomo.delete(0, tk.END)
        entry_tomo.insert(0, ultimo_tomo)
        entry_tomo.config(state="disabled")
        
        entry_folio.config(state="normal")
        entry_folio.delete(0, tk.END)
        entry_folio.insert(0, ultimo_folio)
        entry_folio.config(state="disabled")
        
        entry_partida.config(state="normal")
        entry_partida.delete(0, tk.END)
        entry_partida.insert(0, ultima_partida)
        entry_partida.config(state="disabled")
        
        lbl_contador.config(
            text=f"Historial cargado: Tomo {ultimo_tomo} ({año_actual}). Registros en esta pestaña: {total_actas}", 
            foreground="#27ae60"
        )
    else:
        # --- AÑO NUEVO O SIN REGISTROS ---
        var_edit_tfp.set(1)
        
        entry_tomo.config(state="normal")
        entry_tomo.delete(0, tk.END)
        
        entry_folio.config(state="normal")
        entry_folio.delete(0, tk.END)
        entry_folio.insert(0, "1")
        
        entry_partida.config(state="normal")
        entry_partida.delete(0, tk.END)
        entry_partida.insert(0, "1")
        
        lbl_contador.config(
            text=f"Año {año_actual} nuevo o sin pestañas. Asigne el Tomo para comenzar.", 
            foreground="#e67e22"
        )

    recalcular_partidas_dinamicas()


def formatear_fecha_evento(event):
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
    for formato in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(texto_fecha, formato)
        except ValueError:
            continue
    return None


def validar_y_avanzar(event, entry_actual, siguiente_elemento, tipo_validacion):
    valor = entry_actual.get().strip()
    if tipo_validacion == "fecha":
        if len(valor) < 8 or parsear_fecha(valor) is None:
            root.bell()
            return "break"
    elif tipo_validacion == "texto_largo":
        if len(valor) < 4:
            root.bell()
            return "break"
    siguiente_elemento.focus()
    return "break"


def alternar_edicion_tfp():
    estado = "normal" if var_edit_tfp.get() == 1 else "disabled"
    entry_tomo.config(state=estado)
    entry_folio.config(state=estado)
    entry_partida.config(state=estado)
    if var_edit_tfp.get() == 0:
        escanear_historial_año()


def ajustar_pantalla_y_formularios():
    num_partidas = var_num_partidas.get()
    if num_partidas == 1:
        root.geometry("450x660")
        frame_col2.pack_forget()
        frame_col3.pack_forget()
    elif num_partidas == 2:
        root.geometry("820x660")
        frame_col2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        frame_col3.pack_forget()
    elif num_partidas == 3:
        root.geometry("1180x660")
        frame_col2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        frame_col3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
    recalcular_partidas_dinamicas()


def recalcular_partidas_dinamicas(*args):
    try:
        partida_inicial = int(entry_partida.get().strip())
    except ValueError:
        partida_inicial = 1
    lbl_partida_v1.config(text=f"PARTIDA N°: {partida_inicial}")
    lbl_partida_v2.config(text=f"PARTIDA N°: {partida_inicial + 1}")
    lbl_partida_v3.config(text=f"PARTIDA N°: {partida_inicial + 2}")


def guardar_folio_completo():
    rubro = combo_rubro.get()
    tomo = entry_tomo.get().strip()
    año = entry_año.get().strip()
    folio = entry_folio.get().strip()
    partida_base = entry_partida.get().strip()
    
    if not tomo or not año or not folio or not partida_base:
        messagebox.showwarning("Faltan datos", "Por favor complete todos los parámetros del Libro (Tomo, Año, Folio, Partida).")
        return

    try:
        p_inicial = int(partida_base)
        f_actual = int(folio)
    except ValueError:
        messagebox.showerror("Error numérico", "Folio y Partida deben ser números enteros válidos.")
        return

    nombre_hoja = obtener_nombre_hoja(tomo, año)
    
    # Estructura para leer o crear el ecosistema de hojas del archivo central
    diccionario_hojas = {}
    if os.path.exists(EXCEL_CENTRAL):
        try:
            with pd.ExcelFile(EXCEL_CENTRAL) as xls:
                for sheet in xls.sheet_names:
                    diccionario_hojas[sheet] = pd.read_excel(xls, sheet_name=sheet)
        except Exception as e:
            messagebox.showerror("Error de lectura", f"No se pudo leer el archivo maestro central: {e}")
            return

    # Obtener los datos previos exclusivamente de ESTA pestaña
    df_existente_hoja = diccionario_hojas.get(nombre_hoja, pd.DataFrame())

    num_formularios = var_num_partidas.get()
    registros_a_guardar = []
    
    formularios = [
        {"insc": entry_insc1, "nac": entry_nac1, "nom": entry_nombre1, "com": entry_comunidad1, "mad": entry_madre1, "pad_chk": var_padre1, "pad_ent": entry_padre1, "mar_chk": var_marginacion1, "mar_ent": entry_marginacion1, "offset": 0},
        {"insc": entry_insc2, "nac": entry_nac2, "nom": entry_nombre2, "com": entry_comunidad2, "mad": entry_madre2, "pad_chk": var_padre2, "pad_ent": entry_padre2, "mar_chk": var_marginacion2, "mar_ent": entry_marginacion2, "offset": 1},
        {"insc": entry_insc3, "nac": entry_nac3, "nom": entry_nombre3, "com": entry_comunidad3, "mad": entry_madre3, "pad_chk": var_padre3, "pad_ent": entry_padre3, "mar_chk": var_marginacion3, "mar_ent": entry_marginacion3, "offset": 2}
    ]

    for i in range(num_formularios):
        f = formularios[i]
        f_insc = f["insc"].get().strip()
        f_nac = f["nac"].get().strip()
        f_nom = f["nom"].get().strip()
        f_com = f["com"].get().strip()
        f_mad = f["mad"].get().strip()
        
        f_padre = f["pad_ent"].get().strip() if f["pad_chk"].get() == 1 else ""
        f_marginacion = f["mar_ent"].get().strip() if f["mar_chk"].get() == 1 else ""
        
        if len(f_nom) < 4 or len(f_com) < 4 or len(f_mad) < 4:
            messagebox.showwarning("Datos Insuficientes", f"Error en Partida {p_inicial + f['offset']}: Campos obligatorios deben tener al menos 4 caracteres.")
            f["nom"].focus()
            return

        dt_insc = parsear_fecha(f_insc)
        dt_nac = parsear_fecha(f_nac)
        
        if not dt_insc or not dt_nac:
            messagebox.showerror("Error de Formato", f"Error en Partida {p_inicial + f['offset']}: Formato de fecha incorrecto.")
            f["insc"].focus()
            return

        partida_actual_calcular = p_inicial + f["offset"]

        # Duplicados locales en la pestaña
        es_repetido = False
        if not df_existente_hoja.empty and "Partida" in df_existente_hoja.columns:
            coincidencias = df_existente_hoja[
                (df_existente_hoja["Folio"].astype(int) == f_actual) & 
                (df_existente_hoja["Partida"].astype(int) == partida_actual_calcular)
            ]
            if not coincidencias.empty:
                es_repetido = True

        if es_repetido:
            f_marginacion = f"[REPETIDO] {f_marginacion}".strip()

        registros_a_guardar.append({
            "Tomo": tomo,
            "Folio": f_actual,
            "Partida": partida_actual_calcular,
            "Año Libro": año,
            "Rubro": rubro,
            "Fecha Inscripción": f_insc,
            "Fecha Nacimiento": f_nac,
            "Nombre Completo": f_nom,
            "Comunidad/Barrio": f_com,
            "Nombre de la Madre": f_mad,
            "Nombre del Padre": f_padre if f_padre else "No Registrado",
            "Marginación / Notas": f_marginacion,
        })

    try:
        df_nuevos = pd.DataFrame(registros_a_guardar)
        df_hoja_final = pd.concat([df_existente_hoja, df_nuevos], ignore_index=True) if not df_existente_hoja.empty else df_nuevos
        
        # Insertar o actualizar la hoja en nuestro diccionario general
        diccionario_hojas[nombre_hoja] = df_hoja_final
        
        # Reescribir el archivo conservando absolutamente todas las pestañas intactas
        with pd.ExcelWriter(EXCEL_CENTRAL, engine="openpyxl") as writer:
            for name_sheet, df_sheet in diccionario_hojas.items():
                df_sheet.to_excel(writer, sheet_name=name_sheet, index=False)

        # Avanzar contadores
        nuevo_folio = f_actual + 1
        nueva_partida_inicial = p_inicial + num_formularios
        
        var_edit_tfp.set(0)
        
        entry_folio.config(state="normal")
        entry_folio.delete(0, tk.END)
        entry_folio.insert(0, str(nuevo_folio))
        entry_folio.config(state="disabled")
        
        entry_partida.config(state="normal")
        entry_partida.delete(0, tk.END)
        entry_partida.insert(0, str(nueva_partida_inicial))
        entry_partida.config(state="disabled")
        
        entry_tomo.config(state="disabled")
        
        limpiar_todos_los_formularios()
        
        lbl_contador.config(
            text=f"Guardado en pestaña '{nombre_hoja}'. Total registros en hoja: {len(df_hoja_final)}", 
            foreground="#27ae60"
        )
        
        messagebox.showinfo("Guardado", f"Datos guardados con éxito en la pestaña '{nombre_hoja}'.")

    except Exception as e:
        messagebox.showerror("Error de Archivo", f"Asegúrate de cerrar el Excel central antes de guardar.\nError: {e}")


def limpiar_todos_los_formularios():
    lista_entries = [
        (entry_insc1, entry_nac1, entry_nombre1, entry_comunidad1, entry_madre1, entry_padre1, var_padre1, entry_marginacion1, var_marginacion1),
        (entry_insc2, entry_nac2, entry_nombre2, entry_comunidad2, entry_madre2, entry_padre2, var_padre2, entry_marginacion2, var_marginacion2),
        (entry_insc3, entry_nac3, entry_nombre3, entry_comunidad3, entry_madre3, entry_padre3, var_padre3, entry_marginacion3, var_marginacion3)
    ]
    for ins, nac, nom, com, mad, pad_e, pad_v, mar_e, mar_v in lista_entries:
        ins.delete(0, tk.END)
        nac.delete(0, tk.END)
        nom.delete(0, tk.END)
        com.delete(0, tk.END)
        mad.delete(0, tk.END)
        pad_e.config(state="normal")
        pad_e.delete(0, tk.END)
        pad_e.config(state="disabled")
        pad_v.set(0)
        mar_e.config(state="normal")
        mar_e.delete(0, tk.END)
        mar_e.config(state="disabled")
        mar_v.set(0)
        
    recalcular_partidas_dinamicas()
    entry_insc1.focus()


def crear_bloque_captura(parent, titulo_partida_ref):
    lbl_p = ttk.Label(parent, text="PARTIDA N°: -", font=("Arial", 11, "bold"), foreground="#1b4f72")
    lbl_p.pack(pady=5)
    
    frame_inputs = ttk.Frame(parent)
    frame_inputs.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
    
    ent_insc = ttk.Entry(frame_inputs, font=("Arial", 10))
    ent_nac = ttk.Entry(frame_inputs, font=("Arial", 10))
    ent_nom = ttk.Entry(frame_inputs, font=("Arial", 10))
    ent_com = ttk.Entry(frame_inputs, font=("Arial", 10))
    ent_mad = ttk.Entry(frame_inputs, font=("Arial", 10))
    
    labels_text = ["F. Inscripción:", "F. Nacimiento:", "Nombre Persona:", "Comunidad/Origen:", "Nombre Madre:"]
    entries_obj = [ent_insc, ent_nac, ent_nom, ent_com, ent_mad]
    
    for idx, texto in enumerate(labels_text):
        ttk.Label(frame_inputs, text=texto, font=("Arial", 9, "bold")).grid(row=idx, column=0, sticky=tk.W, pady=5, padx=2)
        entries_obj[idx].grid(row=idx, column=1, pady=5, sticky="ew")
        
    frame_inputs.columnconfigure(1, weight=1)
    
    # Padre dinámico
    ttk.Label(frame_inputs, text="¿Registra Padre?:", font=("Arial", 9, "bold")).grid(row=5, column=0, sticky=tk.W, pady=5)
    f_p = ttk.Frame(frame_inputs)
    f_p.grid(row=5, column=1, sticky="ew", pady=5)
    ent_pad = ttk.Entry(f_p, font=("Arial", 10), state="disabled")
    v_pad = tk.IntVar()
    chk_pad = ttk.Checkbutton(f_p, variable=v_pad, command=lambda: ent_pad.config(state="normal") if v_pad.get()==1 else (ent_pad.delete(0,tk.END), ent_pad.config(state="disabled")))
    chk_pad.grid(row=0, column=0, padx=(0,2))
    ent_pad.grid(row=0, column=1, sticky="ew")
    f_p.columnconfigure(1, weight=1)
    
    # Marginacion dinámica
    ttk.Label(frame_inputs, text="¿Marginación?:", font=("Arial", 9, "bold")).grid(row=6, column=0, sticky=tk.W, pady=5)
    f_m = ttk.Frame(frame_inputs)
    f_m.grid(row=6, column=1, sticky="ew", pady=5)
    ent_mar = ttk.Entry(f_m, font=("Arial", 10), state="disabled")
    v_mar = tk.IntVar()
    chk_mar = ttk.Checkbutton(f_m, variable=v_mar, command=lambda: ent_mar.config(state="normal") if v_mar.get()==1 else (ent_mar.delete(0,tk.END), ent_mar.config(state="disabled")))
    chk_mar.grid(row=0, column=0, padx=(0,2))
    ent_mar.grid(row=0, column=1, sticky="ew")
    f_m.columnconfigure(1, weight=1)
    
    ent_insc.bind("<KeyRelease>", formatear_fecha_evento)
    ent_nac.bind("<KeyRelease>", formatear_fecha_evento)
    
    return lbl_p, ent_insc, ent_nac, ent_nom, ent_com, ent_mad, chk_pad, v_pad, ent_pad, chk_mar, v_mar, ent_mar


# --- Inicialización de la Interfaz ---
root = tk.Tk()
root.title("Indexador de Registro Civil (Pestañas por Libro)")
root.geometry("450x660")
root.resizable(True, False)

style = ttk.Style()
style.theme_use("clam")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

# PANEL SUPERIOR: CONFIGURACIÓN
frame_control_top = ttk.LabelFrame(main_frame, text=" Parámetros de Libro Inteligentes ", padding="8")
frame_control_top.pack(fill=tk.X, pady=(0, 8))

frame_control_top.columnconfigure(1, weight=1)
frame_control_top.columnconfigure(3, weight=1)

ttk.Label(frame_control_top, text="Rubro:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=2, pady=3)
combo_rubro = ttk.Combobox(frame_control_top, values=["Nacimiento", "Defunción", "Matrimonio"], state="readonly", width=12)
combo_rubro.set("Nacimiento")
combo_rubro.grid(row=0, column=1, sticky=tk.W, padx=2, pady=3)

ttk.Label(frame_control_top, text="Año Libro:", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=2, pady=3)
entry_año = ttk.Entry(frame_control_top, width=10)
entry_año.insert(0, "1901")
entry_año.grid(row=0, column=3, sticky=tk.W, padx=2, pady=3)

ttk.Label(frame_control_top, text="Tomo:", font=("Arial", 9, "bold")).grid(row=1, column=0, sticky=tk.W, padx=2, pady=3)
entry_tomo = ttk.Entry(frame_control_top, width=10, state="disabled")
entry_tomo.grid(row=1, column=1, sticky=tk.W, padx=2, pady=3)

ttk.Label(frame_control_top, text="Folio:", font=("Arial", 9, "bold")).grid(row=1, column=2, sticky=tk.W, padx=2, pady=3)
entry_folio = ttk.Entry(frame_control_top, width=10, state="disabled")
entry_folio.grid(row=1, column=3, sticky=tk.W, padx=2, pady=3)

ttk.Label(frame_control_top, text="Partida Inic:", font=("Arial", 9, "bold")).grid(row=1, column=4, sticky=tk.W, padx=2, pady=3)
entry_partida = ttk.Entry(frame_control_top, width=8, state="disabled")
entry_partida.grid(row=1, column=5, sticky=tk.W, padx=2, pady=3)

# Selectores de cantidad de partidas en el Folio
frame_selectores = ttk.Frame(frame_control_top)
frame_selectores.grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=5)

ttk.Label(frame_selectores, text="Partidas en este Folio:  ", font=("Arial", 9, "bold"), foreground="#c0392b").pack(side=tk.LEFT)
var_num_partidas = tk.IntVar(value=1)
ttk.Radiobutton(frame_selectores, text="1 Partida", variable=var_num_partidas, value=1, command=ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(frame_selectores, text="2 Partidas", variable=var_num_partidas, value=2, command=ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)
ttk.Radiobutton(frame_selectores, text="3 Partidas", variable=var_num_partidas, value=3, command=ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)

var_edit_tfp = tk.IntVar()
check_edit_tfp = ttk.Checkbutton(frame_control_top, text="Forzar edición manual de parámetros (T, A, F, P)", variable=var_edit_tfp, command=alternar_edicion_tfp)
check_edit_tfp.grid(row=3, column=0, columnspan=5, sticky=tk.W, pady=2)


# PANEL DE TRABAJO RECIPIENTE
frame_contenedor_columnas = ttk.Frame(main_frame)
frame_contenedor_columnas.pack(fill=tk.BOTH, expand=True, pady=5)

frame_col1 = ttk.LabelFrame(frame_contenedor_columnas, text=" Acta Sección A ", padding="5")
frame_col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

frame_col2 = ttk.LabelFrame(frame_contenedor_columnas, text=" Acta Sección B ", padding="5")
frame_col3 = ttk.LabelFrame(frame_contenedor_columnas, text=" Acta Sección C ", padding="5")

lbl_partida_v1, entry_insc1, entry_nac1, entry_nombre1, entry_comunidad1, entry_madre1, chk_padre1, var_padre1, entry_padre1, chk_marginacion1, var_marginacion1, entry_marginacion1 = crear_bloque_captura(frame_col1, 1)
lbl_partida_v2, entry_insc2, entry_nac2, entry_nombre2, entry_comunidad2, entry_madre2, chk_padre2, var_padre2, entry_padre2, chk_marginacion2, var_marginacion2, entry_marginacion2 = crear_bloque_captura(frame_col2, 2)
lbl_partida_v3, entry_insc3, entry_nac3, entry_nombre3, entry_comunidad3, entry_madre3, chk_padre3, var_padre3, entry_padre3, chk_marginacion3, var_marginacion3, entry_marginacion3 = crear_bloque_captura(frame_col3, 3)


# --- CAPTURA DE TECLAS ENLAZADAS (ENTER) ---
entry_insc1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_insc1, entry_nac1, "fecha"))
entry_nac1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nac1, entry_nombre1, "fecha"))
entry_nombre1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nombre1, entry_comunidad1, "texto_largo"))
entry_comunidad1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_comunidad1, entry_madre1, "texto_largo"))
entry_madre1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_madre1, chk_padre1, "texto_largo"))
chk_padre1.bind("<Return>", lambda e: entry_padre1.focus() if var_padre1.get() == 1 else chk_marginacion1.focus())
entry_padre1.bind("<Return>", lambda e: validar_y_avanzar(e, entry_padre1, chk_marginacion1, "texto_largo"))
chk_marginacion1.bind("<Return>", lambda e: entry_marginacion1.focus() if var_marginacion1.get() == 1 else (entry_insc2.focus() if var_num_partidas.get() >= 2 else guardar_folio_completo()))
entry_marginacion1.bind("<Return>", lambda e: entry_insc2.focus() if var_num_partidas.get() >= 2 else guardar_folio_completo())

entry_insc2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_insc2, entry_nac2, "fecha"))
entry_nac2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nac2, entry_nombre2, "fecha"))
entry_nombre2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nombre2, entry_comunidad2, "texto_largo"))
entry_comunidad2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_comunidad2, entry_madre2, "texto_largo"))
entry_madre2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_madre2, chk_padre2, "texto_largo"))
chk_padre2.bind("<Return>", lambda e: entry_padre2.focus() if var_padre2.get() == 1 else chk_marginacion2.focus())
entry_padre2.bind("<Return>", lambda e: validar_y_avanzar(e, entry_padre2, chk_marginacion2, "texto_largo"))
chk_marginacion2.bind("<Return>", lambda e: entry_marginacion2.focus() if var_marginacion2.get() == 1 else (entry_insc3.focus() if var_num_partidas.get() == 3 else guardar_folio_completo()))
entry_marginacion2.bind("<Return>", lambda e: entry_insc3.focus() if var_num_partidas.get() == 3 else guardar_folio_completo())

entry_insc3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_insc3, entry_nac3, "fecha"))
entry_nac3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nac3, entry_nombre3, "fecha"))
entry_nombre3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_nombre3, entry_comunidad3, "texto_largo"))
entry_comunidad3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_comunidad3, entry_madre3, "texto_largo"))
entry_madre3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_madre3, chk_padre3, "texto_largo"))
chk_padre3.bind("<Return>", lambda e: entry_padre3.focus() if var_padre3.get() == 1 else chk_marginacion3.focus())
entry_padre3.bind("<Return>", lambda e: validar_y_avanzar(e, entry_padre3, chk_marginacion3, "texto_largo"))
chk_marginacion3.bind("<Return>", lambda e: entry_marginacion3.focus() if var_marginacion3.get() == 1 else guardar_folio_completo())
entry_marginacion3.bind("<Return>", guardar_folio_completo)


# PANEL INFERIOR
lbl_contador = ttk.Label(main_frame, text="Escaneando archivo central...", font=("Arial", 10, "italic"))
lbl_contador.pack(pady=5)

btn_guardar_folio = ttk.Button(main_frame, text="Guardar Folio Completo", command=guardar_folio_completo)
btn_guardar_folio.pack(ipadx=30, ipady=5, pady=5)


# --- GESTORES DE EVENTOS EN TIEMPO REAL ---
entry_año.bind("<KeyRelease>", escanear_historial_año)
entry_partida.bind("<KeyRelease>", recalcular_partidas_dinamicas)

# Carga inicial segura
escanear_historial_año()
entry_insc1.focus()
root.mainloop()
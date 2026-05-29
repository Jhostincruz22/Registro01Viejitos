import os
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

from .buscador import VentanaBuscador
from .componentes import BloqueCapturaActa
from .config import EXCEL_CENTRAL
from .utils import crear_backup_seguro, guardar_archivo_central, leer_archivo_central, parsear_fecha, validar_coherencia_fechas


class VentanaPrincipal(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Indexador de Registro Civil")
        self.geometry("450x660")
        self.resizable(True, False)

        style = ttk.Style()
        style.theme_use("clam")

        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.var_num_partidas = tk.IntVar(value=1)
        self.var_edit_tfp = tk.IntVar(value=0)

        self._construir_controles()
        self._construir_bloques()
        self._configurar_bindings()

        self.escanear_historial_año()
        self.bloque_a.ent_insc.focus()

    def _construir_controles(self):
        frame_control_top = ttk.LabelFrame(self.main_frame, text=" Parámetros de Libro Inteligentes ", padding="8")
        frame_control_top.pack(fill=tk.X, pady=(0, 8))
        frame_control_top.columnconfigure(1, weight=1)
        frame_control_top.columnconfigure(3, weight=1)

        ttk.Label(frame_control_top, text="Rubro:", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky=tk.W, padx=2, pady=3)
        self.combo_rubro = ttk.Combobox(frame_control_top, values=["Nacimiento", "Defunción", "Matrimonio"], state="readonly", width=12)
        self.combo_rubro.set("Nacimiento")
        self.combo_rubro.grid(row=0, column=1, sticky=tk.W, padx=2, pady=3)

        ttk.Label(frame_control_top, text="Año Libro:", font=("Arial", 9, "bold")).grid(row=0, column=2, sticky=tk.W, padx=2, pady=3)
        self.entry_año = ttk.Entry(frame_control_top, width=10)
        self.entry_año.insert(0, "1901")
        self.entry_año.grid(row=0, column=3, sticky=tk.W, padx=2, pady=3)

        ttk.Label(frame_control_top, text="Tomo:", font=("Arial", 9, "bold")).grid(row=1, column=0, sticky=tk.W, padx=2, pady=3)
        self.entry_tomo = ttk.Entry(frame_control_top, width=10, state="disabled")
        self.entry_tomo.grid(row=1, column=1, sticky=tk.W, padx=2, pady=3)

        ttk.Label(frame_control_top, text="Folio:", font=("Arial", 9, "bold")).grid(row=1, column=2, sticky=tk.W, padx=2, pady=3)
        self.entry_folio = ttk.Entry(frame_control_top, width=10, state="disabled")
        self.entry_folio.grid(row=1, column=3, sticky=tk.W, padx=2, pady=3)

        ttk.Label(frame_control_top, text="Partida Inic:", font=("Arial", 9, "bold")).grid(row=1, column=4, sticky=tk.W, padx=2, pady=3)
        self.entry_partida = ttk.Entry(frame_control_top, width=8, state="disabled")
        self.entry_partida.grid(row=1, column=5, sticky=tk.W, padx=2, pady=3)

        frame_selectores = ttk.Frame(frame_control_top)
        frame_selectores.grid(row=2, column=0, columnspan=6, sticky=tk.W, pady=5)

        ttk.Label(frame_selectores, text="Partidas en este Folio:  ", font=("Arial", 9, "bold"), foreground="#c0392b").pack(side=tk.LEFT)
        ttk.Radiobutton(frame_selectores, text="1 Partida", variable=self.var_num_partidas, value=1, command=self.ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(frame_selectores, text="2 Partidas", variable=self.var_num_partidas, value=2, command=self.ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(frame_selectores, text="3 Partidas", variable=self.var_num_partidas, value=3, command=self.ajustar_pantalla_y_formularios).pack(side=tk.LEFT, padx=5)

        ttk.Checkbutton(
            frame_control_top,
            text="Forzar edición manual de parámetros (T, A, F, P)",
            variable=self.var_edit_tfp,
            command=self.alternar_edicion_tfp,
        ).grid(row=3, column=0, columnspan=6, sticky=tk.W, pady=2)

        self.frame_contenedor_columnas = ttk.Frame(self.main_frame)
        self.frame_contenedor_columnas.pack(fill=tk.BOTH, expand=True, pady=5)

        self.frame_col1 = ttk.Frame(self.frame_contenedor_columnas)
        self.frame_col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        self.frame_col2 = ttk.Frame(self.frame_contenedor_columnas)
        self.frame_col3 = ttk.Frame(self.frame_contenedor_columnas)

        self.lbl_contador = ttk.Label(self.main_frame, text="Preparando entorno...", font=("Arial", 10, "italic"))
        self.lbl_contador.pack(pady=5)

        frame_acciones = ttk.Frame(self.main_frame)
        frame_acciones.pack(pady=5)

        ttk.Button(frame_acciones, text="🔍 Abrir Buscador", command=self.abrir_buscador_vent).pack(side=tk.LEFT, padx=5, ipady=3)
        ttk.Button(frame_acciones, text="💾 Guardar Folio Completo", command=self.guardar_folio_completo).pack(side=tk.LEFT, padx=5, ipadx=15, ipady=3)

    def _construir_bloques(self):
        self.bloque_a = BloqueCapturaActa(self.frame_col1, "Acta Sección A", 0)
        self.bloque_a.pack(fill=tk.BOTH, expand=True)

        self.bloque_b = BloqueCapturaActa(self.frame_col2, "Acta Sección B", 1)
        self.bloque_c = BloqueCapturaActa(self.frame_col3, "Acta Sección C", 2)

        self.bloques = [self.bloque_a, self.bloque_b, self.bloque_c]
        self.ajustar_pantalla_y_formularios()

    def _configurar_bindings(self):
        self.entry_año.bind("<KeyRelease>", self.escanear_historial_año)
        self.entry_partida.bind("<KeyRelease>", self.recalcular_partidas_dinamicas)

        for bloque in self.bloques:
            bloque.ent_insc.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_insc, bloque_actual.ent_nac, "fecha"))
            bloque.ent_nac.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_nac, bloque_actual.ent_nom, "fecha"))
            bloque.ent_nom.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_nom, bloque_actual.ent_com, "texto_largo"))
            bloque.ent_com.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_com, bloque_actual.ent_mad, "texto_largo"))
            bloque.ent_mad.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_mad, bloque_actual.chk_pad, "texto_largo"))
            bloque.chk_pad.bind("<Return>", lambda event, bloque_actual=bloque: bloque_actual.ent_pad.focus() if bloque_actual.v_pad.get() == 1 else bloque_actual.chk_mar.focus())
            bloque.ent_pad.bind("<Return>", lambda event, bloque_actual=bloque: self.validar_y_avanzar(event, bloque_actual.ent_pad, bloque_actual.chk_mar, "texto_largo"))
            bloque.chk_mar.bind("<Return>", lambda event, bloque_actual=bloque: self._avanzar_desde_marginacion(bloque_actual))
            bloque.ent_mar.bind("<Return>", lambda event, bloque_actual=bloque: self._avanzar_desde_marginacion(bloque_actual))

    def _avanzar_desde_marginacion(self, bloque_actual):
        indice = self.bloques.index(bloque_actual)
        siguiente = self.bloques[indice + 1] if indice + 1 < len(self.bloques) else None
        if siguiente and self.var_num_partidas.get() > indice + 1:
            siguiente.ent_insc.focus()
            return "break"
        self.guardar_folio_completo()
        return "break"

    def obtener_nombre_hoja(self, tomo_val=None, año_val=None):
        tomo = tomo_val if tomo_val is not None else self.entry_tomo.get().strip()
        año = año_val if año_val is not None else self.entry_año.get().strip()

        tomo_limpio = "".join([c for c in tomo if c.isalnum()])
        año_limpio = "".join([c for c in año if c.isalnum()])

        if not tomo_limpio or not año_limpio:
            return None

        return f"T{tomo_limpio}_{año_limpio}"

    def validar_y_avanzar(self, event, entry_actual, siguiente_elemento, tipo_validacion):
        valor = entry_actual.get().strip()

        if valor == "[ILEGIBLE]":
            siguiente_elemento.focus()
            return "break"

        if tipo_validacion == "fecha":
            if len(valor) < 8 or parsear_fecha(valor) is None:
                self.bell()
                return "break"
        elif tipo_validacion == "texto_largo":
            if len(valor) < 4:
                self.bell()
                return "break"

        siguiente_elemento.focus()
        return "break"

    def alternar_edicion_tfp(self):
        estado = "normal" if self.var_edit_tfp.get() == 1 else "disabled"
        self.entry_tomo.config(state=estado)
        self.entry_folio.config(state=estado)
        self.entry_partida.config(state=estado)

        if self.var_edit_tfp.get() == 0:
            self.escanear_historial_año()

    def ajustar_pantalla_y_formularios(self):
        num_partidas = self.var_num_partidas.get()

        if num_partidas == 1:
            self.geometry("450x660")
            self.bloque_b.pack_forget()
            self.bloque_c.pack_forget()
        elif num_partidas == 2:
            self.geometry("820x660")
            self.frame_col2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            self.bloque_b.pack(fill=tk.BOTH, expand=True)
            self.bloque_c.pack_forget()
        elif num_partidas == 3:
            self.geometry("1180x660")
            self.frame_col2.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            self.frame_col3.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
            self.bloque_b.pack(fill=tk.BOTH, expand=True)
            self.bloque_c.pack(fill=tk.BOTH, expand=True)

        self.recalcular_partidas_dinamicas()

    def recalcular_partidas_dinamicas(self, *args):
        try:
            partida_inicial = int(self.entry_partida.get().strip())
        except ValueError:
            partida_inicial = 1

        for indice, bloque in enumerate(self.bloques):
            bloque.set_partida(partida_inicial + indice)

    def escanear_historial_año(self, event=None):
        año_actual = self.entry_año.get().strip()

        if len(año_actual) < 4:
            self.lbl_contador.config(text="Ingrese un año válido de 4 dígitos.", foreground="#7f8c8d")
            return

        registro_encontrado = False
        ultimo_tomo = ""
        ultimo_folio = "1"
        ultima_partida = "1"
        total_actas = 0

        if os.path.exists(EXCEL_CENTRAL):
            try:
                with pd.ExcelFile(EXCEL_CENTRAL) as xls:
                    max_partida = -1
                    for hoja in xls.sheet_names:
                        if not hoja.endswith(f"_{año_actual}"):
                            continue
                        try:
                            df = pd.read_excel(xls, sheet_name=hoja)
                            if df.empty or "Partida" not in df.columns:
                                continue
                            ultima_fila = df.iloc[-1]
                            partida_act = int(str(ultima_fila.get("Partida", 0)).split(".")[0])
                            if partida_act > max_partida:
                                max_partida = partida_act
                                ultimo_tomo = str(ultima_fila.get("Tomo", ""))
                                ultimo_folio = str(ultima_fila.get("Folio", "1"))
                                ultima_partida = str(partida_act + 1)
                                total_actas = len(df)
                                registro_encontrado = True
                        except Exception:
                            continue
            except Exception as error:
                print(f"Error al escanear el archivo: {error}")

        if registro_encontrado:
            self.var_edit_tfp.set(0)

            self.entry_tomo.config(state="normal")
            self.entry_tomo.delete(0, tk.END)
            self.entry_tomo.insert(0, ultimo_tomo)
            self.entry_tomo.config(state="disabled")

            self.entry_folio.config(state="normal")
            self.entry_folio.delete(0, tk.END)
            self.entry_folio.insert(0, ultimo_folio)
            self.entry_folio.config(state="disabled")

            self.entry_partida.config(state="normal")
            self.entry_partida.delete(0, tk.END)
            self.entry_partida.insert(0, ultima_partida)
            self.entry_partida.config(state="disabled")

            self.lbl_contador.config(
                text=f"Historial cargado: Tomo {ultimo_tomo} ({año_actual}). Registros en la pestaña: {total_actas}",
                foreground="#27ae60",
            )
        else:
            self.var_edit_tfp.set(1)

            self.entry_tomo.config(state="normal")
            self.entry_tomo.delete(0, tk.END)

            self.entry_folio.config(state="normal")
            self.entry_folio.delete(0, tk.END)
            self.entry_folio.insert(0, "1")

            self.entry_partida.config(state="normal")
            self.entry_partida.delete(0, tk.END)
            self.entry_partida.insert(0, "1")

            self.lbl_contador.config(
                text=f"Año {año_actual} nuevo o sin pestañas. Asigne el Tomo para comenzar.",
                foreground="#e67e22",
            )

        self.recalcular_partidas_dinamicas()

    def guardar_folio_completo(self):
        rubro = self.combo_rubro.get()
        tomo = self.entry_tomo.get().strip()
        año = self.entry_año.get().strip()
        folio = self.entry_folio.get().strip()
        partida_base = self.entry_partida.get().strip()

        if not tomo or not año or not folio or not partida_base:
            messagebox.showwarning("Faltan datos", "Complete Tomo, Año, Folio y Partida antes de guardar.")
            return

        try:
            p_inicial = int(partida_base)
            f_actual = int(folio)
        except ValueError:
            messagebox.showerror("Error numérico", "Folio y Partida deben ser enteros válidos.")
            return

        nombre_hoja = self.obtener_nombre_hoja(tomo, año)
        if not nombre_hoja:
            messagebox.showwarning("Hoja inválida", "El nombre de la pestaña no pudo generarse con los datos actuales.")
            return

        diccionario_hojas = leer_archivo_central()
        df_existente_hoja = diccionario_hojas.get(nombre_hoja, pd.DataFrame())

        num_formularios = self.var_num_partidas.get()
        registros_a_guardar = []

        for indice in range(num_formularios):
            bloque = self.bloques[indice]
            partida_actual = p_inicial + bloque.offset

            fecha_inscripcion = bloque.ent_insc.get().strip()
            fecha_nacimiento = bloque.ent_nac.get().strip()
            nombre_completo = bloque.ent_nom.get().strip()
            comunidad = bloque.ent_com.get().strip()
            madre = bloque.ent_mad.get().strip()

            padre = bloque.ent_pad.get().strip() if bloque.v_pad.get() == 1 else ""
            marginacion = bloque.ent_mar.get().strip() if bloque.v_mar.get() == 1 else ""

            es_ilegible_fecha_insc = fecha_inscripcion == "[ILEGIBLE]"
            es_ilegible_fecha_nac = fecha_nacimiento == "[ILEGIBLE]"
            es_ilegible_nombre = nombre_completo == "[ILEGIBLE]"
            es_ilegible_comunidad = comunidad == "[ILEGIBLE]"
            es_ilegible_madre = madre == "[ILEGIBLE]"

            if not es_ilegible_nombre and len(nombre_completo) < 4:
                messagebox.showwarning("Datos insuficientes", f"La partida {partida_actual} necesita un nombre válido.")
                bloque.ent_nom.focus()
                return

            if not es_ilegible_comunidad and len(comunidad) < 4:
                messagebox.showwarning("Datos insuficientes", f"La partida {partida_actual} necesita una comunidad/origen válida.")
                bloque.ent_com.focus()
                return

            if not es_ilegible_madre and len(madre) < 4:
                messagebox.showwarning("Datos insuficientes", f"La partida {partida_actual} necesita el nombre de la madre.")
                bloque.ent_mad.focus()
                return

            if not es_ilegible_fecha_insc:
                dt_insc = parsear_fecha(fecha_inscripcion)
                if not dt_insc:
                    messagebox.showerror("Error de formato", f"La fecha de inscripción en la partida {partida_actual} es inválida.")
                    bloque.ent_insc.focus()
                    return
            else:
                dt_insc = None

            if not es_ilegible_fecha_nac:
                dt_nac = parsear_fecha(fecha_nacimiento)
                if not dt_nac:
                    messagebox.showerror("Error de formato", f"La fecha de nacimiento en la partida {partida_actual} es inválida.")
                    bloque.ent_nac.focus()
                    return
            else:
                dt_nac = None

            if dt_insc and dt_nac and not validar_coherencia_fechas(fecha_nacimiento, fecha_inscripcion):
                continuar = messagebox.askyesno(
                    "Coherencia cronológica",
                    f"La partida {partida_actual} tiene inscripción anterior al nacimiento. ¿Desea guardar de todas formas?",
                )
                if not continuar:
                    return

            if not df_existente_hoja.empty and "Partida" in df_existente_hoja.columns:
                duplicados = df_existente_hoja[
                    (df_existente_hoja["Folio"].astype(str) == str(f_actual))
                    & (df_existente_hoja["Partida"].astype(str) == str(partida_actual))
                ]
                if not duplicados.empty:
                    marginacion = f"[REPETIDO] {marginacion}".strip()

            registros_a_guardar.append({
                "Tomo": tomo,
                "Folio": f_actual,
                "Partida": partida_actual,
                "Año Libro": año,
                "Rubro": rubro,
                "Fecha Inscripción": fecha_inscripcion,
                "Fecha Nacimiento": fecha_nacimiento,
                "Nombre Completo": nombre_completo,
                "Comunidad/Barrio": comunidad,
                "Nombre de la Madre": madre,
                "Nombre del Padre": padre if padre else "No Registrado",
                "Marginación / Notas": marginacion,
            })

        try:
            backup_path = crear_backup_seguro()
            df_nuevos = pd.DataFrame(registros_a_guardar)
            df_final = pd.concat([df_existente_hoja, df_nuevos], ignore_index=True) if not df_existente_hoja.empty else df_nuevos
            diccionario_hojas[nombre_hoja] = df_final
            guardar_archivo_central(diccionario_hojas)

            nuevo_folio = f_actual + 1
            nueva_partida_inicial = p_inicial + num_formularios

            self.entry_folio.config(state="normal")
            self.entry_folio.delete(0, tk.END)
            self.entry_folio.insert(0, str(nuevo_folio))
            self.entry_folio.config(state="disabled")

            self.entry_partida.config(state="normal")
            self.entry_partida.delete(0, tk.END)
            self.entry_partida.insert(0, str(nueva_partida_inicial))
            self.entry_partida.config(state="disabled")

            self.entry_tomo.config(state="disabled")

            for bloque in self.bloques[:num_formularios]:
                bloque.limpiar()

            self.lbl_contador.config(
                text=f"Guardado en '{nombre_hoja}'. Registros: {len(df_final)}. Backup: {os.path.basename(backup_path) if backup_path else 'sin respaldo'}",
                foreground="#27ae60",
            )
            messagebox.showinfo("Guardado", f"Datos guardados con éxito en la pestaña '{nombre_hoja}'.")
        except Exception as error:
            messagebox.showerror("Error de Archivo", f"No se pudo guardar el archivo central.\nError: {error}")

    def abrir_buscador_vent(self):
        VentanaBuscador(self)


if __name__ == "__main__":
    app = VentanaPrincipal()
    app.mainloop()

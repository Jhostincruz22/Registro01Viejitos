import os
import re
from difflib import SequenceMatcher

import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

from .config import EXCEL_CENTRAL


def _normalizar_texto(valor):
    if valor is None:
        return ""

    texto = str(valor).strip().lower()
    if texto in ["", "[ilegible]", "no registrado", "noregistrado"]:
        return ""

    reemplazos = (
        ("á", "a"),
        ("é", "e"),
        ("í", "i"),
        ("ó", "o"),
        ("ú", "u"),
        ("ü", "u"),
    )
    for a, b in reemplazos:
        texto = texto.replace(a, b)

    for nexo in [" de ", " del ", " la ", " las ", " los ", " y "]:
        texto = texto.replace(nexo, " ")

    texto = re.sub(r"[^a-z0-9ñ\s]", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


def _calcular_similitud(txt1, txt2):
    t1 = _normalizar_texto(txt1)
    t2 = _normalizar_texto(txt2)
    if not t1 or not t2:
        return 0.0
    return SequenceMatcher(None, t1, t2).ratio()


def _similitud_por_tokens(txt1, txt2):
    t1 = _normalizar_texto(txt1).split()
    t2 = _normalizar_texto(txt2).split()

    if not t1 or not t2:
        return 0.0

    set1 = set(t1)
    set2 = set(t2)
    interseccion = set1 & set2
    union = set1 | set2
    if not union:
        return 0.0

    return len(interseccion) / len(union)


def _similitud_heuristica(txt1, txt2):
    similitud_global = _calcular_similitud(txt1, txt2)
    similitud_tokens = _similitud_por_tokens(txt1, txt2)
    return round((similitud_global * 0.6) + (similitud_tokens * 0.4), 4)


class VentanaArbolAvanzado(tk.Toplevel):
    def __init__(self, parent, sujeto, todos_los_registros, callback_recargar=None):
        super().__init__(parent)
        self.title(f"Red Familiar Avanzada (1° a 3° Grado) - {sujeto.get('Nombre', 'Registro')}")
        self.geometry("860x620")
        self.minsize(760, 520)

        self.transient(parent)
        self.grab_set()

        self.sujeto = dict(sujeto)
        self.sujeto["Fecha Nacimiento"] = self.sujeto.get("Fecha Nacimiento") or self.sujeto.get("Fecha Nac.")
        self.sujeto["anio_nacimiento"] = self._extraer_anio(self.sujeto.get("Fecha Nacimiento"))

        if not self.sujeto.get("clave_registro"):
            self.sujeto["clave_registro"] = (
                f"{self.sujeto.get('Pestaña', '')}|"
                f"{self.sujeto.get('Tomo', '')}|"
                f"{self.sujeto.get('Folio', '')}|"
                f"{self.sujeto.get('Partida', '')}"
            )

        self.registros = self._normalizar_fila(todos_los_registros.copy())
        self.callback_recargar = callback_recargar
        self.vinculos_a_guardar = {}
        self.rechazados = {}
        self.descartados = {}
        self.estados_vinculos = {}

        self._construir_interfaz_arbol()
        self._analizar_red_completa()

    def _extraer_anio(self, valor):
        if pd.isna(valor):
            return None

        texto = str(valor).strip()
        if not texto:
            return None

        fecha = pd.to_datetime(texto, errors="coerce")
        if pd.notna(fecha):
            return int(fecha.year)

        try:
            return int(texto.split("/")[-1])
        except ValueError:
            try:
                return int(texto)
            except ValueError:
                return None

    def _normalizar_fila(self, df):
        df = df.copy()

        if "Nombre" not in df.columns and "Nombre Completo" in df.columns:
            df["Nombre"] = df["Nombre Completo"]
        if "Madre" not in df.columns and "Nombre de la Madre" in df.columns:
            df["Madre"] = df["Nombre de la Madre"]
        if "Padre" not in df.columns and "Nombre del Padre" in df.columns:
            df["Padre"] = df["Nombre del Padre"]
        if "Fecha Nacimiento" not in df.columns and "Fecha Nac." in df.columns:
            df["Fecha Nacimiento"] = df["Fecha Nac."]

        if "clave_registro" not in df.columns:
            df["clave_registro"] = (
                df["Pestaña"].astype(str)
                + "|"
                + df["Tomo"].astype(str)
                + "|"
                + df["Folio"].astype(str)
                + "|"
                + df["Partida"].astype(str)
            )

        df["clave_registro"] = df["clave_registro"].astype(str)
        df["Nombre"] = df["Nombre"].fillna("").astype(str)
        df["Madre"] = df["Madre"].fillna("").astype(str)
        df["Padre"] = df["Padre"].fillna("").astype(str)
        df["Fecha Nacimiento"] = df["Fecha Nacimiento"].fillna("").astype(str)
        df["anio_nacimiento"] = pd.to_numeric(pd.to_datetime(df["Fecha Nacimiento"], errors="coerce").dt.year, errors="coerce")

        if "nombre_normalizado" not in df.columns:
            df["nombre_normalizado"] = ""
        if "madre_normalizada" not in df.columns:
            df["madre_normalizada"] = ""
        if "padre_normalizado" not in df.columns:
            df["padre_normalizado"] = ""

        df["nombre_normalizado"] = df["nombre_normalizado"].replace("", pd.NA).fillna(df["Nombre"].map(_normalizar_texto))
        df["madre_normalizada"] = df["madre_normalizada"].replace("", pd.NA).fillna(df["Madre"].map(_normalizar_texto))
        df["padre_normalizado"] = df["padre_normalizado"].replace("", pd.NA).fillna(df["Padre"].map(_normalizar_texto))

        df["nombre_normalizado"] = df["nombre_normalizado"].astype(str)
        df["madre_normalizada"] = df["madre_normalizada"].astype(str)
        df["padre_normalizado"] = df["padre_normalizado"].astype(str)

        return df

    def _construir_interfaz_arbol(self):
        top_panel = ttk.LabelFrame(self, text=" Persona Central Evaluada ", padding=10)
        top_panel.pack(fill=tk.X, padx=10, pady=(10, 5))

        lugar_registro = self.sujeto.get("Pestaña", "")
        tomo = self.sujeto.get("Tomo", "")
        folio = self.sujeto.get("Folio", "")
        partida = self.sujeto.get("Partida", "")
        fecha_nacimiento = self.sujeto.get("Fecha Nacimiento") or self.sujeto.get("Fecha Nac.", "")

        lbl_info = (
            f"Sujeto: {self.sujeto.get('Nombre', '')}\n"
            f"Madre: {self.sujeto.get('Madre', '')} | Padre: {self.sujeto.get('Padre', '')}\n"
            f"Nacimiento: {fecha_nacimiento} | Ubicación: {lugar_registro} - Tomo {tomo} - Folio {folio} - Partida {partida}"
        )
        ttk.Label(top_panel, text=lbl_info, font=("Arial", 9, "bold"), foreground="#2c3e50", justify=tk.LEFT).pack(side=tk.LEFT)

        self.lbl_resumen = ttk.Label(
            top_panel,
            text="Pendientes: 0 | Confirmados: 0 | Rechazados: 0 — Doble clic para cambiar estado.",
            foreground="#2f4f4f",
            justify=tk.LEFT,
        )
        self.lbl_resumen.pack(side=tk.LEFT, padx=20)

        btn_guardar = ttk.Button(top_panel, text="💾 Guardar Éxitos en Excel", command=self._guardar_vinculos_en_excel)
        btn_guardar.pack(side=tk.RIGHT, padx=10, pady=5)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.tab_descendientes = ttk.Frame(self.notebook)
        self.tab_colaterales = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_descendientes, text="Descendencia (Hijos, Nietos y Bisnietos)")
        self.notebook.add(self.tab_colaterales, text="Colaterales (Hermanos y Sobrinos)")

        self.tabla_desc = self._crear_tabla_parentesco(self.tab_descendientes)
        self.tabla_colat = self._crear_tabla_parentesco(self.tab_colaterales)

    def _crear_tabla_parentesco(self, contenedor):
        frame = ttk.Frame(contenedor)
        frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        columnas = ("Nombre", "Parentesco", "Grado", "Puntaje", "Madre Inscrita", "Padre Inscrito", "Estado", "Observación")
        tabla = ttk.Treeview(frame, columns=columnas, show="headings", selectmode="browse")

        anchos = {
            "Nombre": 190,
            "Parentesco": 150,
            "Grado": 55,
            "Puntaje": 60,
            "Madre Inscrita": 150,
            "Padre Inscrito": 150,
            "Estado": 125,
            "Observación": 180,
        }

        for columna in columnas:
            tabla.heading(columna, text=columna)
            tabla.column(columna, width=anchos.get(columna, 110), anchor=tk.CENTER if columna in ["Grado", "Puntaje", "Estado", "Parentesco"] else tk.W)

        scroll_y = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tabla.yview)
        scroll_x = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tabla.xview)
        tabla.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        tabla.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        tabla.tag_configure("marcado_exito", background="#d1fae5", foreground="#065f46")
        tabla.tag_configure("rechazado", background="#fee2e2", foreground="#991b1b")
        tabla.tag_configure("descartado", background="#fef3c7", foreground="#92400e")
        tabla.bind("<Double-1>", lambda event, tabla=tabla: self._alternar_exito_vinculo(tabla))
        tabla.bind("<Button-3>", lambda event, tabla=tabla: self._mostrar_menu_descartar(event, tabla))

        return tabla

    def _ajustar_puntaje_por_homonimos(self, fila, puntaje):
        nombre = _normalizar_texto(fila.get("Nombre", ""))
        if not nombre:
            return puntaje

        frecuencia = self.nombre_frecuencia.get(nombre, 0)
        if frecuencia > 3:
            penalizacion = min(12.0, (frecuencia - 2) * 2.0)
            puntaje = max(0.0, puntaje - penalizacion)

        return puntaje

    def _alerta_edad(self, fila, parentesco):
        anio_persona = fila.get("anio_nacimiento")
        anio_sujeto = self.sujeto.get("anio_nacimiento")
        if pd.isna(anio_persona) or pd.isna(anio_sujeto):
            return ""

        alerta = []
        diferencia = anio_persona - anio_sujeto

        if parentesco.startswith("Hijo/a"):
            if diferencia <= 0:
                alerta.append("⚠️ REVISAR (nacimiento anterior al sujeto)")
            if diferencia < 12:
                alerta.append("⚠️ REVISAR (madre menor de 12 años al parto)")
            if diferencia > 60:
                alerta.append("⚠️ REVISAR (madre mayor de 60 años al parto)")

        if anio_persona <= anio_sujeto:
            alerta.append("⚠️ REVISAR (cronología incompatible)")

        return " | ".join(alerta)

    def _calcular_observacion(self, fila, parentesco, anchor=None):
        observaciones = []
        anio_persona = fila.get("anio_nacimiento")
        anio_sujeto = self.sujeto.get("anio_nacimiento")

        alerta_edad = self._alerta_edad(fila, parentesco)
        if alerta_edad:
            observaciones.append(alerta_edad)

        if pd.notna(anio_persona):
            if pd.notna(anio_sujeto):
                if parentesco == "Hijo/a" and anio_persona <= anio_sujeto:
                    observaciones.append("Nacimiento incompatible con la edad del sujeto")
                elif parentesco.startswith("Nieto/a") and anio_persona <= anio_sujeto:
                    observaciones.append("Cronología incompatible con la descendencia")
                elif parentesco.startswith("Bisnieto/a") and anio_persona <= anio_sujeto:
                    observaciones.append("Cronología incompatible con el bisnieto/a")
                elif parentesco.startswith("Sobrino/a") and anio_persona <= anio_sujeto:
                    observaciones.append("Cronología incompatible con el sobrino/a")

            if anchor and pd.notna(anchor.get("anio_nacimiento")) and anio_persona <= anchor["anio_nacimiento"]:
                observaciones.append("Nacimiento no posterior al pariente de referencia")

        if not observaciones:
            observaciones.append("Sin observaciones")

        return " | ".join(observaciones)

    def _agregar_fila(self, tabla, fila, parentesco, grado, puntaje, observacion="Sin observaciones"):
        clave = str(fila["clave_registro"])
        fila_id = f"{clave}::{grado}::{parentesco}"
        valores = (
            fila["Nombre"],
            parentesco,
            grado,
            f"{puntaje:.1f}",
            fila["Madre"],
            fila["Padre"],
            "Pendiente",
            observacion,
        )
        tabla.insert("", tk.END, iid=fila_id, values=valores)

        self.estados_vinculos[fila_id] = {
            "estado": "pendiente",
            "clave_registro": clave,
            "Parentesco": parentesco,
            "Grado": grado,
            "Puntaje": puntaje,
            "Nombre": fila["Nombre"],
            "Madre": fila["Madre"],
            "Padre": fila["Padre"],
            "Observacion": observacion,
        }

        return fila_id

    def _mostrar_menu_descartar(self, event, tabla):
        item = tabla.identify_row(event.y)
        if not item:
            return

        tabla.selection_set(item)
        tabla.focus(item)

        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label="❌ Marcar como descartado", command=lambda tabla=tabla: self._marcar_descartado(tabla))
        menu.post(event.x_root, event.y_root)

    def _marcar_descartado(self, tabla):
        item = tabla.focus()
        if not item or item not in self.estados_vinculos:
            return

        self.vinculos_a_guardar.pop(item, None)
        self.rechazados.pop(item, None)
        self.descartados[item] = dict(self.estados_vinculos[item])

        self.estados_vinculos[item]["estado"] = "descartado"
        self.estados_vinculos[item]["Observacion"] = "Descartado manualmente por el registrador"

        valores = list(tabla.item(item, "values"))
        valores[6] = "Descartado"
        valores[7] = "Descartado manualmente por el registrador"
        tabla.item(item, values=valores, tags=("descartado",))

        self._actualizar_resumen()

    def _analizar_red_completa(self):
        self.tabla_desc.delete(*self.tabla_desc.get_children())
        self.tabla_colat.delete(*self.tabla_colat.get_children())
        self.vinculos_a_guardar = {}
        self.rechazados = {}
        self.descartados = {}
        self.estados_vinculos = {}

        base = self.registros.copy()
        base = base[base["clave_registro"] != str(self.sujeto.get("clave_registro"))].copy()
        descartados_historicos = self._cargar_descartes_historicos()
        if descartados_historicos:
            base = base[~base["clave_registro"].isin(descartados_historicos)].copy()

        self.nombre_frecuencia = base["nombre_normalizado"].value_counts().to_dict()

        sujeto_nombre_norm = _normalizar_texto(self.sujeto.get("Nombre", ""))
        sujeto_madre_norm = _normalizar_texto(self.sujeto.get("Madre", ""))
        sujeto_padre_norm = _normalizar_texto(self.sujeto.get("Padre", ""))

        hijos_detectados = []
        nietos_detectados = []
        bisnietos_detectados = []
        hermanos_detectados = []
        sobrinos_detectados = []

        asignados = set()

        for _, fila in base.iterrows():
            clave = str(fila["clave_registro"])
            if clave in asignados:
                continue

            score_hijo = max(
                _calcular_similitud(fila["madre_normalizada"], sujeto_nombre_norm),
                _calcular_similitud(fila["padre_normalizado"], sujeto_nombre_norm),
            ) * 100
            score_hijo = self._ajustar_puntaje_por_homonimos(fila, score_hijo)

            if score_hijo >= 78:
                fila_dict = fila.to_dict()
                fila_dict["anio_nacimiento"] = fila["anio_nacimiento"]
                hijos_detectados.append({
                    "clave_registro": clave,
                    "Nombre": fila["Nombre"],
                    "Madre": fila["Madre"],
                    "Padre": fila["Padre"],
                    "anio_nacimiento": fila["anio_nacimiento"],
                    "score": round(score_hijo, 1),
                })
                observacion = self._calcular_observacion(fila_dict, "Hijo/a", self.sujeto)
                self._agregar_fila(self.tabla_desc, fila, "Hijo/a", "1°", round(score_hijo, 1), observacion)
                asignados.add(clave)
                continue

            score_hermano = max(
                _calcular_similitud(fila["madre_normalizada"], sujeto_madre_norm),
                _calcular_similitud(fila["padre_normalizado"], sujeto_padre_norm),
            ) * 100
            score_hermano = self._ajustar_puntaje_por_homonimos(fila, score_hermano)

            if score_hermano >= 82:
                fila_dict = fila.to_dict()
                fila_dict["anio_nacimiento"] = fila["anio_nacimiento"]
                hermanos_detectados.append({
                    "clave_registro": clave,
                    "Nombre": fila["Nombre"],
                    "Madre": fila["Madre"],
                    "Padre": fila["Padre"],
                    "anio_nacimiento": fila["anio_nacimiento"],
                    "score": round(score_hermano, 1),
                })
                observacion = self._calcular_observacion(fila_dict, "Hermano/a", self.sujeto)
                self._agregar_fila(self.tabla_colat, fila, "Hermano/a", "2°", round(score_hermano, 1), observacion)
                asignados.add(clave)

        for _, fila in base.iterrows():
            clave = str(fila["clave_registro"])
            if clave in asignados:
                continue

            mejor_nieto = None
            mejor_puntaje = 0

            for hijo in hijos_detectados:
                sim_madre_hijo = _calcular_similitud(fila["madre_normalizada"], _normalizar_texto(hijo["Nombre"]))
                sim_padre_hijo = _calcular_similitud(fila["padre_normalizado"], _normalizar_texto(hijo["Nombre"]))
                puntaje = max(sim_madre_hijo, sim_padre_hijo) * 100
                puntaje = self._ajustar_puntaje_por_homonimos(fila, puntaje)
                if puntaje >= 77 and puntaje > mejor_puntaje:
                    mejor_nieto = hijo
                    mejor_puntaje = puntaje

            if mejor_nieto is not None:
                fila_dict = fila.to_dict()
                fila_dict["anio_nacimiento"] = fila["anio_nacimiento"]
                anchor_dict = dict(mejor_nieto)
                anchor_dict["anio_nacimiento"] = mejor_nieto["anio_nacimiento"]
                nietos_detectados.append({
                    "clave_registro": clave,
                    "Nombre": fila["Nombre"],
                    "Madre": fila["Madre"],
                    "Padre": fila["Padre"],
                    "anio_nacimiento": fila["anio_nacimiento"],
                    "score": round(mejor_puntaje, 1),
                    "parentesco": f"Nieto/a (de {mejor_nieto['Nombre']})",
                })
                observacion = self._calcular_observacion(fila_dict, f"Nieto/a (de {mejor_nieto['Nombre']})", anchor_dict)
                self._agregar_fila(self.tabla_desc, fila, f"Nieto/a (de {mejor_nieto['Nombre']})", "2°", round(mejor_puntaje, 1), observacion)
                asignados.add(clave)
                continue

            mejor_sobrino = None
            mejor_puntaje_sobrino = 0

            for hermano in hermanos_detectados:
                sim_madre_hermano = _calcular_similitud(fila["madre_normalizada"], _normalizar_texto(hermano["Nombre"]))
                sim_padre_hermano = _calcular_similitud(fila["padre_normalizado"], _normalizar_texto(hermano["Nombre"]))
                puntaje = max(sim_madre_hermano, sim_padre_hermano) * 100
                puntaje = self._ajustar_puntaje_por_homonimos(fila, puntaje)
                if puntaje >= 77 and puntaje > mejor_puntaje_sobrino:
                    mejor_sobrino = hermano
                    mejor_puntaje_sobrino = puntaje

            if mejor_sobrino is not None:
                fila_dict = fila.to_dict()
                fila_dict["anio_nacimiento"] = fila["anio_nacimiento"]
                anchor_dict = dict(mejor_sobrino)
                anchor_dict["anio_nacimiento"] = mejor_sobrino["anio_nacimiento"]
                sobrinos_detectados.append({
                    "clave_registro": clave,
                    "Nombre": fila["Nombre"],
                    "Madre": fila["Madre"],
                    "Padre": fila["Padre"],
                    "anio_nacimiento": fila["anio_nacimiento"],
                    "score": round(mejor_puntaje_sobrino, 1),
                    "parentesco": f"Sobrino/a (de {mejor_sobrino['Nombre']})",
                })
                observacion = self._calcular_observacion(fila_dict, f"Sobrino/a (de {mejor_sobrino['Nombre']})", anchor_dict)
                self._agregar_fila(self.tabla_colat, fila, f"Sobrino/a (de {mejor_sobrino['Nombre']})", "3°", round(mejor_puntaje_sobrino, 1), observacion)
                asignados.add(clave)

        for _, fila in base.iterrows():
            clave = str(fila["clave_registro"])
            if clave in asignados:
                continue

            mejor_bisnieto = None
            mejor_puntaje = 0

            for nieto in nietos_detectados:
                sim_madre_nieto = _calcular_similitud(fila["madre_normalizada"], _normalizar_texto(nieto["Nombre"]))
                sim_padre_nieto = _calcular_similitud(fila["padre_normalizado"], _normalizar_texto(nieto["Nombre"]))
                puntaje = max(sim_madre_nieto, sim_padre_nieto) * 100
                puntaje = self._ajustar_puntaje_por_homonimos(fila, puntaje)
                if puntaje >= 75 and puntaje > mejor_puntaje:
                    mejor_bisnieto = nieto
                    mejor_puntaje = puntaje

            if mejor_bisnieto is not None:
                fila_dict = fila.to_dict()
                fila_dict["anio_nacimiento"] = fila["anio_nacimiento"]
                anchor_dict = dict(mejor_bisnieto)
                anchor_dict["anio_nacimiento"] = mejor_bisnieto["anio_nacimiento"]
                bisnietos_detectados.append({
                    "clave_registro": clave,
                    "Nombre": fila["Nombre"],
                    "Madre": fila["Madre"],
                    "Padre": fila["Padre"],
                    "anio_nacimiento": fila["anio_nacimiento"],
                    "score": round(mejor_puntaje, 1),
                    "parentesco": f"Bisnieto/a (de {mejor_bisnieto['Nombre']})",
                })
                observacion = self._calcular_observacion(fila_dict, f"Bisnieto/a (de {mejor_bisnieto['Nombre']})", anchor_dict)
                self._agregar_fila(self.tabla_desc, fila, f"Bisnieto/a (de {mejor_bisnieto['Nombre']})", "3°", round(mejor_puntaje, 1), observacion)
                asignados.add(clave)

        self._actualizar_resumen()

    def _cargar_descartes_historicos(self):
        descartados = set()
        if not os.path.exists(EXCEL_CENTRAL):
            return descartados

        try:
            with pd.ExcelFile(EXCEL_CENTRAL) as xls:
                for nombre_hoja in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=nombre_hoja)
                    if "Veredicto_Historico" not in df.columns:
                        continue

                    descartados_df = df[df["Veredicto_Historico"].fillna("").astype(str).str.contains("Descartado manualmente por el registrador", case=False, na=False)].copy()
                    if descartados_df.empty:
                        continue

                    for _, fila in descartados_df.iterrows():
                        clave = (
                            f"{nombre_hoja}|"
                            f"{str(fila.get('Tomo', '')).strip()}|"
                            f"{str(fila.get('Folio', '')).strip()}|"
                            f"{str(fila.get('Partida', '')).strip()}"
                        )
                        descartados.add(clave)
        except Exception:
            return descartados

        return descartados

    def _actualizar_resumen(self):
        total = len(self.estados_vinculos)
        confirmados = sum(1 for estado in self.estados_vinculos.values() if estado["estado"] == "confirmado")
        rechazados = sum(1 for estado in self.estados_vinculos.values() if estado["estado"] == "rechazado")
        descartados = sum(1 for estado in self.estados_vinculos.values() if estado["estado"] == "descartado")
        pendientes = total - confirmados - rechazados - descartados

        self.lbl_resumen.config(
            text=(
                f"Pendientes: {pendientes} | Confirmados: {confirmados} | Rechazados: {rechazados} | Descartados: {descartados} "
                "— Doble clic para cambiar estado, clic derecho para descartar."
            )
        )

    def _alternar_exito_vinculo(self, tabla):
        item = tabla.focus()
        if not item:
            return

        if item not in self.estados_vinculos:
            return

        estado_actual = self.estados_vinculos[item]["estado"]
        valores = list(tabla.item(item, "values"))

        if estado_actual == "pendiente":
            nuevo_estado = "confirmado"
            self.vinculos_a_guardar[item] = dict(self.estados_vinculos[item])
            self.rechazados.pop(item, None)
            self.descartados.pop(item, None)
            valores[6] = "Confirmado"
            valores[7] = self.estados_vinculos[item]["Observacion"]
            tabla.item(item, values=valores, tags=("marcado_exito",))
        elif estado_actual == "confirmado":
            nuevo_estado = "rechazado"
            self.vinculos_a_guardar.pop(item, None)
            self.rechazados[item] = dict(self.estados_vinculos[item])
            self.descartados.pop(item, None)
            valores[6] = "Rechazado"
            valores[7] = self.estados_vinculos[item]["Observacion"]
            tabla.item(item, values=valores, tags=("rechazado",))
        elif estado_actual == "rechazado":
            nuevo_estado = "pendiente"
            self.vinculos_a_guardar.pop(item, None)
            self.rechazados.pop(item, None)
            self.descartados.pop(item, None)
            valores[6] = "Pendiente"
            valores[7] = self.estados_vinculos[item]["Observacion"]
            tabla.item(item, values=valores, tags=())
        else:
            nuevo_estado = "pendiente"
            self.vinculos_a_guardar.pop(item, None)
            self.rechazados.pop(item, None)
            self.descartados.pop(item, None)
            valores[6] = "Pendiente"
            valores[7] = "Sin observaciones"
            tabla.item(item, values=valores, tags=())

        self.estados_vinculos[item]["estado"] = nuevo_estado
        self._actualizar_resumen()

    def _guardar_vinculos_en_excel(self):
        if not self.vinculos_a_guardar and not self.rechazados and not self.descartados:
            messagebox.showwarning("Sin Cambios", "No has marcado ningún parentesco confirmado, rechazado o descartado todavía.")
            return

        if not os.path.exists(EXCEL_CENTRAL):
            messagebox.showerror("Error", "No se encuentra el archivo Excel maestro para guardar.")
            return

        try:
            id_familia = f"FAM-{int(pd.Timestamp.now().timestamp())}"
            todas_las_hojas = {}

            xls_lector = pd.ExcelFile(EXCEL_CENTRAL)
            for nombre_hoja in xls_lector.sheet_names:
                df_hoja = pd.read_excel(xls_lector, sheet_name=nombre_hoja)
                df_hoja = df_hoja.copy()

                if "Nombre" not in df_hoja.columns and "Nombre Completo" in df_hoja.columns:
                    df_hoja["Nombre"] = df_hoja["Nombre Completo"]
                if "Madre" not in df_hoja.columns and "Nombre de la Madre" in df_hoja.columns:
                    df_hoja["Madre"] = df_hoja["Nombre de la Madre"]
                if "Padre" not in df_hoja.columns and "Nombre del Padre" in df_hoja.columns:
                    df_hoja["Padre"] = df_hoja["Nombre del Padre"]

                if "ID_Familia_Confirmado" not in df_hoja.columns:
                    df_hoja["ID_Familia_Confirmado"] = ""
                if "Veredicto_Historico" not in df_hoja.columns:
                    df_hoja["Veredicto_Historico"] = ""

                df_hoja["ID_Familia_Confirmado"] = df_hoja["ID_Familia_Confirmado"].fillna("").astype(str)
                df_hoja["Veredicto_Historico"] = df_hoja["Veredicto_Historico"].fillna("").astype(str)

                df_hoja["nombre_normalizado"] = df_hoja["nombre_normalizado"].fillna(df_hoja["Nombre"].fillna("").map(_normalizar_texto)) if "nombre_normalizado" in df_hoja.columns else df_hoja["Nombre"].fillna("").map(_normalizar_texto)
                df_hoja["madre_normalizada"] = df_hoja["madre_normalizada"].fillna(df_hoja["Madre"].fillna("").map(_normalizar_texto)) if "madre_normalizada" in df_hoja.columns else df_hoja["Madre"].fillna("").map(_normalizar_texto)
                df_hoja["padre_normalizado"] = df_hoja["padre_normalizado"].fillna(df_hoja["Padre"].fillna("").map(_normalizar_texto)) if "padre_normalizado" in df_hoja.columns else df_hoja["Padre"].fillna("").map(_normalizar_texto)

                todas_las_hojas[nombre_hoja] = df_hoja

            claves_confirmadas = {valor["clave_registro"]: valor for valor in self.vinculos_a_guardar.values()}
            claves_rechazadas = {valor["clave_registro"]: valor for valor in self.rechazados.values()}
            claves_descartadas = {valor["clave_registro"]: valor for valor in self.descartados.values()}
            cambios_efectuados = 0

            for hoja, df_hoja in todas_las_hojas.items():
                for idx, fila in df_hoja.iterrows():
                    clave_fila = (
                        f"{hoja}|"
                        f"{str(fila.get('Tomo', '')).strip()}|"
                        f"{str(fila.get('Folio', '')).strip()}|"
                        f"{str(fila.get('Partida', '')).strip()}"
                    )

                    if clave_fila == self.sujeto.get("clave_registro"):
                        df_hoja.at[idx, "ID_Familia_Confirmado"] = id_familia
                        df_hoja.at[idx, "Veredicto_Historico"] = "Sujeto Tronco Familiar"
                        cambios_efectuados += 1
                        continue

                    if clave_fila in claves_descartadas:
                        info_vinc = claves_descartadas[clave_fila]
                        df_hoja.at[idx, "Veredicto_Historico"] = (
                            f"Descartado manualmente por el registrador: {info_vinc['Parentesco']} de {info_vinc['Nombre']}"
                        )
                        cambios_efectuados += 1
                        continue

                    if clave_fila in claves_confirmadas:
                        info_vinc = claves_confirmadas[clave_fila]
                        df_hoja.at[idx, "ID_Familia_Confirmado"] = id_familia
                        df_hoja.at[idx, "Veredicto_Historico"] = (
                            f"Confirmado {info_vinc['Parentesco']} de {info_vinc['Nombre']}"
                        )
                        cambios_efectuados += 1
                        continue

                    if clave_fila in claves_rechazadas:
                        info_vinc = claves_rechazadas[clave_fila]
                        df_hoja.at[idx, "Veredicto_Historico"] = (
                            f"Rechazado por auditoría: {info_vinc['Parentesco']} de {info_vinc['Nombre']}"
                        )
                        cambios_efectuados += 1

            with pd.ExcelWriter(EXCEL_CENTRAL, engine="openpyxl") as writer:
                for hoja, df_hoja in todas_las_hojas.items():
                    df_hoja.to_excel(writer, sheet_name=hoja, index=False)

            messagebox.showinfo(
                "Éxito al Guardar",
                (
                    f"Se han guardado {cambios_efectuados} anotaciones en el Excel bajo el código: {id_familia}.\n"
                    f"Confirmados: {len(self.vinculos_a_guardar)} | Rechazados: {len(self.rechazados)} | Descartados: {len(self.descartados)}"
                ),
            )

            self.vinculos_a_guardar.clear()
            self.rechazados.clear()
            self.descartados.clear()
            if self.callback_recargar:
                self.callback_recargar()
            self.destroy()

        except Exception as error:
            messagebox.showerror("Error Crítico", f"No se pudo escribir en el Excel central. Verifica que no esté abierto por otra persona.\nDetalle: {error}")


class VentanaBuscador(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Central de inteligencia genealógica")
        self.geometry("1180x640")
        self.minsize(980, 520)

        self.transient(parent)
        self.grab_set()

        self.registros_df = pd.DataFrame()
        self.resultados_df = pd.DataFrame()

        self._construir_interfaz()
        self.cargar_registros()

    def _construir_interfaz(self):
        self.left_frame = ttk.Frame(self)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        frame_busqueda = ttk.LabelFrame(self.left_frame, text=" Búsqueda de registros y vínculos familiares ", padding=10)
        frame_busqueda.pack(fill=tk.X, padx=10, pady=(10, 5))

        ttk.Label(frame_busqueda, text="Buscar nombre, madre, padre o comunidad:", font=("Arial", 9, "bold")).pack(anchor=tk.W)

        controles_busqueda = ttk.Frame(frame_busqueda)
        controles_busqueda.pack(fill=tk.X, pady=(5, 0))

        self.entry_termino = ttk.Entry(controles_busqueda, width=45)
        self.entry_termino.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.entry_termino.bind("<Return>", lambda _event: self.ejecutar_busqueda())
        self.entry_termino.bind("<KeyRelease>", lambda _event: self.ejecutar_busqueda())

        ttk.Button(controles_busqueda, text="🔍 Buscar", command=self.ejecutar_busqueda).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(controles_busqueda, text="🧹 Limpiar", command=self.limpiar_busqueda).pack(side=tk.LEFT)

        self.lbl_estado = ttk.Label(self.left_frame, text="Esperando búsqueda...", foreground="#2c3e50")
        self.lbl_estado.pack(anchor=tk.W, padx=10, pady=(0, 5))

        frame_tabla = ttk.Frame(self.left_frame)
        frame_tabla.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        columnas = (
            "Pestaña",
            "Tomo",
            "Folio",
            "Partida",
            "Nombre",
            "Fecha Nac.",
            "Madre",
            "Padre",
            "Comunidad",
            "Fecha Inscripción",
        )
        self.tabla = ttk.Treeview(frame_tabla, columns=columnas, show="headings")

        anchos = {
            "Pestaña": 70,
            "Tomo": 55,
            "Folio": 55,
            "Partida": 60,
            "Nombre": 200,
            "Fecha Nac.": 90,
            "Madre": 180,
            "Padre": 180,
            "Comunidad": 140,
            "Fecha Inscripción": 95,
        }

        for columna in columnas:
            self.tabla.heading(columna, text=columna)
            self.tabla.column(columna, width=anchos.get(columna, 100), anchor=tk.W if columna in ["Nombre", "Madre", "Padre", "Comunidad"] else tk.CENTER)

        self.scroll_y = ttk.Scrollbar(frame_tabla, orient=tk.VERTICAL, command=self.tabla.yview)
        self.scroll_x = ttk.Scrollbar(frame_tabla, orient=tk.HORIZONTAL, command=self.tabla.xview)
        self.tabla.configure(yscrollcommand=self.scroll_y.set, xscrollcommand=self.scroll_x.set)

        self.tabla.grid(row=0, column=0, sticky="nsew")
        self.scroll_y.grid(row=0, column=1, sticky="ns")
        self.scroll_x.grid(row=1, column=0, sticky="ew")

        frame_tabla.rowconfigure(0, weight=1)
        frame_tabla.columnconfigure(0, weight=1)

        self.tabla.bind("<Double-1>", self._seleccionar_fila_desde_tabla)

        self.menu_contextual = tk.Menu(self, tearoff=0)
        self.menu_contextual.add_command(label="👀 Ver árbol genealógico avanzado", command=self._abrir_arbol_avanzado)
        self.tabla.bind("<Button-3>", self._mostrar_menu_contextual)
        self.tabla.bind("<Button-2>", self._mostrar_menu_contextual)

        self.right_frame = ttk.LabelFrame(self, text=" Mapa mental familiar ", padding=10)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=False, padx=(0, 10), pady=(10, 10))
        self.right_frame.configure(width=360)

        self.lbl_mapa = ttk.Label(self.right_frame, text="Seleccione una madre o busque un nombre para ver la red.")
        self.lbl_mapa.pack(anchor=tk.W, pady=(0, 5))

        self.canvas = tk.Canvas(self.right_frame, bg="#f8fafc", highlightthickness=0)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self._dibujar_mapa_vacio()

    def cargar_registros(self):
        self.registros_df = pd.DataFrame()

        if not os.path.exists(EXCEL_CENTRAL):
            self.lbl_estado.config(text="Archivo maestro no encontrado.")
            self._dibujar_mapa_vacio()
            return

        try:
            with pd.ExcelFile(EXCEL_CENTRAL) as xls:
                registros = []
                for nombre_hoja in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=nombre_hoja)
                    if df.empty:
                        continue

                    df = df.copy()
                    df["Pestaña"] = nombre_hoja

                    for columna in [
                        "Nombre Completo",
                        "Nombre de la Madre",
                        "Nombre del Padre",
                        "Comunidad/Barrio",
                        "Fecha Nacimiento",
                        "Fecha Inscripción",
                        "Tomo",
                        "Folio",
                        "Partida",
                    ]:
                        if columna not in df.columns:
                            df[columna] = ""

                    df["Nombre"] = df["Nombre Completo"].fillna("")
                    df["Madre"] = df["Nombre de la Madre"].fillna("")
                    df["Padre"] = df["Nombre del Padre"].fillna("")
                    df["Comunidad"] = df["Comunidad/Barrio"].fillna("")
                    df["Fecha Nacimiento"] = df["Fecha Nacimiento"].fillna("")
                    df["Fecha Inscripción"] = df["Fecha Inscripción"].fillna("")
                    df["anio_nacimiento"] = pd.to_numeric(pd.to_datetime(df["Fecha Nacimiento"], errors="coerce").dt.year, errors="coerce")

                    if "nombre_normalizado" not in df.columns:
                        df["nombre_normalizado"] = ""
                    if "madre_normalizada" not in df.columns:
                        df["madre_normalizada"] = ""
                    if "padre_normalizado" not in df.columns:
                        df["padre_normalizado"] = ""

                    df["nombre_normalizado"] = df["nombre_normalizado"].replace("", pd.NA).fillna(df["Nombre"].map(_normalizar_texto))
                    df["madre_normalizada"] = df["madre_normalizada"].replace("", pd.NA).fillna(df["Madre"].map(_normalizar_texto))
                    df["padre_normalizado"] = df["padre_normalizado"].replace("", pd.NA).fillna(df["Padre"].map(_normalizar_texto))
                    df["identificador_registro"] = (
                        df["Pestaña"].astype(str)
                        + "|"
                        + df["Tomo"].astype(str)
                        + "|"
                        + df["Folio"].astype(str)
                        + "|"
                        + df["Partida"].astype(str)
                    )
                    df["clave_registro"] = df["identificador_registro"]
                    registros.append(df)

                if registros:
                    self.registros_df = pd.concat(registros, ignore_index=True)
                    if "clave_registro" not in self.registros_df.columns:
                        self.registros_df["clave_registro"] = (
                            self.registros_df["Pestaña"].astype(str)
                            + "|"
                            + self.registros_df["Tomo"].astype(str)
                            + "|"
                            + self.registros_df["Folio"].astype(str)
                            + "|"
                            + self.registros_df["Partida"].astype(str)
                        )
                    self.registros_df["madre_normalizada"] = self.registros_df["Madre"].map(_normalizar_texto)
                    self.registros_df["padre_normalizado"] = self.registros_df["Padre"].map(_normalizar_texto)
                    self.registros_df["nombre_normalizado"] = self.registros_df["Nombre"].map(_normalizar_texto)
                    self.registros_df["comunidad_normalizada"] = self.registros_df["Comunidad"].map(_normalizar_texto)

                    self.lbl_estado.config(text=f"Registros cargados: {len(self.registros_df)}")
                    self.ejecutar_busqueda()
                else:
                    self.lbl_estado.config(text="No se encontraron registros en el archivo maestro.")
        except Exception as error:
            messagebox.showerror("Error al cargar registros", f"No se pudo leer el archivo maestro:\n{error}")

    def limpiar_busqueda(self):
        self.entry_termino.delete(0, tk.END)
        self.ejecutar_busqueda()

    def ejecutar_busqueda(self):
        if self.registros_df.empty:
            self.resultados_df = pd.DataFrame()
            self._poblar_tabla(pd.DataFrame())
            self._dibujar_mapa_vacio()
            return

        termino = _normalizar_texto(self.entry_termino.get())
        base = self.registros_df.copy()

        if termino:
            columnas_busqueda = [
                "nombre_normalizado",
                "madre_normalizada",
                "padre_normalizado",
                "comunidad_normalizada",
            ]
            mascara = pd.Series(False, index=base.index)
            for columna in columnas_busqueda:
                mascara = mascara | base[columna].fillna("").astype(str).str.contains(termino, regex=False, na=False)
            base = base[mascara].copy()

        if base.empty:
            self.resultados_df = pd.DataFrame()
            self.lbl_estado.config(text="No se encontraron coincidencias.")
            self._poblar_tabla(base)
            self._dibujar_mapa_vacio()
            return

        base = base.sort_values(["Madre", "anio_nacimiento", "Nombre"], na_position="last").reset_index(drop=True)
        self.resultados_df = base.copy()

        self._poblar_tabla(base)
        self._dibujar_familia_activa(base)

    def _poblar_tabla(self, df):
        for item in self.tabla.get_children():
            self.tabla.delete(item)

        self.resultados_df = df.reset_index(drop=True).copy()

        if self.resultados_df.empty:
            self.lbl_estado.config(text="No hay resultados para mostrar.")
            return

        for _, fila in self.resultados_df.iterrows():
            values = (
                fila["Pestaña"],
                fila["Tomo"],
                fila["Folio"],
                fila["Partida"],
                fila["Nombre"],
                fila["Fecha Nacimiento"],
                fila["Madre"],
                fila["Padre"],
                fila["Comunidad"],
                fila["Fecha Inscripción"],
            )
            iid = fila["clave_registro"]
            self.tabla.insert("", tk.END, iid=iid, values=values)

        total = len(self.resultados_df)
        vinculadas = self.resultados_df["Madre"].astype(str).str.strip().ne("").sum()
        self.lbl_estado.config(text=f"Mostrando {total} resultados. Familias vinculadas: {int(vinculadas)}")

    def _familia_activa(self, df):
        familiares = df[df["Madre"].astype(str).str.strip() != ""].copy()
        if familiares.empty:
            return None, pd.DataFrame()

        grupos = familiares.groupby("madre_normalizada", sort=False)
        madre_elegida = None
        mejor_grupo = None

        for madre_norm, grupo in grupos:
            grupo = grupo.copy()
            años = grupo["anio_nacimiento"].dropna()
            if len(grupo) < 2:
                continue
            if años.empty:
                rango = 0
            else:
                rango = int(años.max() - años.min())
            if rango <= 5:
                if mejor_grupo is None or len(grupo) > len(mejor_grupo):
                    madre_elegida = madre_norm
                    mejor_grupo = grupo

        if madre_elegida is None:
            madre_elegida = familiares["madre_normalizada"].iloc[0]
            mejor_grupo = familiares[familiares["madre_normalizada"] == madre_elegida].copy()

        madre_original = mejor_grupo["Madre"].iloc[0]
        return madre_original, mejor_grupo.sort_values(["anio_nacimiento", "Nombre"], na_position="last").reset_index(drop=True)

    def _dibujar_familia_activa(self, df):
        madre_nombre, familia = self._familia_activa(df)

        if madre_nombre is None or familia.empty:
            self._dibujar_mapa_vacio()
            return

        self.lbl_mapa.config(text=f"Familia activa: {madre_nombre} ({len(familia)} registros vinculados)")
        self.canvas.delete("all")

        ancho = self.canvas.winfo_width() or 360
        centro_x = ancho / 2
        centro_y = 70

        self.canvas.create_oval(
            centro_x - 40,
            centro_y - 22,
            centro_x + 40,
            centro_y + 22,
            fill="#2c3e50",
            outline="#1f2937",
            width=2,
            tags=("madre", "nodo"),
        )
        self.canvas.create_text(centro_x, centro_y, text="MADRE", fill="white", font=("Arial", 10, "bold"))
        self.canvas.create_text(centro_x, centro_y + 30, text=madre_nombre, fill="#1f2937", font=("Arial", 9, "bold"))

        hijos = familia.drop_duplicates(subset=["clave_registro"]).copy()
        hijos = hijos.sort_values(["anio_nacimiento", "Nombre"], na_position="last").reset_index(drop=True)

        if hijos.empty:
            self.canvas.create_text(centro_x, centro_y + 80, text="Sin hijos vinculados detectados", fill="#95a5a6", font=("Arial", 9, "italic"))
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))
            return

        columnas = 2
        paso_x = 130
        paso_y = 105

        for indice, fila in hijos.iterrows():
            columna = indice % columnas
            fila_idx = indice // columnas
            x = 70 + columna * paso_x if columna == 0 else 290
            y = 150 + fila_idx * paso_y

            node_id = self.canvas.create_oval(
                x - 26,
                y - 18,
                x + 26,
                y + 18,
                fill="#3498db",
                outline="#21618c",
                width=2,
                tags=(f"nodo_{fila['clave_registro']}", "nodo", "hijo"),
            )
            self.canvas.create_text(x, y, text="HIJO", fill="white", font=("Arial", 8, "bold"))
            self.canvas.create_text(x, y + 24, text=f"{fila['Nombre']}", fill="#1f2937", font=("Arial", 8, "bold"))
            self.canvas.create_text(x, y + 40, text=f"{int(fila['anio_nacimiento']) if pd.notna(fila['anio_nacimiento']) else 'Sin año'}", fill="#566573", font=("Arial", 7))

            self.canvas.create_line(
                centro_x + 15,
                centro_y + 24,
                x - 24,
                y - 16,
                arrow=tk.LAST,
                width=2,
                fill="#95a5a6",
            )

            self.canvas.tag_bind(node_id, "<Double-1>", lambda event, clave=fila["clave_registro"]: self._seleccionar_fila_por_clave(clave))

        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _dibujar_mapa_vacio(self):
        self.lbl_mapa.config(text="Seleccione una madre o busque un nombre para ver la red.")
        self.canvas.delete("all")
        self.canvas.create_text(
            180,
            160,
            text="Busque a una persona o a una madre\npara activar el mapa familiar",
            fill="#95a5a6",
            font=("Arial", 10, "italic"),
            justify=tk.CENTER,
        )

    def _mostrar_menu_contextual(self, event):
        item = self.tabla.identify_row(event.y)
        if not item:
            return
        self.tabla.selection_set(item)
        self.tabla.focus(item)
        self._seleccionar_fila_por_clave(item)
        self.menu_contextual.post(event.x_root, event.y_root)

    def _abrir_arbol_avanzado(self):
        item = self.tabla.focus()
        if not item or self.registros_df.empty:
            return

        fila = self.registros_df[self.registros_df["clave_registro"] == item]
        if fila.empty:
            return

        sujeto = fila.iloc[0].to_dict()
        sujeto["Fecha Nac."] = sujeto.get("Fecha Nacimiento")
        VentanaArbolAvanzado(self, sujeto, self.registros_df, callback_recargar=self.ejecutar_busqueda)

    def _seleccionar_fila_desde_tabla(self, event):
        item = self.tabla.focus()
        if not item:
            return
        self._seleccionar_fila_por_clave(item)

    def _seleccionar_fila_por_clave(self, clave_registro):
        if clave_registro not in self.tabla.get_children():
            return

        self.tabla.selection_set(clave_registro)
        self.tabla.focus(clave_registro)
        self.tabla.see(clave_registro)

        fila = self.resultados_df[self.resultados_df["clave_registro"] == clave_registro]
        if fila.empty:
            return

        madre = fila["Madre"].iloc[0]
        if madre:
            familia = self.resultados_df[self.resultados_df["madre_normalizada"] == _normalizar_texto(madre)].copy()
            if not familia.empty:
                self._dibujar_familia_activa(familia)

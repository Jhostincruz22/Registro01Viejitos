import tkinter as tk
from tkinter import ttk

from .utils import formatear_fecha_evento


class BloqueCapturaActa(ttk.LabelFrame):
    def __init__(self, parent, titulo, offset):
        super().__init__(parent, text=f" {titulo} ", padding="5")
        self.offset = offset

        self.lbl_partida = ttk.Label(self, text="PARTIDA N°: -", font=("Arial", 11, "bold"), foreground="#1b4f72")
        self.lbl_partida.pack(pady=5)

        self.frame_inputs = ttk.Frame(self)
        self.frame_inputs.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.frame_inputs.columnconfigure(1, weight=1)

        self.ent_insc = ttk.Entry(self.frame_inputs, font=("Arial", 10))
        self.ent_nac = ttk.Entry(self.frame_inputs, font=("Arial", 10))
        self.ent_nom = ttk.Entry(self.frame_inputs, font=("Arial", 10))
        self.ent_com = ttk.Entry(self.frame_inputs, font=("Arial", 10))
        self.ent_mad = ttk.Entry(self.frame_inputs, font=("Arial", 10))

        self.configurar_fila_con_ilegible("F. Inscripción:", self.ent_insc, 0, es_fecha=True)
        self.configurar_fila_con_ilegible("F. Nacimiento:", self.ent_nac, 1, es_fecha=True)
        self.configurar_fila_con_ilegible("Nombre Persona:", self.ent_nom, 2)
        self.configurar_fila_con_ilegible("Comunidad/Origen:", self.ent_com, 3)
        self.configurar_fila_con_ilegible("Nombre Madre:", self.ent_mad, 4)

        ttk.Label(self.frame_inputs, text="¿Registra Padre?:", font=("Arial", 9, "bold")).grid(row=5, column=0, sticky=tk.W, pady=5)
        frame_padre = ttk.Frame(self.frame_inputs)
        frame_padre.grid(row=5, column=1, sticky="ew", pady=5)
        self.ent_pad = ttk.Entry(frame_padre, font=("Arial", 10), state="disabled")
        self.v_pad = tk.IntVar()
        self.chk_pad = ttk.Checkbutton(frame_padre, variable=self.v_pad, command=self.alternar_padre)
        self.chk_pad.grid(row=0, column=0, padx=(0, 2))
        self.ent_pad.grid(row=0, column=1, sticky="ew")
        frame_padre.columnconfigure(1, weight=1)

        ttk.Label(self.frame_inputs, text="¿Marginación?:", font=("Arial", 9, "bold")).grid(row=6, column=0, sticky=tk.W, pady=5)
        frame_marginacion = ttk.Frame(self.frame_inputs)
        frame_marginacion.grid(row=6, column=1, sticky="ew", pady=5)
        self.ent_mar = ttk.Entry(frame_marginacion, font=("Arial", 10), state="disabled")
        self.v_mar = tk.IntVar()
        self.chk_mar = ttk.Checkbutton(frame_marginacion, variable=self.v_mar, command=self.alternar_marginacion)
        self.chk_mar.grid(row=0, column=0, padx=(0, 2))
        self.ent_mar.grid(row=0, column=1, sticky="ew")
        frame_marginacion.columnconfigure(1, weight=1)

    def configurar_fila_con_ilegible(self, texto_label, entry_obj, fila, es_fecha=False):
        ttk.Label(self.frame_inputs, text=texto_label, font=("Arial", 9, "bold")).grid(row=fila, column=0, sticky=tk.W, pady=5, padx=2)
        entry_obj.grid(row=fila, column=1, pady=5, sticky="ew", padx=(0, 2))

        var_ilegible = tk.IntVar()
        chk_ilegible = ttk.Checkbutton(
            self.frame_inputs,
            text="⚠️",
            variable=var_ilegible,
            command=lambda: self.alternar_ilegible(var_ilegible, entry_obj),
        )
        chk_ilegible.grid(row=fila, column=2, pady=5, sticky=tk.W)

        if not hasattr(self, "dict_ilegibles"):
            self.dict_ilegibles = {}
        self.dict_ilegibles[entry_obj] = var_ilegible

        if es_fecha:
            entry_obj.bind("<KeyRelease>", formatear_fecha_evento)

    def alternar_ilegible(self, var_chk, entry_obj):
        if var_chk.get() == 1:
            entry_obj.config(state="normal")
            entry_obj.delete(0, tk.END)
            entry_obj.insert(0, "[ILEGIBLE]")
            entry_obj.config(state="disabled")
        else:
            entry_obj.config(state="normal")
            entry_obj.delete(0, tk.END)

    def alternar_padre(self):
        if self.v_pad.get() == 1:
            self.ent_pad.config(state="normal")
            self.ent_pad.focus()
        else:
            self.ent_pad.delete(0, tk.END)
            self.ent_pad.config(state="disabled")

    def alternar_marginacion(self):
        if self.v_mar.get() == 1:
            self.ent_mar.config(state="normal")
            self.ent_mar.focus()
        else:
            self.ent_mar.delete(0, tk.END)
            self.ent_mar.config(state="disabled")

    def limpiar(self):
        for entry, var_il in self.dict_ilegibles.items():
            entry.config(state="normal")
            entry.delete(0, tk.END)
            var_il.set(0)

        self.ent_pad.config(state="normal")
        self.ent_pad.delete(0, tk.END)
        self.ent_pad.config(state="disabled")
        self.v_pad.set(0)

        self.ent_mar.config(state="normal")
        self.ent_mar.delete(0, tk.END)
        self.ent_mar.config(state="disabled")
        self.v_mar.set(0)

    def set_partida(self, numero_partida):
        self.lbl_partida.config(text=f"PARTIDA N°: {numero_partida}")

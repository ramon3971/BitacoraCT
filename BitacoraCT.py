import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import openpyxl
from openpyxl import Workbook
import os
import re
import pyperclip
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class BitacoraCTApp:
    def __init__(self, root):
        self.root = root
        self.version = "1.0.0"
        self.ruta_base_datos = r"\\mexhome03\Data\Prototype Engineering\Public\MC Front End\Mold\PROCESOS MOLDEO\ramon\01- Documentos\Bitacora CT\Bitacora CT.xlsx"
        self.id_a_modificar = None
        
        # Definir las listas para los Combobox
        self.combo_list_Equipo = ["Towa20", "Towa21", "Towa22", "Towa23", "Towa24", "Towa25",
                                 "Towa26", "Towa27", "Towa28", "Towa29", "Towa30", "Towa31",
                                 "Towa32", "Towa33", "Otro"]
        self.combo_list_Turno = ["A", "B", "C", "D"]
        self.combo_list_RCA = ["Si", "No", "Pendiente"]
        self.combo_list_Estatus = ["Quitar Hold y Mover", "Descontar Quitar Hold y Mover", "Para NCMR"]
        
        self.configurar_interfaz()
        self.crear_widgets()
        self.actualizar_treeview()
        self.centrar_ventana()

    def configurar_interfaz(self):
        self.root.geometry("1200x650")
        self.root.title(f"Bitácora CT Hold - v{self.version}")
        
        # Configuración del tema
        style = ttk.Style(self.root)
        self.root.option_add("*tearOff", False)
        self.root.tk.call("source", self.recurso_relativo("forest-light.tcl"))
        self.root.tk.call("source", self.recurso_relativo("forest-dark.tcl"))
        style.theme_use("forest-dark")

        # Frame principal
        self.frame = ttk.Frame(self.root)
        self.frame.pack(fill="both", expand=True, padx=10, pady=10)

    def recurso_relativo(self, ruta):
        """Para asegurar compatibilidad en .exe"""
        import sys
        import os
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, ruta)
        return ruta

    def crear_widgets(self):
        self.crear_formulario()
        self.crear_treeview()
        self.crear_botones_adicionales()

    def crear_formulario(self):
        # Frame principal del formulario
        widgets_frame = ttk.LabelFrame(self.frame, text="Insertar Evento")
        widgets_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        
        # Configurar grid para mejor distribución
        widgets_frame.columnconfigure(0, weight=1)
        widgets_frame.columnconfigure(1, weight=1)
        
        # Frame para campos del formulario
        campos_frame = ttk.Frame(widgets_frame)
        campos_frame.grid(row=0, column=0, columnspan=2, sticky="nsew", pady=(0, 10))
        
        # Lote
        self.lot_entry = ttk.Entry(campos_frame)
        self.lot_entry.insert(0, "Lote")
        self.lot_entry.bind("<FocusIn>", lambda e: self.lot_entry.delete(0, "end") if self.lot_entry.get() == "Lote" else None)
        self.lot_entry.grid(row=0, column=0, padx=5, pady=5, sticky="ew", columnspan=2)

        # Part ID
        self.part_id_entry = ttk.Entry(campos_frame)
        self.part_id_entry.insert(0, "Part ID")
        self.part_id_entry.bind("<FocusIn>", lambda e: self.part_id_entry.delete(0, "end") if self.part_id_entry.get() == "Part ID" else None)
        self.part_id_entry.grid(row=1, column=0, padx=5, pady=5, sticky="ew", columnspan=2)

        # Equipo (Combobox)
        self.equipo_combo = ttk.Combobox(campos_frame, values=self.combo_list_Equipo)
        self.equipo_combo.set("Equipo")
        self.equipo_combo.grid(row=2, column=0, padx=5, pady=5, sticky="ew", columnspan=2)

        # CT y CP
        ct_cp_frame = ttk.Frame(campos_frame)
        ct_cp_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5)
        ct_cp_frame.columnconfigure(0, weight=1)
        ct_cp_frame.columnconfigure(1, weight=1)

        self.ct_entry = ttk.Entry(ct_cp_frame, justify="center")
        self.ct_entry.insert(0, "CT")
        self.ct_entry.bind("<FocusIn>", lambda e: self.ct_entry.delete(0, "end") if self.ct_entry.get() == "CT" else None)
        self.ct_entry.grid(row=0, column=0, sticky="ew", padx=(0,2))

        self.cp_entry = ttk.Entry(ct_cp_frame, justify="center")
        self.cp_entry.insert(0, "CP")
        self.cp_entry.bind("<FocusIn>", lambda e: self.cp_entry.delete(0, "end") if self.cp_entry.get() == "CP" else None)
        self.cp_entry.grid(row=0, column=1, sticky="ew", padx=(2,0))

        # Comentario
        ttk.Label(campos_frame, text="Comentario:").grid(row=4, column=0, padx=5, sticky="w", columnspan=2)
        self.cometario_entry = tk.Text(campos_frame, height=6, width=30, wrap="word")
        self.cometario_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

        # PCB
        self.pcb_entry = ttk.Entry(campos_frame, justify="center")
        self.pcb_entry.insert(0, "PCB")
        self.pcb_entry.bind("<FocusIn>", lambda e: self.pcb_entry.delete(0, "end") if self.pcb_entry.get() == "PCB" else None)
        self.pcb_entry.grid(row=6, column=0, sticky="ew", pady=5, columnspan=2)

        # Fecha y Hora
        fecha_hora_frame = ttk.Frame(campos_frame)
        fecha_hora_frame.grid(row=7, column=0, columnspan=2, sticky="ew", pady=5)
        fecha_hora_frame.columnconfigure(0, weight=1)
        fecha_hora_frame.columnconfigure(1, weight=1)

        ttk.Label(fecha_hora_frame, text="Fecha:").grid(row=0, column=0, padx=(5,2), sticky="w")
        ttk.Label(fecha_hora_frame, text="Hora (24h):").grid(row=0, column=1, padx=(2,5), sticky="w")
        
        self.calendario = DateEntry(fecha_hora_frame, date_pattern="mm/dd/yyyy")
        self.calendario.grid(row=1, column=0, padx=(5,2), sticky="ew")
        
        self.hora_entry = ttk.Entry(fecha_hora_frame)
        self.hora_entry.grid(row=1, column=1, padx=(2,5), sticky="ew")

        # Turno y RCA
        turno_rca_frame = ttk.Frame(campos_frame)
        turno_rca_frame.grid(row=8, column=0, columnspan=2, sticky="ew", pady=5)
        turno_rca_frame.columnconfigure(0, weight=1)
        turno_rca_frame.columnconfigure(1, weight=1)

        self.turno_combo = ttk.Combobox(turno_rca_frame, values=self.combo_list_Turno)
        self.turno_combo.set("Turno")
        self.turno_combo.grid(row=0, column=0, padx=(5,2), sticky="ew")

        self.rca_combo = ttk.Combobox(turno_rca_frame, values=self.combo_list_RCA)
        self.rca_combo.set("RCA")
        self.rca_combo.grid(row=0, column=1, padx=(2,5), sticky="ew")

        # Estatus
        self.estatus_combo = ttk.Combobox(campos_frame, values=self.combo_list_Estatus)
        self.estatus_combo.set("Estatus")
        self.estatus_combo.grid(row=9, column=0, columnspan=2, sticky="ew", pady=5)

        # Separador visual
        ttk.Separator(widgets_frame).grid(row=1, column=0, columnspan=2, sticky="ew", pady=5)

        # Frame para botones principales (ahora en la parte inferior)
        self.button_frame = ttk.Frame(widgets_frame)
        self.button_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(0, 5))

        # Botones principales
        self.insert_button = ttk.Button(self.button_frame, text="Insertar Evento", command=self.guardar_datos, style="Accent.TButton")
        self.insert_button.pack(side="left", padx=2, expand=True)

        ttk.Button(self.button_frame, text="Limpiar", command=self.limpiar_campos).pack(side="left", padx=2, expand=True)
        ttk.Button(self.button_frame, text="Buscar", command=self.buscar_por_lote).pack(side="left", padx=2, expand=True)

        # Frame para botones de gráficas
        self.graph_button_frame = ttk.Frame(widgets_frame)
        self.graph_button_frame.grid(row=3, column=0, columnspan=2, sticky="ew")

        # Botones de gráficas
        ttk.Button(self.graph_button_frame, text="Gráfica Turnos", command=self.generar_grafica_turnos, style="Accent.TButton").pack(side="left", padx=2, expand=True)
        ttk.Button(self.graph_button_frame, text="Gráfica Equipos", command=self.generar_grafica_equipos, style="Accent.TButton").pack(side="left", padx=2, expand=True)

        # Botón de guardar cambios (oculto inicialmente)
        self.guardar_button = ttk.Button(self.button_frame, text="Guardar Cambios", command=self.guardar_cambios, style="Accent.TButton")
        self.guardar_button.pack(side="left", padx=2, expand=True)
        self.guardar_button.pack_forget()

    def crear_treeview(self):
        # Frame para el Treeview
        treeFrame = ttk.LabelFrame(self.frame, text="Registros")
        treeFrame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")  

        # Scrollbar para el Treeview
        treescroll = ttk.Scrollbar(treeFrame)
        treescroll.pack(side="right", fill="y")

        # Configurar el Treeview
        cols = ("Lote", "Part ID", "Equipo", "Comentario", "Estatus")
        self.treeview = ttk.Treeview(
            treeFrame,
            show="headings",
            yscrollcommand=treescroll.set,
            columns=cols,
            height=12,
            selectmode="browse"
        )

        # Configurar columnas
        self.treeview.column("Lote", anchor="w", width=80)
        self.treeview.column("Part ID", anchor="w", width=80)  
        self.treeview.column("Equipo", anchor="w", width=80)
        self.treeview.column("Comentario", anchor="w", width=200)
        self.treeview.column("Estatus", anchor="w", width=200)

        # Configurar encabezados
        for col in cols:
            self.treeview.heading(col, text=col)

        self.treeview.pack(fill="both", expand=True)
        treescroll.config(command=self.treeview.yview)

    def crear_botones_adicionales(self):
        # Frame para botones del treeview
        botones_tree_frame = ttk.Frame(self.treeview.master)
        botones_tree_frame.pack(fill="x", pady=5)

        # Botón para copiar selección
        ttk.Button(
            botones_tree_frame,
            text="Copiar Selección",
            command=self.copiar_seleccion,
            style="Accent.TButton"
        ).pack(side="left", padx=2, expand=True)

        # Botón para modificar evento
        ttk.Button(
            botones_tree_frame,
            text="Modificar",
            command=self.modificar_evento,
            style="Accent.TButton"
        ).pack(side="left", padx=2, expand=True)

    def centrar_ventana(self):
        self.root.update()
        self.root.minsize(1300, 650)
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")

    # ============ Funcionalidades principales ============

    def modificar_evento(self):
        """Cargar los datos del registro seleccionado en el formulario para editar"""
        seleccion = self.treeview.selection()
        
        if not seleccion:
            messagebox.showwarning("Advertencia", "Selecciona un registro para modificar")
            return

        item = self.treeview.item(seleccion[0])
        valores = item["values"]
        lote = valores[0]

        # Cargar todos los datos desde Excel
        datos = self.load_data()
        for fila in datos:
            if str(fila[1]) == str(lote):
                self.id_a_modificar = fila[0]  # Guardamos el ID del registro
                
                # Actualizar campos del formulario
                self.lot_entry.delete(0, tk.END)
                self.lot_entry.insert(0, str(fila[1]))

                self.part_id_entry.delete(0, tk.END)
                self.part_id_entry.insert(0, str(fila[2]))

                self.equipo_combo.set(fila[3])
                self.ct_entry.delete(0, tk.END)
                self.ct_entry.insert(0, str(fila[4]))

                self.cp_entry.delete(0, tk.END)
                self.cp_entry.insert(0, str(fila[5]))

                self.cometario_entry.delete("1.0", tk.END)
                self.cometario_entry.insert("1.0", str(fila[6]))

                self.pcb_entry.delete(0, tk.END)
                self.pcb_entry.insert(0, str(fila[7]))

                try:
                    self.calendario.set_date(fila[8])
                except:
                    pass

                self.hora_entry.delete(0, tk.END)
                if fila[9] is not None:
                    self.hora_entry.insert(0, str(fila[9]))

                self.turno_combo.set(fila[10])
                self.rca_combo.set(fila[11])
                self.estatus_combo.set(fila[12])

                # Ocultar botón Insertar y mostrar Guardar Cambios
                self.insert_button.pack_forget()
                self.guardar_button.pack(side="left", padx=2, expand=True)
                return

        messagebox.showerror("Error", "No se pudo encontrar el registro completo en el Excel")

    def guardar_cambios(self):
        """Actualizar un registro existente en el Excel"""
        if self.id_a_modificar is None:
            messagebox.showerror("Error", "No hay registro en modo edición")
            return

        try:
            wb = openpyxl.load_workbook(self.ruta_base_datos)
            ws = wb.active

            for fila in ws.iter_rows(min_row=2):
                if fila[0].value == self.id_a_modificar:
                    fila[1].value = self.lot_entry.get()
                    fila[2].value = self.part_id_entry.get()
                    fila[3].value = self.equipo_combo.get()
                    fila[4].value = self.ct_entry.get()
                    fila[5].value = self.cp_entry.get()
                    fila[6].value = self.cometario_entry.get("1.0", "end-1c")
                    fila[7].value = self.pcb_entry.get()
                    fila[8].value = self.calendario.get_date().strftime("%m/%d/%Y")
                    fila[9].value = self.hora_entry.get()
                    fila[10].value = self.turno_combo.get()
                    fila[11].value = self.rca_combo.get()
                    fila[12].value = self.estatus_combo.get()
                    break

            wb.save(self.ruta_base_datos)
            messagebox.showinfo("Éxito", f"Registro #{self.id_a_modificar} actualizado correctamente")

            # Resetear estado
            self.id_a_modificar = None
            self.limpiar_campos()
            self.actualizar_treeview()
            self.guardar_button.pack_forget()
            self.insert_button.pack(side="left", padx=2, expand=True)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la modificación:\n{str(e)}")

    def load_data(self):
        """Cargar datos desde el archivo Excel"""
        if not os.path.exists(self.ruta_base_datos):
            return []
        
        try:
            wb = openpyxl.load_workbook(self.ruta_base_datos)
            sheet = wb.active
            data = []
            
            # Comenzar desde la fila 2 (omitir encabezados)
            for row in sheet.iter_rows(min_row=2, values_only=True):
                # Filtrar filas vacías
                if any(cell is not None for cell in row):
                    data.append(row)
            
            return data
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel:\n{str(e)}")
            return []

    def validar_lot(self, lot):
        """Validar formato del Lote (7 dígitos + . + 1 dígito)"""
        return re.match(r'^\d{7}\.\d{1}$', lot) is not None

    def guardar_datos(self):
        """Guardar los datos en el archivo Excel"""
        # Validar campos obligatorios
        if not self.lot_entry.get() or not self.part_id_entry.get() or not self.cometario_entry.get("1.0", "end-1c"):
            messagebox.showerror("Error", "¡Completa los campos obligatorios (Lote, Part ID, Comentario)!")
            return

        # Validar formato del Lote
        if not self.validar_lot(self.lot_entry.get()):
            messagebox.showerror("Error", "Formato de Lote inválido. Debe ser: 7 dígitos + . + 1 dígito")
            return

        # Validar hora (opcional)
        hora_val = self.hora_entry.get()
        if hora_val:
            try:
                datetime.strptime(hora_val, "%H:%M")
            except ValueError:
                messagebox.showerror("Error", "Formato de hora inválido. Usa HH:MM (24 hrs)")
                return

        try:
            if not os.path.exists(self.ruta_base_datos):
                wb = Workbook()
                ws = wb.active
                ws.append([
                    "ID", "Lote", "Part ID", "Equipo", "CT", "CP", "Comentario",
                    "PCB", "Fecha", "Hora", "Turno", "RCA", "Estatus"
                ])
                wb.save(self.ruta_base_datos)
            
            wb = openpyxl.load_workbook(self.ruta_base_datos)
            ws = wb.active
            
            # Encontrar el último ID válido
            max_id = 0
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                if row[0].value is not None:
                    try:
                        current_id = int(row[0].value)
                        if current_id > max_id:
                            max_id = current_id
                    except (ValueError, TypeError):
                        continue
            
            siguiente_id = max_id + 1

            ws.append([
                siguiente_id,
                self.lot_entry.get(),
                self.part_id_entry.get(),
                self.equipo_combo.get(),
                self.ct_entry.get(),
                self.cp_entry.get(),
                self.cometario_entry.get("1.0", "end-1c"),
                self.pcb_entry.get(),
                self.calendario.get_date().strftime("%m/%d/%Y"),
                hora_val,
                self.turno_combo.get(),
                self.rca_combo.get(),
                self.estatus_combo.get()
            ])
            wb.save(self.ruta_base_datos)
            messagebox.showinfo("Éxito", f"Registro #{siguiente_id} guardado correctamente")
            self.limpiar_campos()
            self.actualizar_treeview()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")

    def limpiar_campos(self):
        """Limpiar todos los campos del formulario"""
        respuesta = messagebox.askyesno( "Confirmar Limpieza", "¿Estás seguro que deseas limpiar todos los campos? ")
    
        if respuesta:
            self.lot_entry.delete(0, tk.END)
            self.lot_entry.insert(0, "Lote")
            self.part_id_entry.delete(0, tk.END)
            self.part_id_entry.insert(0, "Part ID")
            self.equipo_combo.set("Equipo")
            self.ct_entry.delete(0, tk.END)
            self.ct_entry.insert(0, "CT")
            self.cp_entry.delete(0, tk.END)
            self.cp_entry.insert(0, "CP")
            self.pcb_entry.delete(0, tk.END)
            self.pcb_entry.insert(0, "PCB")
            self.cometario_entry.delete("1.0", tk.END)
            self.turno_combo.set("Turno")
            self.rca_combo.set("RCA")
            self.estatus_combo.set("Estatus")
            self.hora_entry.delete(0, tk.END)

    def guardar_datos(self):
        """Guardar los datos en el archivo Excel con confirmación"""
        # Validar campos obligatorios
        if not self.lot_entry.get() or not self.part_id_entry.get() or not self.cometario_entry.get("1.0", "end-1c"):
            messagebox.showerror("Error", "¡Completa los campos obligatorios (Lote, Part ID, Comentario)!")
            return

        # Validar formato del Lote
        if not self.validar_lot(self.lot_entry.get()):
            messagebox.showerror("Error", "Formato de Lote inválido. Debe ser: 7 dígitos + . + 1 dígito")
            return

        # Validar hora (opcional)
        hora_val = self.hora_entry.get()
        if hora_val:
            try:
                datetime.strptime(hora_val, "%H:%M")
            except ValueError:
                messagebox.showerror("Error", "Formato de hora inválido. Usa HH:MM (24 hrs)")
                return

        # Mostrar resumen antes de guardar
        resumen = (
            f"Lote: {self.lot_entry.get()}\n"
            f"Part ID: {self.part_id_entry.get()}\n"
            f"Equipo: {self.equipo_combo.get()}\n"
            f"Comentario: {self.cometario_entry.get('1.0', 'end-1c')[:50]}...\n\n"
            "¿Deseas guardar este registro?"
        )
        
        confirmar = messagebox.askyesno("Confirmar Registro", resumen)
        if not confirmar:
            return

        try:
            if not os.path.exists(self.ruta_base_datos):
                wb = Workbook()
                ws = wb.active
                ws.append([
                    "ID", "Lote", "Part ID", "Equipo", "CT", "CP", "Comentario",
                    "PCB", "Fecha", "Hora", "Turno", "RCA", "Estatus"
                ])
                wb.save(self.ruta_base_datos)
            
            wb = openpyxl.load_workbook(self.ruta_base_datos)
            ws = wb.active
            
            # Encontrar el último ID válido
            max_id = 0
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=1):
                if row[0].value is not None:
                    try:
                        current_id = int(row[0].value)
                        if current_id > max_id:
                            max_id = current_id
                    except (ValueError, TypeError):
                        continue
            
            siguiente_id = max_id + 1

            ws.append([
                siguiente_id,
                self.lot_entry.get(),
                self.part_id_entry.get(),
                self.equipo_combo.get(),
                self.ct_entry.get(),
                self.cp_entry.get(),
                self.cometario_entry.get("1.0", "end-1c"),
                self.pcb_entry.get(),
                self.calendario.get_date().strftime("%m/%d/%Y"),
                hora_val,
                self.turno_combo.get(),
                self.rca_combo.get(),
                self.estatus_combo.get()    
            ])
            wb.save(self.ruta_base_datos)
            
            # Mensaje de confirmación después de guardar
            mensaje_exito = (
                f"Registro #{siguiente_id} guardado correctamente\n\n"
                f"Lote: {self.lot_entry.get()}\n"
                f"Part ID: {self.part_id_entry.get()}\n"
                f"Equipo: {self.equipo_combo.get()}\n"
                f"Fecha: {self.calendario.get_date().strftime('%d/%m/%Y')}"
            )
            messagebox.showinfo("Éxito", mensaje_exito)
            
            self.limpiar_campos()
            self.actualizar_treeview()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")

    def actualizar_treeview(self):
        """Actualizar el Treeview con los datos del Excel"""
        for item in self.treeview.get_children():
            self.treeview.delete(item)
        
        try:
            data = self.load_data()
            for row in data:
                # Mostrar solo las columnas seleccionadas en el Treeview
                self.treeview.insert("", "end", values=(row[1], row[2], row[3], row[6], row[12]))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos:\n{str(e)}")

    def copiar_seleccion(self):
        """Copiar los datos seleccionados al portapapeles"""
        seleccion = self.treeview.selection()
        if not seleccion:
            messagebox.showwarning("Advertencia", "No hay datos seleccionados")
            return

        item = self.treeview.item(seleccion[0])
        texto = "\t".join(str(val) for val in item["values"])
        pyperclip.copy(texto)
        messagebox.showinfo("Éxito", "Datos copiados al portapapeles")

    def buscar_por_lote(self):
        """Buscar registros por número de lote"""
        lote_buscado = self.lot_entry.get().strip()
        if not lote_buscado or lote_buscado == "Lote":
            messagebox.showwarning("Advertencia", "Ingresa un Lote para buscar")
            return
        
        for item in self.treeview.get_children():
            if self.treeview.item(item)["values"][0] == lote_buscado:
                self.treeview.selection_set(item)
                self.treeview.focus(item)
                self.treeview.see(item)
                return
        
        messagebox.showinfo("Información", f"No se encontró el Lote: {lote_buscado}")

    # ============ Funciones para gráficas ============

    def generar_grafica_turnos(self):
        """Generar gráfica de frecuencia de turnos por mes"""
        try:
            # Crear ventana emergente
            ventana_grafica = tk.Toplevel(self.root)
            ventana_grafica.title("Gráfica de Turnos por Mes")
            ventana_grafica.geometry("800x600")
            
            # Frame para controles
            control_frame = ttk.Frame(ventana_grafica)
            control_frame.pack(pady=10, fill="x", padx=20)
            
            # Etiqueta y combobox para selección de mes
            ttk.Label(control_frame, text="Seleccionar Mes:").grid(row=0, column=0, padx=5)
            
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            mes_combo = ttk.Combobox(control_frame, values=meses, width=15)
            mes_combo.grid(row=0, column=1, padx=5)
            mes_combo.current(0)  # Seleccionar enero por defecto
            
            # Etiqueta y combobox para selección de año
            ttk.Label(control_frame, text="Seleccionar Año:").grid(row=0, column=2, padx=5)
            
            # Obtener el año actual
            año_actual = datetime.now().year
            años = [str(año_actual - 1), str(año_actual), str(año_actual + 1)]
            año_combo = ttk.Combobox(control_frame, values=años, width=8)
            año_combo.grid(row=0, column=3, padx=5)
            año_combo.set(str(año_actual))  # Seleccionar año actual por defecto
            
            # Botón para generar gráfica
            btn_generar = ttk.Button(control_frame, text="Generar Gráfica", 
                                    command=lambda: self.mostrar_grafica_turnos(mes_combo.get(), año_combo.get(), ventana_grafica),
                                    style="Accent.TButton")
            btn_generar.grid(row=0, column=4, padx=10)
            
            # Frame para la gráfica
            grafica_frame = ttk.Frame(ventana_grafica)
            grafica_frame.pack(fill="both", expand=True, padx=20, pady=10)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la ventana de gráficas:\n{str(e)}")

    def mostrar_grafica_turnos(self, mes_seleccionado, año_seleccionado, ventana):
        """Mostrar gráfica de frecuencia de turnos para el mes seleccionado"""
        try:
            # Obtener datos del Excel
            data = self.load_data()
            if not data:
                messagebox.showinfo("Información", "No hay datos para mostrar")
                return
            
            # Obtener número de mes (1-12)
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                     "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            num_mes = meses.index(mes_seleccionado) + 1

            # Inicializar contadores
            turnos = {"A": 0, "B": 0, "C": 0, "D": 0}
            total_eventos = 0
            
            for fila in data:
                try:
                    fecha = fila[8]  # Columna de Fecha

                    # Si es cadena, convertirla a datetime
                    if isinstance(fecha, str):
                        fecha = datetime.strptime(fecha, "%m/%d/%Y")

                    # Verificar mes y año
                    if fecha.month == num_mes and fecha.year == int(año_seleccionado):
                        turno = str(fila[10]).strip().upper()  # Normalizar turno
                        if turno in turnos:
                            turnos[turno] += 1
                            total_eventos += 1
                except Exception as e:
                    print(f"Error al procesar fila: {fila}\n{e}")
                    continue

            if total_eventos == 0:
                messagebox.showinfo("Información", f"No hay eventos para {mes_seleccionado} {año_seleccionado}")
                return

            # Crear figura
            fig, ax = plt.subplots(figsize=(8, 5))
            turnos_keys = list(turnos.keys())
            valores = list(turnos.values())

            bars = ax.bar(turnos_keys, valores, color=['#4e79a7', '#f28e2c', '#e15759', '#76b7b2'])

            for bar in bars:
                height = bar.get_height()
                ax.annotate(f'{height}',
                            xy=(bar.get_x() + bar.get_width() / 2, height),
                            xytext=(0, 3),
                            textcoords="offset points",
                            ha='center', va='bottom')

            ax.set_title(f'Frecuencia de Turnos - {mes_seleccionado} {año_seleccionado}', fontsize=14)
            ax.set_xlabel('Turno', fontsize=12)
            ax.set_ylabel('Cantidad de Eventos', fontsize=12)
            ax.set_ylim(0, max(valores) * 1.2)
            ax.grid(axis='y', linestyle='--', alpha=0.7)
            plt.tight_layout()

            # Limpiar frame anterior
            for widget in ventana.winfo_children():
                if isinstance(widget, ttk.Frame) and widget.winfo_name() == "!frame2":
                    widget.destroy()

            grafica_frame = ttk.Frame(ventana, name="frame2")
            grafica_frame.pack(fill="both", expand=True, padx=20, pady=10)

            canvas = FigureCanvasTkAgg(fig, master=grafica_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar la gráfica:\n{str(e)}")

    def generar_grafica_equipos(self):
        """Ventana para seleccionar mes y año para la gráfica por equipos"""
        try:
            ventana_equipos = tk.Toplevel(self.root)
            ventana_equipos.title("Gráfica por Equipos")
            ventana_equipos.geometry("900x600")

            control_frame = ttk.Frame(ventana_equipos)
            control_frame.pack(pady=10, fill="x", padx=20)

            # Selección de mes
            ttk.Label(control_frame, text="Mes:").grid(row=0, column=0, padx=5)
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                     "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            mes_combo = ttk.Combobox(control_frame, values=meses, width=15)
            mes_combo.grid(row=0, column=1, padx=5)
            mes_combo.current(0)

            # Selección de año
            ttk.Label(control_frame, text="Año:").grid(row=0, column=2, padx=5)
            año_actual = datetime.now().year
            años = [str(año_actual - 1), str(año_actual), str(año_actual + 1)]
            año_combo = ttk.Combobox(control_frame, values=años, width=8)
            año_combo.grid(row=0, column=3, padx=5)
            año_combo.set(str(año_actual))

            # Botón generar
            btn_generar = ttk.Button(control_frame, text="Generar Gráfica",
                                     command=lambda: self.mostrar_grafica_equipos(mes_combo.get(), año_combo.get(), ventana_equipos),
                                     style="Accent.TButton")
            btn_generar.grid(row=0, column=4, padx=10)

            # Frame para la gráfica
            grafica_frame = ttk.Frame(ventana_equipos)
            grafica_frame.pack(fill="both", expand=True, padx=20, pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo crear la ventana de gráficas:\n{str(e)}")

    def mostrar_grafica_equipos(self, mes_seleccionado, año_seleccionado, ventana):
        """Genera gráfica de eventos por equipo, separados por turno"""
        try:
            data = self.load_data()
            if not data:
                messagebox.showinfo("Información", "No hay datos para mostrar")
                return

            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", 
                     "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            num_mes = meses.index(mes_seleccionado) + 1

            # Diccionario: equipo -> {turno -> cantidad}
            equipos_data = {}

            for fila in data:
                try:
                    fecha = fila[8]
                    if isinstance(fecha, str):
                        fecha = datetime.strptime(fecha, "%m/%d/%Y")
                    if fecha.month != num_mes or fecha.year != int(año_seleccionado):
                        continue

                    equipo = str(fila[3]).strip()
                    turno = str(fila[10]).strip().upper()

                    if equipo and turno in ["A", "B", "C", "D"]:
                        if equipo not in equipos_data:
                            equipos_data[equipo] = {"A": 0, "B": 0, "C": 0, "D": 0}
                        equipos_data[equipo][turno] += 1

                except Exception as e:
                    print(f"Error al procesar fila: {fila}\n{e}")
                    continue

            if not equipos_data:
                messagebox.showinfo("Información", f"No hay eventos para {mes_seleccionado} {año_seleccionado}")
                return

            # Preparar datos para gráfica
            equipos = sorted(equipos_data.keys())
            turnos = ["A", "B", "C", "D"]
            colores = ['#4e79a7', '#f28e2c', '#e15759', '#76b7b2']

            valores_turnos = {t: [equipos_data[e].get(t, 0) for e in equipos] for t in turnos}

            fig, ax = plt.subplots(figsize=(10, 6))

            bottom = [0] * len(equipos)
            for i, turno in enumerate(turnos):
                ax.bar(equipos, valores_turnos[turno], bottom=bottom, label=f'Turno {turno}', color=colores[i])
                bottom = [sum(x) for x in zip(bottom, valores_turnos[turno])]

            ax.set_title(f'Eventos por Equipo y Turno - {mes_seleccionado} {año_seleccionado}', fontsize=14)
            ax.set_xlabel("Equipo", fontsize=12)
            ax.set_ylabel("Cantidad de Eventos", fontsize=12)
            ax.legend()
            ax.set_xticklabels(equipos, rotation=45, ha='right')
            ax.grid(axis='y', linestyle='--', alpha=0.5)

            plt.tight_layout()

            # Mostrar en ventana
            grafica_frame = ttk.Frame(ventana)
            grafica_frame.pack(fill="both", expand=True, padx=20, pady=10)

            canvas = FigureCanvasTkAgg(fig, master=grafica_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except Exception as e:
            messagebox.showerror("Error", f"No se pudo generar la gráfica:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = BitacoraCTApp(root)
    root.mainloop()

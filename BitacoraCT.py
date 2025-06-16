import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
import os

class FormularioHoldCT:
    def __init__(self, root):
        self.root = root
        self.root.title("Bitácora CT Hold")
        self.root.geometry("800x750")
        
        # Configuración del archivo
        self.excel_file = "Python Bitácora CT Hold.xlsm"
        self.sheet_name = "Data"
        self.header_row = 4  # Fila de encabezados
        self.start_data_row = 5  # Fila donde comienzan los datos
        
        # Variables para combobox
        self.turno_opciones = ["A", "B", "C", "D"]
        self.status_opciones = [
            "Quitar Hold y Mover",
            "Descontar Quitar Hold y Mover para NCMR"
        ]

        # Crear interfaz
        self.crear_interfaz()
        
        # Inicializar archivo Excel
        self.inicializar_archivo()
        
        # Cargar próximo ID
        self.cargar_proximo_id()
    
    def crear_interfaz(self):
        """Crea todos los elementos de la interfaz gráfica"""
        # Frame principal con scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas y scrollbar
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Título
        ttk.Label(scrollable_frame, text="Bitácora CT Hold", font=('Helvetica', 16, 'bold')).grid(
            row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Campos del formulario
        campos = [
            ("ID:", "id_entry", False),
            ("Lot*:", "lot_entry", True),
            ("Part ID*:", "part_id_entry", True),
            ("Equipo*:", "equipo_entry", True),
            ("CT:", "ct_entry", False),
            ("CP:", "cp_entry", False),
            ("Comentario:", "comentario_text", False),
            ("PCB:", "pcb_entry", False),
            ("Fecha del evento*:", "fecha_entry", True),
            ("Hora del evento*:", "hora_entry", True),
            ("Turno*:", "turno_combobox", True),
            ("RCA:", "rca_entry", False),
            ("Status*:", "status_combobox", True)
        ]
        
        for i, (label, var_name, obligatorio) in enumerate(campos, start=1):
            # Etiqueta con asterisco si es obligatorio
            lbl_text = label.replace("*", "") + ("*" if obligatorio else "")
            ttk.Label(scrollable_frame, text=lbl_text).grid(row=i, column=0, sticky="e", padx=5, pady=5)
            
            if "comentario" in var_name:
                setattr(self, var_name, tk.Text(scrollable_frame, width=50, height=4))
                getattr(self, var_name).grid(row=i, column=1, sticky="w", pady=5)
            elif "combobox" in var_name:
                valores = self.turno_opciones if "turno" in var_name else self.status_opciones
                combobox = ttk.Combobox(scrollable_frame, values=valores, state="readonly", width=47)
                combobox.grid(row=i, column=1, sticky="w", pady=5)
                setattr(self, var_name, combobox)
            else:
                entry = ttk.Entry(scrollable_frame, width=50)
                entry.grid(row=i, column=1, sticky="w", pady=5)
                setattr(self, var_name, entry)
        
        # Separador
        ttk.Separator(scrollable_frame).grid(row=len(campos)+1, column=0, columnspan=2, pady=20, sticky="ew")
        
        # Botones
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=len(campos)+2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Guardar", command=self.guardar_datos).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Limpiar", command=self.limpiar_formulario).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Ver Registros", command=self.ver_registros).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Salir", command=self.root.quit).pack(side="left", padx=5)
        
        # Nota
        ttk.Label(scrollable_frame, text="* Campos obligatorios", font=('Helvetica', 8)).grid(
            row=len(campos)+3, column=0, columnspan=2, pady=(10, 0))
    
    def inicializar_archivo(self):
        """Inicializa el archivo Excel si no existe o está corrupto"""
        if not os.path.exists(self.excel_file):
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = self.sheet_name
                
                # Crear estructura inicial
                for _ in range(3):  # Filas de presentación
                    ws.append([])
                
                # Encabezados en fila 4
                encabezados = [
                    "ID", "Lot", "Part ID", "Equipo", "CT", "CP", 
                    "Comentario", "PCB", "Fecha", "Hora", "Turno", "RCA", "Status"
                ]
                ws.append(encabezados)
                
                # Formato para encabezados
                for col in range(1, len(encabezados)+1):
                    celda = ws.cell(row=self.header_row, column=col)
                    celda.font = celda.font.copy(bold=True)
                    celda.alignment = celda.alignment.copy(horizontal="center")
                
                # Guardar como .xlsm (macro-enabled)
                wb.save(self.excel_file)
                messagebox.showinfo("Información", f"Se creó nuevo archivo: {self.excel_file}")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear el archivo: {str(e)}")
    
    def cargar_proximo_id(self):
        """Carga el próximo ID disponible"""
        proximo_id = self.obtener_proximo_id()
        self.id_entry.delete(0, tk.END)
        self.id_entry.insert(0, str(proximo_id))
    
    def obtener_proximo_id(self):
        """Obtiene el próximo ID basado en el máximo existente"""
        try:
            wb = load_workbook(self.excel_file, read_only=True, keep_vba=True)
            ws = wb[self.sheet_name]
            
            max_id = 0
            for row in ws.iter_rows(min_row=self.start_data_row, max_col=1, values_only=True):
                if row[0] and str(row[0]).isdigit():
                    current_id = int(row[0])
                    if current_id > max_id:
                        max_id = current_id
            
            return max_id + 1 if max_id > 0 else 1
        except Exception as e:
            messagebox.showwarning("Advertencia", f"No se pudo leer ID máximo: {str(e)}")
            return 1
    
    def validar_campos(self):
        """Valida los campos obligatorios"""
        campos_obligatorios = {
            "Lot": self.lot_entry.get(),
            "Part ID": self.part_id_entry.get(),
            "Equipo": self.equipo_entry.get(),
            "Turno": self.turno_combobox.get(),
            "Status": self.status_combobox.get(),
            "Fecha": self.fecha_entry.get(),
            "Hora": self.hora_entry.get()
        }
        
        for campo, valor in campos_obligatorios.items():
            if not valor:
                messagebox.showerror("Error", f"El campo {campo} es obligatorio")
                widget = getattr(self, f"{campo.lower().replace(' ', '_')}_entry", None) or \
                         getattr(self, f"{campo.lower().replace(' ', '_')}_combobox", None)
                if widget:
                    widget.focus()
                return False
        return True
    
    def guardar_datos(self):
        """Guarda los datos en el archivo Excel, insertando en la parte superior"""
        if not self.validar_campos():
            return
        
        try:
            # Cargar archivo existente conservando macros
            wb = load_workbook(self.excel_file, keep_vba=True)
            
            # Verificar si existe la hoja Data
            if self.sheet_name not in wb.sheetnames:
                wb.create_sheet(self.sheet_name)
                # Crear encabezados si es nueva hoja
                ws = wb[self.sheet_name]
                encabezados = [
                    "ID", "Lot", "Part ID", "Equipo", "CT", "CP", 
                    "Comentario", "PCB", "Fecha", "Hora", "Turno", "RCA", "Status"
                ]
                for _ in range(3):  # Filas de presentación
                    ws.append([])
                ws.append(encabezados)
            else:
                ws = wb[self.sheet_name]
            
            # Insertar nueva fila después de los encabezados
            ws.insert_rows(self.start_data_row)
            
            # Preparar datos
            datos = [
                self.id_entry.get(),
                self.lot_entry.get(),
                self.part_id_entry.get(),
                self.equipo_entry.get(),
                self.ct_entry.get(),
                self.cp_entry.get(),
                self.comentario_text.get("1.0", tk.END).strip(),
                self.pcb_entry.get(),
                self.fecha_entry.get(),
                self.hora_entry.get(),
                self.turno_combobox.get(),
                self.rca_entry.get(),
                self.status_combobox.get()
            ]
            
            # Escribir datos en la nueva fila (fila 5)
            for col, valor in enumerate(datos, start=1):
                ws.cell(row=self.start_data_row, column=col, value=valor)
            
            # Guardar conservando macros
            wb.save(self.excel_file)
            messagebox.showinfo("Éxito", "Registro guardado correctamente")
            self.limpiar_formulario()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar: {str(e)}")
    
    def limpiar_formulario(self):
        """Limpia todos los campos del formulario"""
        # Guardar ID actual antes de limpiar
        current_id = self.id_entry.get()
        
        # Limpiar todos los campos excepto ID
        self.lot_entry.delete(0, tk.END)
        self.part_id_entry.delete(0, tk.END)
        self.equipo_entry.delete(0, tk.END)
        self.ct_entry.delete(0, tk.END)
        self.cp_entry.delete(0, tk.END)
        self.comentario_text.delete("1.0", tk.END)
        self.pcb_entry.delete(0, tk.END)
        self.fecha_entry.delete(0, tk.END)
        self.hora_entry.delete(0, tk.END)
        self.turno_combobox.set('')
        self.rca_entry.delete(0, tk.END)
        self.status_combobox.set('')
        
        # Restaurar ID (o asignar nuevo si estaba vacío)
        self.id_entry.delete(0, tk.END)
        if current_id and current_id.isdigit():
            next_id = int(current_id) + 1
            self.id_entry.insert(0, str(next_id))
        else:
            self.cargar_proximo_id()
        
        self.lot_entry.focus()
    
    def ver_registros(self):
        """Muestra una ventana con los registros existentes"""
        try:
            wb = load_workbook(self.excel_file, read_only=True, data_only=True)
            
            # Verificar si existe la hoja Data
            if self.sheet_name not in wb.sheetnames:
                messagebox.showwarning("Advertencia", f"No se encontró la hoja '{self.sheet_name}'")
                return
            
            ws = wb[self.sheet_name]
            
            # Crear nueva ventana
            registros_window = tk.Toplevel(self.root)
            registros_window.title("Registros Existentes")
            registros_window.geometry("1300x600")
            
            # Treeview para mostrar datos
            tree = ttk.Treeview(registros_window)
            
            # Configurar columnas
            encabezados = []
            for cell in ws[self.header_row]:
                encabezados.append(cell.value)
            
            tree["columns"] = encabezados
            tree.column("#0", width=0, stretch=tk.NO)
            
            for encabezado in encabezados:
                tree.column(encabezado, anchor=tk.W, width=120)
                tree.heading(encabezado, text=encabezado)
            
            # Agregar datos (mostrando los más recientes primero)
            for row in ws.iter_rows(min_row=self.start_data_row, values_only=True):
                tree.insert("", 0, values=row)  # Insertar al principio
            
            # Scrollbars
            y_scroll = ttk.Scrollbar(registros_window, orient="vertical", command=tree.yview)
            x_scroll = ttk.Scrollbar(registros_window, orient="horizontal", command=tree.xview)
            tree.configure(yscroll=y_scroll.set, xscroll=x_scroll.set)
            
            y_scroll.pack(side="right", fill="y")
            x_scroll.pack(side="bottom", fill="x")
            tree.pack(fill="both", expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los registros: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FormularioHoldCT(root)
    root.mainloop()

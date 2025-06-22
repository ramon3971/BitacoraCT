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
id_a_modificar = None  # Variable global para almacenar el ID del registro a modificar

# =============================================
# CONFIGURACIÓN - EDITA ESTA RUTA
RUTA_BASE_DATOS = r"\\mexhome03\Data\Prototype Engineering\Public\MC Front End\Mold\PROCESOS MOLDEO\ramon\01- Documentos\Bitacora CT\Bitacora CT.xlsx"
# =============================================


# Evento Modificar Variable global para almacenar el ID del registro a modificar
def modificar_evento():
    """Cargar los datos del registro seleccionado en el formulario para editar"""
    global id_a_modificar
    seleccion = treeview.selection()
    
    if not seleccion:
        messagebox.showwarning("Advertencia", "Selecciona un registro para modificar")
        return

    item = treeview.item(seleccion[0])
    valores = item["values"]
    lote = valores[0]

    # Cargar todos los datos desde Excel
    datos = load_data()
    for fila in datos:
        if str(fila[1]) == str(lote):
            id_a_modificar = fila[0]  # Guardamos el ID del registro
            lot_entry.delete(0, tk.END)
            lot_entry.insert(0, str(fila[1]))

            part_id_entry.delete(0, tk.END)
            part_id_entry.insert(0, str(fila[2]))

            equipo_combo.set(fila[3])
            ct_entry.delete(0, tk.END)
            ct_entry.insert(0, str(fila[4]))

            cp_entry.delete(0, tk.END)
            cp_entry.insert(0, str(fila[5]))

            cometario_entry.delete("1.0", tk.END)
            cometario_entry.insert("1.0", str(fila[6]))

            pcb_entry.delete(0, tk.END)
            pcb_entry.insert(0, str(fila[7]))

            try:
                calendario.set_date(fila[8])
            except:
                pass

            hora_entry.delete(0, tk.END)
            if fila[9] is not None:
                hora_entry.insert(0, str(fila[9]))

            turno_combo.set(fila[10])
            rca_combo.set(fila[11])
            estatus_combo.set(fila[12])

            # Ocultar botón Insertar y mostrar Guardar Cambios
            insert_button.pack_forget()
            guardar_button.pack(side="left", padx=2, expand=True)
            return

    messagebox.showerror("Error", "No se pudo encontrar el registro completo en el Excel")



# Guardas cambios en el archivo excel
def guardar_cambios():
    """Actualizar un registro existente en el Excel"""
    global id_a_modificar
    if id_a_modificar is None:
        messagebox.showerror("Error", "No hay registro en modo edición")
        return

    try:
        wb = openpyxl.load_workbook(RUTA_BASE_DATOS)
        ws = wb.active

        for fila in ws.iter_rows(min_row=2):
            if fila[0].value == id_a_modificar:
                fila[1].value = lot_entry.get()
                fila[2].value = part_id_entry.get()
                fila[3].value = equipo_combo.get()
                fila[4].value = ct_entry.get()
                fila[5].value = cp_entry.get()
                fila[6].value = cometario_entry.get("1.0", "end-1c")
                fila[7].value = pcb_entry.get()
                fila[8].value = calendario.get_date().strftime("%m/%d/%Y")
                fila[9].value = hora_entry.get()
                fila[10].value = turno_combo.get()
                fila[11].value = rca_combo.get()
                fila[12].value = estatus_combo.get()
                break

        wb.save(RUTA_BASE_DATOS)
        messagebox.showinfo("Éxito", f"Registro #{id_a_modificar} actualizado correctamente")

        # Resetear estado
        id_a_modificar = None
        limpiar_campos()
        actualizar_treeview()
        guardar_button.pack_forget()
        insert_button.pack(side="left", padx=2, expand=True)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar la modificación:\n{str(e)}")





def load_data():
    """Cargar datos desde el archivo Excel"""
    if not os.path.exists(RUTA_BASE_DATOS):
        return []
    
    try:
        wb = openpyxl.load_workbook(RUTA_BASE_DATOS)
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

def validar_lot(lot):
    """Validar formato del Lote (7 dígitos + . + 1 dígito)"""
    return re.match(r'^\d{7}\.\d{1}$', lot) is not None

def guardar_datos():
    """Guardar los datos en el archivo Excel"""
    # Validar campos obligatorios
    if not lot_entry.get() or not part_id_entry.get() or not cometario_entry.get("1.0", "end-1c"):
        messagebox.showerror("Error", "¡Completa los campos obligatorios (Lote, Part ID, Comentario)!")
        return

    # Validar formato del Lote
    if not validar_lot(lot_entry.get()):
        messagebox.showerror("Error", "Formato de Lote inválido. Debe ser: 7 dígitos + . + 1 dígito")
        return

    # Validar hora (opcional)
    hora_val = hora_entry.get()
    if hora_val:
        try:
            datetime.strptime(hora_val, "%H:%M")
        except ValueError:
            messagebox.showerror("Error", "Formato de hora inválido. Usa HH:MM (24 hrs)")
            return

    try:
        if not os.path.exists(RUTA_BASE_DATOS):
            wb = Workbook()
            ws = wb.active
            ws.append([
                "ID", "Lote", "Part ID", "Equipo", "CT", "CP", "Comentario",
                "PCB", "Fecha", "Hora", "Turno", "RCA", "Estatus"
            ])
            wb.save(RUTA_BASE_DATOS)
        
        wb = openpyxl.load_workbook(RUTA_BASE_DATOS)
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
            lot_entry.get(),
            part_id_entry.get(),
            equipo_combo.get(),
            ct_entry.get(),
            cp_entry.get(),
            cometario_entry.get("1.0", "end-1c"),
            pcb_entry.get(),
            calendario.get_date().strftime("%m/%d/%Y"),
            hora_val,
            turno_combo.get(),
            rca_combo.get(),
            estatus_combo.get()
        ])
        wb.save(RUTA_BASE_DATOS)
        messagebox.showinfo("Éxito", f"Registro #{siguiente_id} guardado correctamente")
        limpiar_campos()
        actualizar_treeview()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo guardar:\n{str(e)}")

def limpiar_campos():
    """Limpiar todos los campos del formulario"""
    lot_entry.delete(0, tk.END)
    lot_entry.insert(0, "Lote")
    part_id_entry.delete(0, tk.END)
    part_id_entry.insert(0, "Part ID")
    equipo_combo.set("Equipo")
    ct_entry.delete(0, tk.END)
    ct_entry.insert(0, "CT")
    cp_entry.delete(0, tk.END)
    cp_entry.insert(0, "CP")
    pcb_entry.delete(0, tk.END)
    pcb_entry.insert(0, "PCB")
    cometario_entry.delete("1.0", tk.END)
    turno_combo.set("Turno")
    rca_combo.set("RCA")
    estatus_combo.set("Estatus")
    hora_entry.delete(0, tk.END)

def actualizar_treeview():
    """Actualizar el Treeview con los datos del Excel"""
    for item in treeview.get_children():
        treeview.delete(item)
    
    try:
        data = load_data()
        for row in data:
            # Mostrar solo las columnas seleccionadas en el Treeview
            treeview.insert("", "end", values=(row[1], row[2], row[3], row[6], row[12]))
    except Exception as e:
        messagebox.showerror("Error", f"No se pudieron cargar los datos:\n{str(e)}")

def copiar_seleccion():
    """Copiar los datos seleccionados al portapapeles"""
    seleccion = treeview.selection()
    if not seleccion:
        messagebox.showwarning("Advertencia", "No hay datos seleccionados")
        return

    item = treeview.item(seleccion[0])
    texto = "\t".join(str(val) for val in item["values"])
    pyperclip.copy(texto)
    messagebox.showinfo("Éxito", "Datos copiados al portapapeles")

def buscar_por_lote():
    """Buscar registros por número de lote"""
    lote_buscado = lot_entry.get().strip()
    if not lote_buscado or lote_buscado == "Lote":
        messagebox.showwarning("Advertencia", "Ingresa un Lote para buscar")
        return
    
    for item in treeview.get_children():
        if treeview.item(item)["values"][0] == lote_buscado:
            treeview.selection_set(item)
            treeview.focus(item)
            treeview.see(item)
            return
    
    messagebox.showinfo("Información", f"No se encontró el Lote: {lote_buscado}")


# Genera gráfica de frecuencia de turnos por mes
def generar_grafica_turnos():
    """Generar gráfica de frecuencia de turnos por mes"""
    try:
        # Crear ventana emergente
        ventana_grafica = tk.Toplevel(root)
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
                                command=lambda: mostrar_grafica(mes_combo.get(), año_combo.get(), ventana_grafica),
                                style="Accent.TButton")
        btn_generar.grid(row=0, column=4, padx=10)
        
        # Frame para la gráfica
        grafica_frame = ttk.Frame(ventana_grafica)
        grafica_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo crear la ventana de gráficas:\n{str(e)}")

# Mostrar gráfica de frecuencia de turnos
def mostrar_grafica(mes_seleccionado, año_seleccionado, ventana):
    """Mostrar gráfica de frecuencia de turnos para el mes seleccionado"""
    try:
        # Obtener datos del Excel
        data = load_data()
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

# Generar_grafica_equipos
def generar_grafica_equipos():
    """Ventana para seleccionar mes y año para la gráfica por equipos"""
    try:
        ventana_equipos = tk.Toplevel(root)
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
                                 command=lambda: mostrar_grafica_por_equipos(mes_combo.get(), año_combo.get(), ventana_equipos),
                                 style="Accent.TButton")
        btn_generar.grid(row=0, column=4, padx=10)

    except Exception as e:
        messagebox.showerror("Error", f"No se pudo crear la ventana de gráficas:\n{str(e)}")

# Mostrar gráfica por equipos
def mostrar_grafica_por_equipos(mes_seleccionado, año_seleccionado, ventana):
    """Genera gráfica de eventos por equipo, separados por turno"""
    try:
        data = load_data()
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



# Crear ventana principal
root = tk.Tk()
root.geometry("1200x650")
root.title("Bitácora CT Hold")

# Configuración del tema Forest
style = ttk.Style(root)
root.option_add("*tearOff", False)
import sys
import os

# Para asegurar compatibilidad en .exe
def recurso_relativo(ruta):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, ruta)
    return ruta

root.tk.call("source", recurso_relativo("forest-light.tcl"))
root.tk.call("source", recurso_relativo("forest-dark.tcl"))
style.theme_use("forest-dark")

# Frame principal
frame = ttk.Frame(root)
frame.pack(fill="both", expand=True, padx=10, pady=10)

# Frame para el formulario
widgets_frame = ttk.LabelFrame(frame, text="Insertar Evento")
widgets_frame.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")

# Configurar columnas para que se distribuyan bien
widgets_frame.columnconfigure(0, weight=1)
widgets_frame.columnconfigure(1, weight=1)

# Listas para los Combobox
combo_list_Equipo = ["Towa20", "Towa21", "Towa22", "Towa23", "Towa24", "Towa25",
                     "Towa26", "Towa27", "Towa28", "Towa29", "Towa30", "Towa31",
                     "Towa32", "Towa33", "Otro"]
combo_list_Turno = ["A", "B", "C", "D"]
combo_list_RCA = ["Si", "No", "Pendiente"]
combo_list_Estatus = ["Quitar Hold y Mover", "Descartar Quitar Hold y Mover", "Para NCMR"]

# Lote
lot_entry = ttk.Entry(widgets_frame)
lot_entry.insert(0, "Lote")
lot_entry.bind("<FocusIn>", lambda e: lot_entry.delete(0, "end") if lot_entry.get() == "Lote" else None)
lot_entry.grid(row=0, column=0, padx=5, pady=(0,5), sticky="ew")

# Part ID
part_id_entry = ttk.Entry(widgets_frame)
part_id_entry.insert(0, "Part ID")
part_id_entry.bind("<FocusIn>", lambda e: part_id_entry.delete(0, "end") if part_id_entry.get() == "Part ID" else None)
part_id_entry.grid(row=1, column=0, padx=5, pady=(0,5), sticky="ew")

# Equipo
equipo_combo = ttk.Combobox(widgets_frame, values=combo_list_Equipo)
equipo_combo.set("Equipo")
equipo_combo.grid(row=2, column=0, padx=5, pady=(0,5), sticky="ew")

# CT y CP
ct_entry = ttk.Entry(widgets_frame, justify="center")
ct_entry.insert(0, "CT")
ct_entry.bind("<FocusIn>", lambda e: ct_entry.delete(0, "end") if ct_entry.get() == "CT" else None)
ct_entry.grid(row=3, column=0, sticky="ew", padx=(0,2), pady=2)

cp_entry = ttk.Entry(widgets_frame, justify="center")
cp_entry.insert(0, "CP")
cp_entry.bind("<FocusIn>", lambda e: cp_entry.delete(0, "end") if cp_entry.get() == "CP" else None)
cp_entry.grid(row=3, column=1, sticky="ew", padx=(2,0), pady=2)

# Comentario
ttk.Label(widgets_frame, text="Comentario:").grid(row=4, column=0, padx=5, pady=5, sticky="w")
cometario_entry = tk.Text(widgets_frame, height=6, width=30, wrap="word")
cometario_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=2)

# PCB
pcb_entry = ttk.Entry(widgets_frame, width=15, justify="center")
pcb_entry.insert(0, "PCB")
pcb_entry.bind("<FocusIn>", lambda e: pcb_entry.delete(0, "end") if pcb_entry.get() == "PCB" else None)
pcb_entry.grid(row=6, column=0, sticky="ew", padx=(0,2), pady=2)

# Fecha y hora
ttk.Label(widgets_frame, text="Fecha:").grid(row=7, column=0, padx=5, pady=4, sticky="w")
calendario = DateEntry(widgets_frame, date_pattern="mm/dd/yyyy")
calendario.grid(row=8, column=0, padx=5, pady=5, sticky="ew")

ttk.Label(widgets_frame, text="Hora (24h):").grid(row=7, column=1, padx=5, pady=4, sticky="w")
hora_entry = ttk.Entry(widgets_frame)
hora_entry.grid(row=8, column=1, padx=5, pady=5, sticky="ew")

# Turno y RCA
turno_combo = ttk.Combobox(widgets_frame, values=combo_list_Turno)
turno_combo.set("Turno")
turno_combo.grid(row=9, column=0, padx=5, pady=5, sticky="ew")

rca_combo = ttk.Combobox(widgets_frame, values=combo_list_RCA)
rca_combo.set("RCA")
rca_combo.grid(row=9, column=1, padx=5, pady=5, sticky="ew")

# Estatus
estatus_combo = ttk.Combobox(widgets_frame, values=combo_list_Estatus)
estatus_combo.set("Estatus")
estatus_combo.grid(row=10, column=0, columnspan=2, sticky="ew", pady=2)

separator = ttk.Separator(widgets_frame)
separator.grid(row=11, column=0, columnspan=2, sticky="ew", pady=10)

# Botones
guardar_button = ttk.Button((button_frame := ttk.Frame(widgets_frame)), text="Guardar Cambios", command=guardar_cambios, style="Accent.TButton")
guardar_button.pack_forget()  # Ocultar inicialmente

button_frame = ttk.Frame(widgets_frame)
button_frame.grid(row=12, column=0, columnspan=2, pady=5, sticky="nsew")

insert_button = ttk.Button(button_frame, text="Insertar Evento", command=guardar_datos, style="Accent.TButton")
insert_button.pack(side="left", padx=2, expand=True)

clear_button = ttk.Button(button_frame, text="Limpiar", command=limpiar_campos)
clear_button.pack(side="left", padx=2, expand=True)

search_button = ttk.Button(button_frame, text="Buscar", command=buscar_por_lote)
search_button.pack(side="left", padx=2, expand=True)

# Botón para generar gráfica de turnos
graph_button = ttk.Button(button_frame, text="Gráfica Turnos", command=generar_grafica_turnos, style="Accent.TButton")
graph_button.pack(side="left", padx=2, expand=True)

# Botón para generar gráfica de equipos
graph_equipos_button = ttk.Button(button_frame, text="Gráfica Equipos", command=generar_grafica_equipos, style="Accent.TButton")
graph_equipos_button.pack(side="left", padx=2, expand=True)

# Frame para el Treeview
treeFrame = ttk.LabelFrame(frame, text="Registros")
treeFrame.grid(row=0, column=1, padx=10, pady=10, sticky="nsew")  

# Scrollbar para el Treeview
treescroll = ttk.Scrollbar(treeFrame)
treescroll.pack(side="right", fill="y")

# Configurar el Treeview
cols = ("Lote", "Part ID", "Equipo", "Comentario", "Estatus")
treeview = ttk.Treeview(
    treeFrame,
    show="headings",
    yscrollcommand=treescroll.set,
    columns=cols,
    height=12,
    selectmode="browse"
)

# Configurar columnas
treeview.column("Lote", anchor="w", width=80)
treeview.column("Part ID", anchor="w", width=80)  
treeview.column("Equipo", anchor="w", width=80)
treeview.column("Comentario", anchor="w", width=200)
treeview.column("Estatus", anchor="w", width=200)

# Configurar encabezados
for col in cols:
    treeview.heading(col, text=col)

treeview.pack(fill="both", expand=True)
treescroll.config(command=treeview.yview)

# Botón para copiar selección
copy_button = ttk.Button(
    treeFrame,
    text="Copiar Selección",
    command=copiar_seleccion,
    style="Accent.TButton"
)
copy_button.pack(pady=5)

#Boton para modificar evento
edit_button = ttk.Button(treeFrame, text="Modificar", command=modificar_evento, style="Accent.TButton")
edit_button.pack(pady=5)


# Cargar datos iniciales
actualizar_treeview()

# Centrar la ventana
root.update()
root.minsize(1300, 650)
x = (root.winfo_screenwidth() // 2) - (root.winfo_width() // 2)
y = (root.winfo_screenheight() // 2) - (root.winfo_height() // 2)
root.geometry(f"+{x}+{y}")

root.mainloop()

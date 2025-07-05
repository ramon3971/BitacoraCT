Bitácora CT Hold - Documentación
Descripción
La aplicación Bitácora CT Hold es una herramienta diseñada para registrar y gestionar eventos relacionados con el proceso de moldeo en equipos Towa. Permite el registro, modificación y análisis de eventos, así como la generación de reportes en diferentes formatos.

Características principales
Registro de eventos con múltiples campos de información

Visualización de registros en formato tabla

Modificación de registros existentes

Generación de gráficas estadísticas por turnos y equipos

Exportación de datos a formato CSV

Generación de reportes RCA en PowerPoint

Interfaz intuitiva con tema oscuro

Requisitos del sistema
Python 3.7 o superior

Bibliotecas requeridas (ver requirements.txt)

Instalación
Clonar el repositorio o descargar los archivos

Instalar las dependencias:

text
pip install -r requirements.txt
Ejecutar la aplicación:

text
python BitacoraCT.py
Uso básico
Insertar un nuevo evento:

Complete los campos obligatorios (Lote, Part ID, Comentario)

Haga clic en "Insertar Evento"

Modificar un evento existente:

Seleccione un registro en la tabla

Haga clic en "Modificar"

Realice los cambios necesarios

Haga clic en "Guardar Cambios"

Generar gráficas:

Utilice los botones "Gráfica Turnos" o "Gráfica Equipos"

Seleccione el mes y año deseado

Visualice las estadísticas

Exportar datos:

CSV: Haga clic en "Generar CSV" para exportar todos los registros

Reporte RCA: Seleccione un registro y haga clic en "Generar Reporte RCA"

Estructura del código
El proyecto consta de dos clases principales:

BitacoraCTApp: Clase principal que maneja la interfaz y funcionalidad de la bitácora

DefectReportApp: Clase para generar reportes RCA en formato PowerPoint

Configuración
La ruta de la base de datos Excel puede modificarse en la variable self.ruta_base_datos

Los temas de la interfaz se encuentran en los archivos .tcl incluidos

Versión
Actual: 1.3.2

Notas importantes
La aplicación requiere acceso al archivo Excel en la ruta especificada

Algunas funcionalidades como los reportes RCA requieren la plantilla PowerPoint (ppt_test.pptx)

Se recomienda verificar los permisos de escritura en la ubicación del archivo Excel

Soporte
Para problemas o sugerencias, contactar al desarrollador responsable.

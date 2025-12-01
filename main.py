import flet as ft
import os
import asyncio
import pandas as pd
import time
from dcf_calculator import DCFModel 

# --- 1. Variables Globales y Configuración de Ruta ---

dcf_model = None
file_path = None
flet_page = None # Variable global para almacenar la referencia a la página

# NOTA: En modo escritorio, la carpeta 'uploads' ya no es estrictamente necesaria
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) 
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads") 

if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

# --- 2. Contenedores y Variables de UI (Globales) ---
valuation_kpis_container = ft.Column(scroll=ft.ScrollMode.ADAPTIVE)
proyection_table_container = ft.Container(ft.Text("Cargue el archivo para ver la tabla."))
fcf_chart_container = ft.Container(ft.Text("Gráfico FCF"))
status_text = ft.Text("Presione 'Cargar Excel' para iniciar.", color=ft.Colors.BLACK)

wacc_input = None
g_perpetuidad_input = None
file_picker = None 

# --- 3. Funciones de Presentación y Helper ---

def format_currency(value):
    """Función auxiliar para formatear valores a moneda."""
    return f"${value:,.0f}"

def create_results_table(df: pd.DataFrame):
    """Crea una tabla Flet a partir del DataFrame de proyecciones."""
    if df.empty:
        return ft.Text("No hay datos de proyección disponibles.")

    columns = ['Ingresos Totales', 'FCF', 'Valor Presente FCF']
    
    header_cells = [ft.DataColumn(ft.Text(col, weight=ft.FontWeight.BOLD)) for col in columns]
    
    rows = []
    for index, row in df.iterrows():
        rows.append(
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(format_currency(row[col])) if col != 'Valor Presente FCF' else ft.DataCell(ft.Text(format_currency(row[col]), weight=ft.FontWeight.BOLD)))
                    for col in columns
                ],
                color={'': ft.Colors.BLUE_GREY_100},
            )
        )

    return ft.DataTable(
        columns=header_cells,
        rows=rows,
        heading_row_color=ft.Colors.CYAN_50,
        border=ft.border.all(1, ft.Colors.BLACK12),
        width=float('inf')
    )

def update_results():
    """Actualiza todos los elementos de la interfaz con los resultados del modelo."""
    summary = dcf_model.get_valuation_summary()
    proyeccion_df = dcf_model.get_detailed_proyection()

    # 1. Actualizar KPIs (Tarjetas de resumen)
    kpi_list = [
        ft.Container(
            content=ft.Column([
                ft.Text("💰 VALOR DE LA EMPRESA (VE)", size=16, weight=ft.FontWeight.BOLD),
                ft.Text(format_currency(summary['Valor_Empresa_VE']), size=32, weight=ft.FontWeight.BOLD, color=ft.Colors.GREEN_800),
            ]),
            padding=15, border_radius=10, bgcolor=ft.Colors.GREEN_50, expand=True
        ),
        ft.Container(
            content=ft.Column([
                ft.Text("VP Valor Terminal", size=14),
                ft.Text(format_currency(summary['VP_Valor_Terminal']), size=24, weight=ft.FontWeight.SEMIBOLD, color=ft.Colors.BLUE_GREY_700),
            ]),
            padding=15, border_radius=10, bgcolor=ft.Colors.BLUE_GREY_50, expand=True
        ),
        ft.Container(
            content=ft.Column([
                ft.Text(f"WACC: {summary['WACC'] * 100:.1f}%", size=14),
                ft.Text(f"g: {summary['G_Perpetuidad'] * 100:.1f}%", size=14),
            ]),
            padding=15, border_radius=10, bgcolor=ft.Colors.CYAN_50, expand=True
        )
    ]
    valuation_kpis_container.controls = [
        ft.ResponsiveRow([
            ft.Container(kpi_list[0], col=5),
            ft.Container(kpi_list[1], col=4),
            ft.Container(kpi_list[2], col=3),
        ])
    ]
    
    proyection_table_container.content = create_results_table(proyeccion_df)
    
    fcf_chart_container.content = ft.Container(
        ft.Column([
            ft.Text("📊 Gráfico Dinámico: FCF por Año", size=18, weight=ft.FontWeight.BOLD),
            ft.Text("El gráfico se actualizaría aquí con Plotly/librería compatible."), 
        ]),
        padding=10, border=ft.border.all(1, ft.Colors.BLACK12)
    )
    
    flet_page.update()

def update_sensitivity_inputs():
    """Rellena los campos de sensibilidad con los valores cargados del Excel."""
    if dcf_model and dcf_model.data:
        base = dcf_model.data['base']
        wacc_input.value = f"{base['WACC'] * 100:.1f}"
        g_perpetuidad_input.value = f"{base['G_Perpetuidad'] * 100:.1f}"

# --- 4. Funciones de Lógica Síncrona ---

def run_recalculation(e):
    """Recalcula el modelo al modificar WACC o g."""
    global dcf_model
    
    if not dcf_model:
        status_text.value = "⚠️ Primero debe cargar un archivo Excel."
        status_text.color = ft.Colors.ORANGE_700
        flet_page.update()
        return

    try:
        recalculate_dcf_model_sync(float(wacc_input.value), float(g_perpetuidad_input.value))
        status_text.value = "✨ Modelo recalculado..."
        status_text.color = ft.Colors.BLUE_700
    except ValueError:
        status_text.value = "❌ Error: WACC y g deben ser valores numéricos."
        status_text.color = ft.Colors.RED_700
        
    flet_page.update()

def recalculate_dcf_model_sync(new_wacc_pct, new_g_pct):
    """Función síncrona que aplica las nuevas variables y recalcula."""
    dcf_model.data['base']['WACC'] = new_wacc_pct / 100
    dcf_model.data['base']['G_Perpetuidad'] = new_g_pct / 100
    
    dcf_model.run_model()
    
    update_results()

def execute_model_after_upload_sync(file_path_absolute, file_name):
    """
    Función síncrona que inicializa y ejecuta el modelo DCF.
    Ejecutada en modo escritorio con la ruta absoluta directa.
    """
    global dcf_model
    
    # 1. Inicializar y Ejecutar el Modelo DCF
    dcf_model = DCFModel(file_path_absolute)
    
    if dcf_model.data and dcf_model.run_model():
        status_text.value = f"✅ Archivo '{file_name}' cargado y modelo ejecutado."
        status_text.color = ft.Colors.GREEN_700
        update_sensitivity_inputs()
        update_results() 
    else:
        status_text.value = f"❌ Error al leer el Excel. Revise la estructura de datos en '{file_name}'."
        status_text.color = ft.Colors.RED_700
    
    flet_page.update()

# --- 5. Handler Principal (Disparador de Archivo) ---

# Función para manejar el evento on_click del botón
def button_click_handler(e):
    """Dispara el diálogo de selección de archivos."""
    print("--- DEBUG: Botón clickeado, llamando a pick_files() ---") # <--- AÑADIR ESTO
    if file_picker:
        print("otro")
        file_picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx"]
        )

# Función para manejar el resultado del FilePicker
def handle_file_pick_consolidated(e: ft.FilePickerResultEvent):
    """Maneja la selección del archivo y ejecuta la lógica de cálculo SÍNCRONO."""
    global dcf_model
    
    if e.files:
        # En modo ESCRITORIO, e.files[0].path contiene la RUTA ABSOLUTA local.
        absolute_path = e.files[0].path
        file_name = e.files[0].name
        
        status_text.value = "⏳ Cargando y ejecutando modelo..."
        flet_page.update() 
        
        # LLAMADA DIRECTA SÍNCRONA: Ejecutamos el cálculo inmediatamente.
        execute_model_after_upload_sync(
            absolute_path, 
            file_name
        )
        
    else:
        status_text.value = "📂 Carga de archivo cancelada."
        status_text.color = ft.Colors.BLACK
        flet_page.update()

# --- 6. Función Principal de Flet ---

def main(page: ft.Page):
    global wacc_input, g_perpetuidad_input, file_picker, flet_page
    
    # 1. Configuración y Asignación Global de Page
    flet_page = page 
    page.title = "Simulador de Valuación DCF"
    page.theme_mode = ft.ThemeMode.LIGHT
    page.padding = 20
    page.vertical_alignment = ft.CrossAxisAlignment.START
    page.scroll = ft.ScrollMode.ADAPTIVE

    # 2. Inicialización Prioritaria del FilePicker
    file_picker = ft.FilePicker(on_result=handle_file_pick_consolidated)
    page.overlay.append(file_picker) # Añadido al overlay antes de construir el layout

    # 3. Inicialización de Otros Componentes
    wacc_input = ft.TextField(label="WACC (%)", width=150, on_change=run_recalculation)
    g_perpetuidad_input = ft.TextField(label="g Perpetuidad (%)", width=150, on_change=run_recalculation)
    
    # 4. Layout Principal
    sensitivity_controls = ft.Row(
        controls=[
            ft.Icon(ft.Icons.TUNE, color=ft.Colors.BLUE_800), 
            ft.Text("Inputs de Sensibilidad:", weight=ft.FontWeight.BOLD),
            wacc_input,
            g_perpetuidad_input,
        ],
        alignment=ft.MainAxisAlignment.START,
        spacing=20
    )

    app_layout = ft.Column(
        controls=[
            ft.ResponsiveRow([
                ft.Container(
                    content=ft.Text("VALUACIÓN DCF", size=30, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_800),
                    col=6
                ),
                ft.Container(
                    content=ft.Row([
                        ft.ElevatedButton(
                            text="Cargar Plantilla Excel",
                            icon=ft.Icons.UPLOAD_FILE, 
                            on_click=button_click_handler, # <-- Usa el handler directo
                        ),
                        status_text
                    ], alignment=ft.MainAxisAlignment.END),
                    col=6
                )
            ], vertical_alignment=ft.CrossAxisAlignment.CENTER),
            
            ft.Divider(height=2, color=ft.Colors.BLUE_GREY_200),
            
            sensitivity_controls,
            
            ft.Divider(height=2, color=ft.Colors.BLUE_GREY_200),
            
            valuation_kpis_container,
            
            ft.ResponsiveRow([
                ft.Container(fcf_chart_container, col=5, padding=10, border_radius=10, bgcolor=ft.Colors.WHITE70),
                ft.Container(proyection_table_container, col=7, padding=10, border_radius=10, bgcolor=ft.Colors.WHITE70),
            ], vertical_alignment=ft.CrossAxisAlignment.START, spacing=15),
        ],
        horizontal_alignment=ft.CrossAxisAlignment.STRETCH
    )

    page.add(app_layout)

if __name__ == '__main__':
    # EJECUCIÓN EN MODO ESCRITORIO
    ft.app(target=main)
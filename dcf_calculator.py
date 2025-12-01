import pandas as pd
import numpy as np
import io
import os
from typing import Union

# --- Funciones Auxiliares ---

def clean_and_convert(value, is_percentage=False):
    """
    Limpia cadenas, maneja valores NaN (celdas vacías) y convierte a float.
    Aplica la conversión a decimal (divide por 100) si es un porcentaje.
    """
    if pd.isna(value) or value == '':
        return 0.0
    
    # Intenta convertir a string y limpiar si no es float/int directo
    if not isinstance(value, (int, float)):
        try:
            # Limpia formatos de miles y decimales
            value_str = str(value).replace('$', '').replace(',', '').strip()
            # Si termina en %, lo quita y lo divide por 100 después
            if value_str.endswith('%'):
                is_percentage = True
                value_str = value_str.replace('%', '')
            
            numeric_value = float(value_str)
        except ValueError:
            return 0.0
    else:
        numeric_value = float(value)

    if is_percentage and numeric_value > 1:
        return numeric_value / 100
    
    return numeric_value


class DCFModel:
    """
    Clase que encapsula la lógica de lectura de datos, proyección financiera
    y cálculo de la valuación por Flujos de Caja Libre Descontados (FCF).
    """
    def __init__(self, file_source: Union[str, io.BytesIO]):
        # 'file_source' puede ser una ruta (string) o un objeto BytesIO (Streamlit)
        self.file_source = file_source
        self.proyeccion = pd.DataFrame()
        self.valoracion = {}
        self.years = 5 # Horizonte de proyección explícita (2026 a 2030)
        self.start_year = 2026
        self.base_year = 2025
        self.data = self._load_data()

    def _load_data(self):
        """
        Carga datos del Excel desde un buffer de memoria (Streamlit) o una ruta,
        utilizando pd.ExcelFile.
        """
        try:
            # 1. Determinar la fuente de datos
            if isinstance(self.file_source, str) and not os.path.exists(self.file_source):
                return None
            
            # Crea el objeto ExcelFile, el cual funciona tanto con rutas como con BytesIO
            xl = pd.ExcelFile(self.file_source)

            # --- A. Hipótesis_Base (Hoja sensible a etiquetas) ---
            df_base = xl.parse(sheet_name='Hipotesis_Base', header=None, index_col=0) # skiprows=2
            
            # --- CORRECCIÓN: Normalizar el índice para búsqueda robusta ---
            df_base.index = df_base.index.fillna('')
            df_base.index = df_base.index.astype(str).str.strip().str.lower().str.replace(' ', '_')

            # Función auxiliar para obtener un valor por etiqueta limpia de fila
            def get_base_value(clean_label, is_percentage=False):
                # Usamos .loc y la columna con índice 1 (la segunda columna del Excel)
                value = df_base.loc[clean_label].iloc[0] 
                return clean_and_convert(value, is_percentage=is_percentage)

            # Mapeo de variables (todas en minúsculas y sin espacios/tildes en el código)
            base_vars = {
                'Ingresos_Totales_2025': get_base_value('ingresos_totales_2025'),
                'Gastos_Fijos_Operativos_2025': get_base_value('gastos_fijos_operativos_2025'),
                'Dias_CxC': get_base_value('dias_cxc'),
                'Dias_Inv': get_base_value('dias_inv'),
                'Dias_CxP': get_base_value('dias_cxp'),
                'WACC': get_base_value('wacc', is_percentage=True),
                'G_Perpetuidad': get_base_value('g-tasa_de_crecimiento_a_perpetuidad', is_percentage=True),
                'Tasa_ISR': get_base_value('tasa_isr', is_percentage=True),
                'Dep_Pct_Base': get_base_value('dep_pct_base', is_percentage=True),
                'CapEx_Pct_Base': get_base_value('capex_pct_base', is_percentage=True),
            }
            
            # --- B. Proyecciones_Detalladas (Usada por índice) ---
            proyecciones_detalladas = xl.parse(sheet_name='Proyecciones_Detalladas', header=0, index_col=None)
            proyecciones_detalladas = proyecciones_detalladas.set_index(proyecciones_detalladas.columns[0])

            # --- C. Impacto_Proyectos (Hoja con estructura fija) ---
            impacto_proyectos = xl.parse(sheet_name='Impacto_Proyectos', header=None, skiprows=2)

            # --- CORRECCIÓN DE AMBITO: Definición de project_data ---
            # Las proyecciones de CapEx, Ingresos, etc., inician en la columna 3 (Año 2026)
            project_data = {
                'capex_inversion': [clean_and_convert(val) for val in impacto_proyectos.iloc[0].values[3:8]], 
                'ingresos_adicionales': [clean_and_convert(val) for val in impacto_proyectos.iloc[4].values[3:8]],
                'gastos_ahorros_operativos': [clean_and_convert(val) for val in impacto_proyectos.iloc[8].values[3:8]],
                'depreciacion_adicional': [clean_and_convert(val) for val in impacto_proyectos.iloc[12].values[3:8]],
            }    
            
            # --- Retorno del diccionario de datos ---
            return {'base': base_vars, 'proyecciones': proyecciones_detalladas, 'proyectos': project_data}

        except KeyError as ke:
             # Captura errores específicos de nombre de hoja o etiqueta
            error_message = f"Error: No se encontró la etiqueta o hoja necesaria. Verifique que exista la hoja 'Hipotesis_Base' y la etiqueta '{ke}'."
            print(error_message)
            return None
        except Exception as e:
            print(f"Error al cargar datos del Excel desde memoria: {e}")
            return None

    def _calculate_projection(self):
        """
        Realiza la proyección de 5 años (2026 - 2030) del FCF, 
        usando acceso por índice numérico fijo.
        """
        data = self.data
        if not data: return False
        
        # Parámetros Base
        base = data['base']
        proyecciones_raw = data['proyecciones']
        proyectos = data['proyectos']
        
        proj_years_idx = range(1, self.years + 1) # t = 1, 2, 3, 4, 5
        years_labels = range(self.start_year, self.start_year + self.years) # 2026 a 2030
        df_proj = pd.DataFrame(index=years_labels)
        
        # Valores base fijos 
        base_gastos_fijos = base['Gastos_Fijos_Operativos_2025']
        
        # --- CORRECCIÓN DE ÍNDICES DE COLUMNA ---
        base_year_col = 0       # 2025 está en la primera columna de datos (índice 0)
        
        # Esta función obtiene el ingreso del año anterior (2025 para t=1)
        def get_prev_ingreso(ingreso_key, base_ingreso_value):
            return df_proj.loc[prev_year_str, ingreso_key] if t > 1 else base_ingreso_value
        
        # Iteración: t va de 1 a 0
        for t in proj_years_idx:
            year_str = self.start_year + t - 1 # El año actual de proyección
            prev_year_str = self.start_year + t - 2 # El año anterior
            
            current_year_col = t # CORREGIDO: Columna 1, 2, 3, 4, 5 (para 2026 a 2030)
            
            total_ingresos_base = 0
            total_costo_venta = 0
            
            # --- Proyecciones Detalladas (Acceso por Índice Fijo) ---
            
            # 1. Binomio 1 (Fila 2, 3, 4)
            ingreso_1_base_2025 = clean_and_convert(proyecciones_raw.iloc[1, base_year_col]) # Fila 1 para Ingresos Base
            g_ingreso_1 = clean_and_convert(proyecciones_raw.iloc[2, current_year_col], is_percentage=True)
            cv_pct_1 = clean_and_convert(proyecciones_raw.iloc[3, current_year_col], is_percentage=True)
            
            ingreso_1_prev = get_prev_ingreso('Ingresos_B1', ingreso_1_base_2025)
            ingreso_1_actual = ingreso_1_prev * (1 + g_ingreso_1)
            df_proj.loc[year_str, 'Ingresos_B1'] = ingreso_1_actual
            total_costo_venta += ingreso_1_actual * cv_pct_1

            # 2. Binomio 2 (Fila 5, 6, 7)
            ingreso_2_base_2025 = clean_and_convert(proyecciones_raw.iloc[5, base_year_col]) 
            g_ingreso_2 = clean_and_convert(proyecciones_raw.iloc[6, current_year_col], is_percentage=True)
            cv_pct_2 = clean_and_convert(proyecciones_raw.iloc[7, current_year_col], is_percentage=True)
            
            ingreso_2_prev = get_prev_ingreso('Ingresos_B2', ingreso_2_base_2025)
            ingreso_2_actual = ingreso_2_prev * (1 + g_ingreso_2)
            df_proj.loc[year_str, 'Ingresos_B2'] = ingreso_2_actual
            total_costo_venta += ingreso_2_actual * cv_pct_2
            
            # 3. Componente General (Fila 9, 10, 11)
            ingreso_G_base_2025 = clean_and_convert(proyecciones_raw.iloc[9, base_year_col])
            g_ingreso_G = clean_and_convert(proyecciones_raw.iloc[10, current_year_col], is_percentage=True)
            cv_pct_G = clean_and_convert(proyecciones_raw.iloc[11, current_year_col], is_percentage=True)
            
            ingreso_G_prev = get_prev_ingreso('Ingresos_G', ingreso_G_base_2025)
            ingreso_G_actual = ingreso_G_prev * (1 + g_ingreso_G)
            df_proj.loc[year_str, 'Ingresos_G'] = ingreso_G_actual
            total_costo_venta += ingreso_G_actual * cv_pct_G

            # Total Ingresos Base (sin proyectos)
            total_ingresos_base = ingreso_1_actual + ingreso_2_actual + ingreso_G_actual
            
            # --- Aplicar Impacto de Proyectos (Flujos Adicionales) ---
            ingresos_adicionales = proyectos['ingresos_adicionales'][t-1]
            gastos_ahorros_operativos = proyectos['gastos_ahorros_operativos'][t-1]
            depreciacion_adicional = proyectos['depreciacion_adicional'][t-1]
            capex_proyectos = proyectos['capex_inversion'][t-1]

            # --- Cálculo de EBIT (Utilidad Operativa) ---
            
            g_fijos_inflacion = 0.03 
            gastos_fijos_actual = base_gastos_fijos * (1 + g_fijos_inflacion)**t
            
            # Gastos Variables Operativos (Gtos. Adm. & Vta.) - Fila 13
            gastos_variables_pct = clean_and_convert(proyecciones_raw.iloc[13, current_year_col], is_percentage=True)
            gastos_variables_actual = total_ingresos_base * gastos_variables_pct
            
            # TOTALES PROYECTADOS
            df_proj.loc[year_str, 'Ingresos Totales'] = total_ingresos_base + ingresos_adicionales
            
            # EBIT
            df_proj.loc[year_str, 'EBIT'] = df_proj.loc[year_str, 'Ingresos Totales'] - total_costo_venta - \
                                            gastos_fijos_actual - gastos_variables_actual + gastos_ahorros_operativos
            
            # ... (El resto del cálculo del FCF, NOPAT, CKT, etc. continúa sin cambios) ...
            # (Se asume que la parte de CKT, CapEx, Depreciación, NOPAT y FCF está correcta.)

            # Depreciación (Base + Proyectos)
            dep_amort_base = total_ingresos_base * base['Dep_Pct_Base']
            df_proj.loc[year_str, 'Depreciacion'] = dep_amort_base + depreciacion_adicional

            # NOPAT (Net Operating Profit After Tax)
            df_proj.loc[year_str, 'NOPAT'] = df_proj.loc[year_str, 'EBIT'] * (1 - base['Tasa_ISR'])
            
            # CKT (Se asume la simplificación: Delta CKT es la diferencia del CKT actual y el anterior)
            # Placeholder del CKT Base 2025 para t=1.
            ck_current = (base['Dias_CxC'] / 365) * df_proj.loc[year_str, 'Ingresos Totales'] + \
                        (base['Dias_Inv'] / 365) * total_costo_venta - \
                        (base['Dias_CxP'] / 365) * total_costo_venta
            
            ck_prev = df_proj.loc[prev_year_str, 'Capital de Trabajo'] if t > 1 else 0
            
            df_proj.loc[year_str, 'Capital de Trabajo'] = ck_current
            df_proj.loc[year_str, 'Delta CTN'] = ck_current - ck_prev

            # CapEx (Base + Proyectos)
            capex_base = total_ingresos_base * base['CapEx_Pct_Base']
            df_proj.loc[year_str, 'CapEx Total'] = capex_base + capex_proyectos 
            
            # FCF
            df_proj.loc[year_str, 'FCF'] = df_proj.loc[year_str, 'NOPAT'] + \
                                            df_proj.loc[year_str, 'Depreciacion'] - \
                                            df_proj.loc[year_str, 'Delta CTN'] - \
                                            df_proj.loc[year_str, 'CapEx Total']
        
        # --- 4. Valor Presente (Calculado después de que todos los años tienen FCF) ---
        wacc = base['WACC']
        proj_years_idx = range(1, self.years + 1) # Aseguramos que es de 1 a 5
        
        # Crear los factores de descuento para cada año (t=1, 2, 3, 4, 5)
        discount_factors = [(1 / (1 + wacc)**i) for i in proj_years_idx]
        
        # Asignar los factores al DataFrame (el DataFrame ya tiene 5 filas, 2026-2030)
        df_proj.loc[:, 'VP Factor'] = discount_factors
        
        # Calcular el VP del FCF para cada año
        df_proj.loc[:, 'Valor Presente FCF'] = df_proj['FCF'] * df_proj['VP Factor']
        
        self.proyeccion = df_proj
        return True

    def _calculate_valuation(self):
        """Calcula el Valor Terminal y el Valor Presente del VE."""
        if self.proyeccion.empty: return False

        base = self.data['base']
        wacc = base['WACC']
        g = base['G_Perpetuidad']
        
        # CORRECCIÓN DE ÍNDICE: El último año proyectado es 2030 (2026 + 5 - 1)
        last_year_label = self.start_year + self.years - 1 
        
        # Último NOPAT proyectado y FCF (Year N = 2030)
        nopat_n = self.proyeccion.loc[last_year_label, 'NOPAT']
        fcf_n = self.proyeccion.loc[last_year_label, 'FCF']
        vp_factor_n = self.proyeccion.loc[last_year_label, 'VP Factor'] # Ya existe si el paso 1 fue correcto
        
        # FCF perpetuo (AÑO N+1)
        fcf_perpetuo = fcf_n * (1 + g)
        
        # Valor Terminal (VT) - Fórmula de Gordon
        valor_terminal = fcf_perpetuo / (wacc - g)
        
        # VP del Valor Terminal (VP_VT)
        vp_vt = valor_terminal * vp_factor_n # Usa el VP Factor del año 2030
        
        # VP de los FCF Proyectados
        vp_fcf_proyectados = self.proyeccion['Valor Presente FCF'].sum()
        
        # Valor de la Empresa (VE)
        valor_empresa_ve = vp_fcf_proyectados + vp_vt
        
        self.valoracion = {
            'VE_FCF_Proyectados': vp_fcf_proyectados,
            'Valor_Terminal': valor_terminal,
            'VP_Valor_Terminal': vp_vt,
            'Valor_Empresa_VE': valor_empresa_ve,
            'WACC': wacc,
            'G_Perpetuidad': g,
            'Tasa_ISR': base['Tasa_ISR'],
        }
        return True

    def run_model(self):
        """Ejecuta toda la secuencia de cálculo."""
        if not self.data: return False
        
        success = self._calculate_projection()
        if success:
            return self._calculate_valuation()
        return False
        
    def get_valuation_summary(self):
        """Retorna el diccionario de resumen de valuación."""
        return self.valoracion

    def get_detailed_proyection(self):
        """Retorna el DataFrame de proyecciones, solo años proyectados."""
        return self.proyeccion
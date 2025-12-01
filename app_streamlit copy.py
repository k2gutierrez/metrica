import streamlit as st
import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go # Necesario para combinar líneas y barras
from dcf_calculator import DCFModel, clean_and_convert 
from typing import Union

# --- Configuración de la Página ---
st.set_page_config(layout="wide", page_title="Simulador de Valuación DCF (Plotly)")

# --- Funciones Auxiliares ---

def format_currency_st(value):
    """Formatea valores a moneda."""
    if pd.isna(value) or value is None:
        return "$0"
    return f"${value:,.0f}"

def to_excel_consolidated(summary_df: pd.DataFrame, proj_df: pd.DataFrame, proy_analysis_df: pd.DataFrame) -> bytes:
    """Convierte múltiples DataFrames a un archivo XLSX con varias hojas."""
    output = io.BytesIO()
    
    # Crea un objeto ExcelWriter y lo dirige al buffer de BytesIO
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Hoja 1: Sumario de Valuación (Múltiplos)
        summary_df.to_excel(writer, sheet_name='Sumario_Valuacion', index=True)
        
        # Hoja 2: Proyecciones Detalladas
        proj_df.to_excel(writer, sheet_name='Proyecciones_Detalladas', index=True)
        
        # Hoja 3: Análisis de Proyectos
        proy_analysis_df.to_excel(writer, sheet_name='Analisis_Proyectos', index=True)
        
    return output.getvalue()

@st.cache_data
def run_dcf_model(uploaded_file, wacc_pct, g_pct, isr_pct):
    """
    Inicializa y ejecuta el modelo DCF con los parámetros de sensibilidad.
    """
    
    try:
        file_buffer = io.BytesIO(uploaded_file.getvalue())
        model = DCFModel(file_buffer)

        if model.data is None:
            return None
            
        # 1. Aplicar Sensibilidad
        model.data['base']['WACC'] = wacc_pct / 100
        model.data['base']['G_Perpetuidad'] = g_pct / 100
        model.data['base']['Tasa_ISR'] = isr_pct / 100
        
        # 2. Ejecutar el cálculo
        model.run_model()
        
        # 3. Preparar datos de Proyectos para la pestaña (Impacto Incremental)
        df_proj = model.get_detailed_proyection()
        
        # Obtener los datos de impacto del proyecto (asumiendo que los flujos incrementales están en el modelo)
        # NOTA: Debes asegurar que tu DCFModel almacene la data del impacto incremental.
        # Aquí usamos los flujos de FCF y los flujos incrementales (Ingresos Adicionales - CapEx Adicional).
        
        # Creamos un DataFrame para el análisis de sensibilidad de proyectos
        proyectos_df = pd.DataFrame({
            'Año': df_proj.index,
            'FCF_Base': df_proj['FCF'] - model.data['proyectos']['ingresos_adicionales'], # FCF sin el impacto
            'Ingresos_Adicionales': model.data['proyectos']['ingresos_adicionales'],
            'CapEx_Proyectos': model.data['proyectos']['capex_inversion'],
            'FCF_Total': df_proj['FCF']
        }).set_index('Año')
        
        return model, proyectos_df
        
    except Exception as e:
        st.error(f"Error al procesar el archivo o la estructura de datos. Verifique la estructura del Excel. Error: {e}")
        return None, None

# --- Interfaz Principal (Streamlit) ---

st.title("💰 Simulador de Valuación por Flujos Descontados (DCF)")
st.caption("Gráficos interactivos y análisis de sensibilidad.")

# --- Barra Lateral (Sensibilidad) ---
st.sidebar.header("⚙️ Ajuste de Sensibilidad")

# Asumimos valores por defecto
default_wacc = 20.0
default_g = 3.0
default_isr = 30.0

wacc_input = st.sidebar.number_input("Costo Capital (WACC) (%)", value=default_wacc, step=0.1)
g_input = st.sidebar.number_input("Tasa Terminal (g) (%)", value=default_g, step=0.1)
isr_input = st.sidebar.number_input("Tasa Impositiva (ISR) (%)", value=default_isr, step=0.1)


# --- Carga de Archivo ---
uploaded_file = st.file_uploader(
    "1. Cargar Plantilla Excel (.xlsx)",
    type=['xlsx'],
    help="Carga tu archivo con las 3 hojas de proyecciones."
)

if uploaded_file is not None:
    
    # 2. Ejecución del Modelo con Inputs
    result = run_dcf_model(uploaded_file, wacc_input, g_input, isr_input)
    
    if result and result[0] and result[0].valoracion:
        model, proyectos_df = result
        summary = model.get_valuation_summary()
        proyeccion_df = model.get_detailed_proyection()

        # --- Pestañas de Análisis ---
        tab1, tab2, tab3 = st.tabs(["📊 Resumen Ejecutivo", "📋 Proyecciones Detalladas", "🧪 Análisis de Proyectos"])

        # ====================================================================
        # PESTAÑA 1: RESUMEN EJECUTIVO
        # ====================================================================
        with tab1:
            st.subheader("✅ Valuación y KPIs Principales")

            # Tarjetas de KPIs
            col_ve, col_vt, col_ratios = st.columns(3)
            
            with col_ve:
                st.metric(label="VALOR DE LA EMPRESA (VE)", 
                          value=format_currency_st(summary['Valor_Empresa_VE']), 
                          delta=f"WACC: {wacc_input:.1f}% / g: {g_input:.1f}%")

            with col_vt:
                st.metric(label="VP Valor Terminal", 
                          value=format_currency_st(summary['VP_Valor_Terminal']))
                
            with col_ratios:
                st.metric(label="VP FCF Proyectados", 
                          value=format_currency_st(summary['VE_FCF_Proyectados']))

            st.markdown("---")
            
            # --- Gráfico Combinado (Líneas y Barras) ---
            st.subheader("Gráfica Dinámica: Evolución de Flujos y Escala Operativa")
            
            # Crear figura de Plotly para el gráfico combinado
            fig = go.Figure()

            # Barras para Ingresos Totales (Escala)
            fig.add_trace(go.Bar(
                x=proyeccion_df.index.astype(str),
                y=proyeccion_df['Ingresos Totales'],
                name='Ingresos Totales',
                marker_color='rgb(158,202,225)',
                yaxis='y1' 
            ))

            # Línea para FCF (Flujo)
            fig.add_trace(go.Scatter(
                x=proyeccion_df.index.astype(str),
                y=proyeccion_df['FCF'],
                name='FCF (Flujo de Caja Libre)',
                mode='lines+markers',
                marker=dict(color='rgb(0,128,128)'),
                line=dict(width=3),
                yaxis='y2' 
            ))

            # Configuración de los Ejes Y
            fig.update_layout(
                title='Ingresos (Barras) vs. FCF (Línea)',
                yaxis=dict(
                    title='Ingresos (Escala Principal)',
                    tickfont=dict(color='rgb(158,202,225)'),
                    #tickfont=dict(color='rgb(158,202,225)'),
                    tickformat=',.0f'
                ),
                yaxis2=dict(
                    title='FCF (Eje Secundario)',
                    #titlefont=dict(color='rgb(0,128,128)'),
                    tickfont=dict(color='rgb(0,128,128)'),
                    overlaying='y',
                    side='right',
                    tickformat=',.0f'
                ),
                legend=dict(x=0.01, y=0.99),
                hovermode="x unified"
            )

            st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")

            # --- Tabla Final de Resumen (Múltiplos) ---
            st.subheader("📋 Sumario de Valuación y Múltiplos")
            
            last_year = model.start_year + model.years - 1 
            last_year_data = proyeccion_df.loc[last_year]

            ve = summary['Valor_Empresa_VE']
            ingresos_n = last_year_data['Ingresos Totales']
            ebitda_n = last_year_data['EBIT'] + last_year_data['Depreciacion']

            data_resumen = {
                'Métrica': [
                    'Valor de la Empresa (VE)',
                    'Valor Presente FCF Proyectado',
                    'Valor Terminal (VP)',
                    'Múltiplo VE / Ingresos (Año Terminal)',
                    'Múltiplo VE / EBITDA (Año Terminal)',
                    'Costo de Capital (WACC)',
                    'Tasa Impositiva (ISR)',
                ],
                'Valor': [
                    format_currency_st(ve),
                    format_currency_st(summary['VE_FCF_Proyectados']),
                    format_currency_st(summary['VP_Valor_Terminal']),
                    f"{ve / ingresos_n:.2f}", 
                    f"{ve / ebitda_n:.2f}",   
                    f"{summary['WACC'] * 100:.2f}%",
                    f"{summary['Tasa_ISR'] * 100:.2f}%",
                ]
            }

            df_final = pd.DataFrame(data_resumen)
            st.dataframe(df_final.set_index('Métrica'), use_container_width=True)

        # ====================================================================
        # PESTAÑA 2: PROYECCIONES DETALLADAS
        # ====================================================================
        with tab2:
            st.subheader("Proyecciones Detalladas por Año (Flujos y Balance)")
            
            # Aplicar formato de moneda al DataFrame de proyección final para la tabla
            # Mostramos todas las columnas relevantes del cálculo
            # NOPAT = Utilidad Operativa Neta Después de Impuestos
            # FCF Free cash flow
            cols_to_show = ['Ingresos Totales', 'EBIT', 'NOPAT', 'Depreciacion', 'Delta CTN', 'CapEx Total', 'FCF', 'Valor Presente FCF']
            df_display = proyeccion_df[cols_to_show].copy()
            
            for col in df_display.columns:
                df_display[col] = df_display[col].apply(format_currency_st)
                
            st.dataframe(df_display, use_container_width=True)

        # ====================================================================
        # PESTAÑA 3: ANÁLISIS DE PROYECTOS
        # ====================================================================
        with tab3:
            st.subheader("Análisis de Impacto de Proyectos (Incremental)")

            # Gráfico de Flujos Incrementales vs. FCF Total
            df_proy_chart = proyectos_df[['Ingresos_Adicionales', 'CapEx_Proyectos', 'FCF_Total']].copy()
            df_proy_chart['CapEx_Proyectos'] = df_proy_chart['CapEx_Proyectos'] * -1 # Mostrar CapEx como negativo
            
            fig_proj = go.Figure()

            # Barras para Ingresos Adicionales
            fig_proj.add_trace(go.Bar(
                x=df_proy_chart.index.astype(str),
                y=df_proy_chart['Ingresos_Adicionales'],
                name='Ingresos Adicionales',
                marker_color='green'
            ))

            # Barras para CapEx de Proyectos
            fig_proj.add_trace(go.Bar(
                x=df_proy_chart.index.astype(str),
                y=df_proy_chart['CapEx_Proyectos'],
                name='CapEx Proyectos (Inversión)',
                marker_color='red'
            ))

            # Línea para FCF Total (Referencia)
            fig_proj.add_trace(go.Scatter(
                x=df_proy_chart.index.astype(str),
                y=df_proy_chart['FCF_Total'],
                name='FCF Total Empresa',
                mode='lines+markers',
                line=dict(color='blue', width=3),
                yaxis='y2'
            ))
            
            fig_proj.update_layout(
                barmode='overlay',
                title='Impacto Financiero de Proyectos (Ingresos vs. Inversión)',
                yaxis=dict(title='Flujos Proyectos', tickformat=',.0f'),
                yaxis2=dict(title='FCF Total', overlaying='y', side='right', tickformat=',.0f'),
                hovermode="x unified"
            )
            
            st.plotly_chart(fig_proj, use_container_width=True)

            st.markdown("---")
            
            st.subheader("Tabla de Impacto Detallado")
            
            # Mostrar la tabla de impacto de proyectos
            proy_table = proyectos_df.copy()
            for col in proy_table.columns:
                proy_table[col] = proy_table[col].apply(format_currency_st)
            
            st.dataframe(proy_table, use_container_width=True)
            
    else:
        st.info("Cargue su archivo Excel en el apartado '1. Cargar Plantilla Excel' para iniciar la simulación.")
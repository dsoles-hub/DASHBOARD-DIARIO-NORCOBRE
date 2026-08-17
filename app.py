import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# ==========================================
# 1. CARGA Y PROCESAMIENTO DE DATOS
# ==========================================
def procesar_datos_produccion(df_planeamiento: pd.DataFrame, df_produccion: pd.DataFrame) -> pd.DataFrame:

    
    # Renombrar TMS-ACUMULADO a EJECUTADO en la hoja de producción
    df_prod = df_produccion[['FECHA', 'TMS-ACUMULADO']].copy()
    df_prod.rename(columns={'TMS-ACUMULADO': 'EJECUTADO'}, inplace=True)
    
    # Asegurar formato de fecha para la unión
    df_planeamiento['FECHA'] = pd.to_datetime(df_planeamiento['FECHA'])
    df_prod['FECHA'] = pd.to_datetime(df_prod['FECHA'])
    
    # Unir ambas tablas por FECHA
    df_merged = pd.merge(df_planeamiento, df_prod, on='FECHA', how='left')
    
    # Condición: Actualizar 'EJEC + PROYEC' con el valor de 'EJECUTADO' donde exista dato
    df_merged['EJEC_PROYEC_ACTUALIZADO'] = np.where(
        df_merged['EJECUTADO'].notnull() & (df_merged['EJECUTADO'] > 0),
        df_merged['EJECUTADO'],
        df_merged['EJEC + PROYEC']
    )
    
    # Formatear la fecha para visualización en gráficos (ej: 26 - Jul)
    df_merged['FECHA_STR'] = df_merged['FECHA'].dt.strftime('%d - %b')
    
    return df_merged

# ==========================================
# 2. GENERACIÓN DE GRÁFICOS Y CARDS
# ==========================================
def generar_panel_produccion(df_data: pd.DataFrame):
    
    # ------------------------------------------
    # A. CÁLCULOS KPI
    # ------------------------------------------
    # Máscara para días que ya tienen dato ejecutado
    mask_ejecutado = df_data['EJECUTADO'].notnull() & (df_data['EJECUTADO'] > 0)
    
    total_ejecutado = df_data.loc[mask_ejecutado, 'EJECUTADO'].sum()
    total_ejec_proy = df_data['EJEC_PROYEC_ACTUALIZADO'].sum()
    total_mensual_completo = df_data['MENSUAL'].sum()
    
    # Mensual acumulado únicamente hasta la fecha que se tiene dato ejecutado
    total_mensual_a_la_fecha = df_data.loc[mask_ejecutado, 'MENSUAL'].sum()
    
    # % Cumplimiento a la fecha
    pct_cumplimiento = (total_ejecutado / total_mensual_a_la_fecha * 100) if total_mensual_a_la_fecha > 0 else 0
    
    # Promedio ejecutado diario (solo de días transcurridos con dato)
    promedio_diario = df_data.loc[mask_ejecutado, 'EJECUTADO'].mean() if mask_ejecutado.sum() > 0 else 0

    # ------------------------------------------
    # B. GRÁFICO 1: PRODUCCIÓN DIARIA (TMS)
    # ------------------------------------------
    fig_diario = go.Figure()

    # Barras de EJECUTADO
    fig_diario.add_trace(go.Bar(
        x=df_data['FECHA_STR'],
        y=df_data['EJECUTADO'],
        name='EJECUTADO',
        marker_color='#3B82F6', # Azul
        text=df_data['EJECUTADO'].apply(lambda v: f"{int(v)}" if pd.notnull(v) and v > 0 else ""),
        textposition='inside',
        textfont=dict(color='white', size=11, family='Arial Black')
    ))

    # Línea SEMANAL
    fig_diario.add_trace(go.Scatter(
        x=df_data['FECHA_STR'],
        y=df_data['SEMANAL'],
        name='SEMANAL',
        mode='lines+markers+text',
        line=dict(color='#EAB308', width=3), # Amarillo / Naranja
        marker=dict(size=6),
        text=df_data['SEMANAL'].apply(lambda v: f"{int(v)}" if pd.notnull(v) else ""),
        textposition='top center',
        textfont=dict(color='#854D0E', size=10)
    ))

    # Línea MENSUAL
    fig_diario.add_trace(go.Scatter(
        x=df_data['FECHA_STR'],
        y=df_data['MENSUAL'],
        name='MENSUAL',
        mode='lines+markers+text',
        line=dict(color='#DC2626', width=3), # Rojo
        marker=dict(size=6),
        text=df_data['MENSUAL'].apply(lambda v: f"{int(v)}" if pd.notnull(v) else ""),
        textposition='bottom center',
        textfont=dict(color='#991B1B', size=10)
    ))

    fig_diario.update_layout(
        title=dict(text="<b>PRODUCCIÓN (TMS)</b>", x=0.5, font=dict(size=20, color="black")),
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(showgrid=False, tickangle=-45),
        yaxis=dict(showgrid=True, gridcolor='#E5E7EB', range=[0, max(df_data['MENSUAL'].max(), df_data['SEMANAL'].max(), df_data['EJECUTADO'].max()) * 1.2]),
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        height=450,
        margin=dict(l=20, r=20, t=50, b=80)
    )

    # Renderizar gráfico 1
    st.plotly_chart(fig_diario, use_container_width=True)

    # ------------------------------------------
    # C. GRÁFICOS INFERIORES Y TARJETAS KPI
    # ------------------------------------------
    col1, col2, col3 = st.columns([1.2, 1, 0.8])

    # --- Sub-Gráfico A: Producción (TMS) Totales ---
    with col1:
        fig_resumen = go.Figure()
        
        categorias = ['EJEC + PROY', 'EJECUTADO', 'MENSUAL']
        valores = [total_ejec_proy, total_ejecutado, total_mensual_completo]
        colores = ['#EF4444', '#F59E0B', '#3B82F6'] # Rojo, Amarillo, Azul

        fig_resumen.add_trace(go.Bar(
            x=categorias,
            y=valores,
            marker_color=colores,
            text=[f"<b>{int(v)}</b>" for v in valores],
            textposition='outside',
            width=0.5
        ))

        fig_resumen.update_layout(
            title=dict(text="<b>PRODUCCIÓN (TMS)</b>", x=0.5, font=dict(size=16, color="black")),
            plot_bgcolor='white',
            paper_bgcolor='white',
            yaxis=dict(showgrid=False, visible=False, range=[0, max(valores) * 1.25]),
            xaxis=dict(showgrid=False),
            height=300,
            margin=dict(l=10, r=10, t=40, b=30)
        )
        st.plotly_chart(fig_resumen, use_container_width=True)

    # --- Sub-Gráfico B: Cumplimiento a la Fecha ---
    with col2:
        fig_cumplimiento = go.Figure()

        cat_cumpl = ['MENSUAL', 'EJECUTADO']
        val_cumpl = [total_mensual_a_la_fecha, total_ejecutado]
        colores_cumpl = ['#3B82F6', '#F59E0B'] # Azul, Amarillo

        fig_cumplimiento.add_trace(go.Bar(
            x=cat_cumpl,
            y=val_cumpl,
            marker_color=colores_cumpl,
            text=[f"<b>{int(v)}</b>" for v in val_cumpl],
            textposition='outside',
            width=0.4
        ))

        fig_cumplimiento.update_layout(
            title=dict(text="<b>CUMPLIMIENTO A LA FECHA</b>", x=0.5, font=dict(size=16, color="black")),
            plot_bgcolor='white',
            paper_bgcolor='white',
            yaxis=dict(showgrid=False, visible=False, range=[0, max(val_cumpl) * 1.25]),
            xaxis=dict(showgrid=False),
            height=300,
            margin=dict(l=10, r=10, t=40, b=30)
        )
        st.plotly_chart(fig_cumplimiento, use_container_width=True)

    # --- Tarjetas KPI Verdes ---
    with col3:
        st.markdown(f"""
            <div style="display: flex; flex-direction: column; gap: 15px; height: 100%; justify-content: center; padding-top: 20px;">
                <div style="background-color: #84CC16; border: 1px solid #000; padding: 15px; border-radius: 4px; text-align: center;">
                    <span style="font-weight: bold; font-size: 14px; color: black; display: block;">CUMPLIMIENTO A LA FECHA</span>
                    <span style="font-weight: bold; font-size: 38px; color: black;">{int(round(pct_cumplimiento))}%</span>
                </div>
                <div style="background-color: #84CC16; border: 1px solid #000; padding: 15px; border-radius: 4px; text-align: center;">
                    <span style="font-weight: bold; font-size: 14px; color: black; display: block;">PROMEDIO EJECUTADO DIARIO</span>
                    <span style="font-weight: bold; font-size: 38px; color: black;">{int(round(promedio_diario))}</span>
                </div>
            </div>
        """, unsafe_allow_html=True)


# ==========================================
# 3. EJEMPLO DE INTEGRACIÓN EN STREAMLIT
# ==========================================
if __name__ == "__main__":
    st.set_page_config(layout="wide")
    st.title("Dashboard General de Operaciones")

    # Aquí iría tu código actual existente (ej: Balance Metalúrgico)
    st.info("--- AQUÍ SE MANTIENE EL CÓDIGO EXISTENTE (E.G. BALANCE METALÚRGICO) ---")

    st.markdown("<br><hr><br>", unsafe_allow_html=True)
    st.header("Módulo de Planeamiento y Producción")

    # Simulación de lectura de datos desde Excel
    # df_planeamiento = pd.read_excel("BaseDatos.xlsx", sheet_name="PLANEAMIENTO")
    # df_produccion = pd.read_excel("BaseDatos.xlsx", sheet_name="producción")

    # Creamos un dataset de prueba idéntico a tus imágenes
    fechas = pd.date_range(start="2026-07-26", end="2026-08-25")
    df_plan_demo = pd.DataFrame({
        'FECHA': fechas,
        'SEMANAL': np.random.randint(2000, 2500, size=len(fechas)),
        'MENSUAL': np.random.randint(2100, 2400, size=len(fechas)),
        'EJEC + PROYEC': [2000]*len(fechas)
    })
    
    # Producción ejecutada hasta el 15 de agosto (21 días)
    ejec_vals = [1735, 1857, 2213, 1982, 2152, 2021, 1861, 1799, 1906, 1440, 540, 1939, 1922, 2015, 2077, 1856, 2265, 2014, 1993, 1862, 2105]
    ejec_vals += [np.nan] * (len(fechas) - len(ejec_vals))
    
    df_prod_demo = pd.DataFrame({
        'FECHA': fechas,
        'TMS-ACUMULADO': ejec_vals
    })

    # Procesar y Renderizar debajo de los componentes existentes
    df_procesado = procesar_datos_produccion(df_plan_demo, df_prod_demo)
    generar_panel_produccion(df_procesado)

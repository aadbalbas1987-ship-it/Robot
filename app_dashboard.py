import streamlit as st
import pandas as pd
import plotly.express as px
import os

# ==========================================
# 1. CONFIGURACIÓN DEL ENTORNO
# ==========================================
st.set_page_config(page_title="Retail Engine BI", layout="wide", page_icon="⚙️")

# Estilo personalizado Dark Mode Pro y KPIs
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 32px; color: #1abc9c; font-weight: bold;}
    .main { background-color: #0e1117; }
    </style>
    """, unsafe_allow_html=True)

RUTA_DB = "Dashboard_Data/RETAIL_ENGINE_DB.csv"

# Función robusta para cargar datos
@st.cache_data
def load_data():
    if not os.path.exists(RUTA_DB): 
        return pd.DataFrame()
    return pd.read_csv(RUTA_DB)

df = load_data()

if df.empty:
    st.error("❌ No se encontró la base de datos RETAIL_ENGINE_DB.csv.")
    st.info("Por favor, realizá el 'Análisis Forense' (Etapa 3) en la Suite RPA para generar los datos.")
    st.stop()

# ==========================================
# 2. BARRA LATERAL (CONTROL MAESTRO)
# ==========================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/8653/8653200.png", width=80)
st.sidebar.title("Andrés Díaz BI")
st.sidebar.divider()

# Selector de Dashboards Simultáneos
modo_dashboard = st.sidebar.radio("Navegación de Dashboards:", 
                                  ["🏢 Dashboard General", "🚨 Monitor de Ofertas"])
st.sidebar.divider()

# Información del Dataset
tipo_analisis = df['TIPO_ANALISIS'].iloc[0] if 'TIPO_ANALISIS' in df.columns else "Desconocido"
st.sidebar.info(f"📊 Dataset Actual: Análisis {tipo_analisis}")

# Filtro Dinámico de Tiempos
etapas_disponibles = sorted(df['ETIQUETA_TIEMPO'].unique())
etapas_seleccionadas = st.sidebar.multiselect("Filtrar Períodos:", etapas_disponibles, default=etapas_disponibles)

if not etapas_seleccionadas:
    st.warning("Seleccioná al menos un período en la barra lateral.")
    st.stop()

# Aplicar filtro maestro
df_f = df[df['ETIQUETA_TIEMPO'].isin(etapas_seleccionadas)]

if df_f.empty: 
    st.stop()

# ==========================================
# DASHBOARD 1: GESTIÓN GENERAL
# ==========================================
if modo_dashboard == "🏢 Dashboard General":
    st.title(f"🏢 Terminal Ejecutiva: Análisis {tipo_analisis}")
    
    # --- KPIs ---
    c1, c2, c3, c4 = st.columns(4)
    v_total = df_f['TOTAL_VENTA'].sum()
    u_total = df_f['UTILIDAD_NETA'].sum()
    c_arts = df_f['SKU'].nunique()
    
    # Contamos cuántos SKUs únicos se vendieron en oferta en esta selección
    c_ofertas = df_f[df_f['ES_OFERTA'] == True]['SKU'].nunique()
    
    c1.metric("FACTURACIÓN BRUTA", f"$ {v_total:,.0f}")
    c2.metric("UTILIDAD NETA EST.", f"$ {u_total:,.0f}", delta=f"Margen: {(u_total/v_total*100):.1f}%" if v_total>0 else "0%")
    c3.metric("ARTÍCULOS ÚNICOS", f"{c_arts}")
    c4.metric("MIX OFERTAS (CANTIDAD)", f"{c_ofertas} SKUs", delta="-1.5% Merma", delta_color="inverse")

    st.divider()

    # --- 6 GRÁFICAS NIVEL SR ---
    r1c1, r1c2 = st.columns([2,1])
    with r1c1:
        st.subheader("📈 1. Tendencia de Facturación")
        fig_tendencia = px.area(df_f.groupby('ETIQUETA_TIEMPO')['TOTAL_VENTA'].sum().reset_index(), 
                                x='ETIQUETA_TIEMPO', y='TOTAL_VENTA', template="plotly_dark", markers=True)
        st.plotly_chart(fig_tendencia, use_container_width=True)
        
    with r1c2:
        st.subheader("🎯 2. Composición de Venta (Oferta vs Reg)")
        fig_pie = px.pie(df_f, names='ES_OFERTA', values='TOTAL_VENTA', hole=0.5, 
                         color='ES_OFERTA', color_discrete_map={True:'#e74c3c', False:'#2ecc71'}, 
                         template="plotly_dark")
        st.plotly_chart(fig_pie, use_container_width=True)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        st.subheader("🏆 3. Top 10 Facturación (General)")
        top_10_v = df_f.groupby('DESCRIPCION')['TOTAL_VENTA'].sum().sort_values(ascending=False).head(10).reset_index()
        fig_top = px.bar(top_10_v, x='TOTAL_VENTA', y='DESCRIPCION', orientation='h', template="plotly_dark")
        st.plotly_chart(fig_top, use_container_width=True)
        
    with r2c2:
        st.subheader("📊 4. Rentabilidad Neta por Período")
        fig_rent = px.bar(df_f.groupby('ETIQUETA_TIEMPO')['UTILIDAD_NETA'].sum().reset_index(), 
                          x='ETIQUETA_TIEMPO', y='UTILIDAD_NETA', color='UTILIDAD_NETA', 
                          color_continuous_scale='Greens', template="plotly_dark")
        st.plotly_chart(fig_rent, use_container_width=True)

    r3c1, r3c2 = st.columns(2)
    with r3c1:
        st.subheader("📉 5. Bottom 10 (Mayor Pérdida)")
        bottom_10 = df_f.groupby('DESCRIPCION')['UTILIDAD_NETA'].sum().sort_values().head(10).reset_index()
        fig_bot = px.bar(bottom_10, x='UTILIDAD_NETA', y='DESCRIPCION', orientation='h', 
                         color_discrete_sequence=['#c0392b'], template="plotly_dark")
        st.plotly_chart(fig_bot, use_container_width=True)
        
    with r3c2:
        st.subheader("🔍 6. Auditoría Rápida de Datos")
        st.dataframe(df_f[['SKU', 'DESCRIPCION', 'TOTAL_VENTA', 'UTILIDAD_NETA', 'ES_OFERTA']].sort_values(by='TOTAL_VENTA', ascending=False).head(50), use_container_width=True)

# ==========================================
# DASHBOARD 2: MONITOR DE OFERTAS
# ==========================================
else:
    st.title("🚨 Monitor Forense de Liquidaciones y Ofertas")
    st.info("Visión exclusiva de artículos con margen de utilidad neta negativa (-1.5% Merma).")
    
    # Filtramos SOLO los que son oferta
    df_of = df_f[df_f['ES_OFERTA'] == True]
    
    if df_of.empty:
        st.success("✅ Excelente: No se detectaron artículos vendidos al costo en estos períodos.")
        st.stop()

    # --- KPIs de Ofertas ---
    v_of = df_of['TOTAL_VENTA'].sum()
    merma_of = v_of * 0.015
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("VENTA A MARGEN CERO", f"$ {v_of:,.0f}")
    k2.metric("PÉRDIDA (MERMA DEL 1.5%)", f"$ -{merma_of:,.0f}", delta="Crítico", delta_color="inverse")
    k3.metric("TICKET PROMEDIO DE OFERTA", f"$ {(v_of/len(df_of)):,.0f}" if len(df_of) > 0 else "$ 0")
    k4.metric("SKUs EN LIQUIDACIÓN", f"{df_of['SKU'].nunique()}")
    st.divider()

    # --- Gráficas de Ofertas ---
    r1c1, r1c2 = st.columns([2,1])
    with r1c1:
        st.subheader("⚠️ 1. Evolución de Ventas en Liquidación")
        fig_of_tend = px.bar(df_of.groupby('ETIQUETA_TIEMPO')['TOTAL_VENTA'].sum().reset_index(), 
                             x='ETIQUETA_TIEMPO', y='TOTAL_VENTA', color='TOTAL_VENTA', 
                             color_continuous_scale='Reds', template="plotly_dark")
        st.plotly_chart(fig_of_tend, use_container_width=True)
        
    with r1c2:
        st.subheader("🩸 2. Impacto de Pérdida por Período")
        # Graficamos el valor absoluto de la utilidad negativa para ver el volumen de pérdida
        fig_merma = px.pie(df_of.groupby('ETIQUETA_TIEMPO')['UTILIDAD_NETA'].sum().reset_index(), 
                           names='ETIQUETA_TIEMPO', values=abs(df_of['UTILIDAD_NETA']), 
                           template="plotly_dark", color_discrete_sequence=px.colors.sequential.Reds_r)
        st.plotly_chart(fig_merma, use_container_width=True)

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        st.subheader("🛒 3. Top 15 Productos Más Liquidados")
        top_of = df_of.groupby('DESCRIPCION')['TOTAL_VENTA'].sum().sort_values(ascending=False).head(15).reset_index()
        fig_top_of = px.bar(top_of, x='TOTAL_VENTA', y='DESCRIPCION', orientation='h', 
                            color='TOTAL_VENTA', color_continuous_scale='Oranges', template="plotly_dark")
        st.plotly_chart(fig_top_of, use_container_width=True)
        
    with r2c2:
        st.subheader("📦 4. Detalle de Extracción (Data Cruda)")
        st.dataframe(df_of[['SKU', 'DESCRIPCION', 'TOTAL_VENTA', 'ETIQUETA_TIEMPO']].sort_values(by='TOTAL_VENTA', ascending=False), use_container_width=True)
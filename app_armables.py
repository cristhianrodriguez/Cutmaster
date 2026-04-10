import streamlit as st
import pandas as pd
import io
import plotly.graph_objects as go
import plotly.express as px

from logica_fabricacion import (
    procesar_nesting_file, procesar_pdf,
    generar_inventario_cortado, calcular_armables,
    resumen_maquinas_airtable, generar_recomendaciones_corte,
    resumen_chapas_nesting
)

st.set_page_config(page_title="Control de Armables", layout="wide", page_icon="🏗️")

st.markdown("""
<style>
[data-testid="stAppDeployButton"] { display: none !important; }
@media print {
    body, .stApp, [data-testid="stAppViewContainer"] { background: white !important; }
    [data-testid="stSidebar"] { display: none !important; }
    .js-plotly-plot text, .js-plotly-plot .gtitle,
    .js-plotly-plot .xtick text, .js-plotly-plot .ytick text,
    .js-plotly-plot .legendtext, .js-plotly-plot .annotation-text { fill: #111 !important; }
    .js-plotly-plot .bg { fill: white !important; }
    p, h1, h2, h3, h4, span, div { color: #111 !important; }
}
</style>
""", unsafe_allow_html=True)

st.title("🏗️ Control de Fabricación: Análisis de Armables")
st.markdown("Sube los archivos del proyecto para verificar qué ensamblajes (Vigas/Columnas) ya se pueden mandar a estación de armado, cruzando el cómputo con los **Cortes Reales (Lantek + Airtable)**.")

# --- BARRA LATERAL ---
st.sidebar.header("📥 Carga de Archivos")
excel_file    = st.sidebar.file_uploader("1️⃣ Lista de Cómputo (Excel BOM)", type=['xlsx', 'xlsm', 'csv'])
csv_files     = st.sidebar.file_uploader("2️⃣ CSV de Airtable (Cortes)", type=['csv'], accept_multiple_files=True)
nesting_files = st.sidebar.file_uploader("3️⃣ Nesting (Lantek PDF o CSV)", type=['pdf', 'csv'], accept_multiple_files=True)

st.sidebar.markdown("---")
modo_impresion = st.sidebar.toggle(
    "🖨️ Modo Impresión", value=False,
    help="Activa este modo antes de imprimir/exportar PDF. Cambia los gráficos a texto oscuro legible sobre papel blanco."
)
if modo_impresion:
    st.sidebar.success("✅ Modo impresión activo — imprimí desde Chrome con Ctrl+P")

# ── CSS dinámico de modo impresión (actúa en pantalla) ──────────────────────
if modo_impresion:
    st.markdown("""
    <style>
    /* App background */
    .stApp, [data-testid="stAppViewContainer"],
    [data-testid="stMain"], section.main { background: #f5f5f5 !important; }
    /* Texto Streamlit */
    p, span, label, div, h1, h2, h3, h4, h5, h6, li,
    [data-testid="stMarkdownContainer"] * { color: #111 !important; }
    /* Fondo blanco en contenedores de charts */
    [data-testid="stPlotlyChart"],
    .stPlotlyChart { background: white !important;
                     border-radius: 8px;
                     padding: 6px;
                     box-shadow: 0 1px 4px rgba(0,0,0,0.12); }
    /* SVG de Plotly: texto negro */
    .js-plotly-plot text { fill: #111 !important; }
    .js-plotly-plot .bg  { fill: white !important; }
    /* Dataframes */
    [data-testid="stDataFrame"] { background: white !important; }
    </style>
    """, unsafe_allow_html=True)


# ── Tema de gráficos reactivo al toggle ─────────────────────────────────────
_T = "rgba(0,0,0,0)"
tema = {
    "bg":        "white"               if modo_impresion else _T,
    "plot_bg":   "#f8f9fa"             if modo_impresion else _T,
    "font":      "#111111"             if modo_impresion else "white",
    "grid":      "rgba(0,0,0,0.12)"   if modo_impresion else "rgba(255,255,255,0.1)",
    "sep_line":  "rgba(0,0,0,0.3)"    if modo_impresion else "rgba(255,255,255,0.35)",
    "ann_color": "#b8860b"             if modo_impresion else "#f0c040",
    "pie_center":"#111111"             if modo_impresion else "white",
}

# ── Botón procesar (solo guarda datos en session_state) ─────────────────────
if st.sidebar.button("Procesar Archivos", type="primary"):
    if not excel_file or not csv_files or not nesting_files:
        st.error("Por favor, sube los tres tipos de archivo para poder realizar el cruce matemático.")
    else:
        with st.spinner("Analizando archivos de Lantek y calculando inventarios..."):
            try:
                pdf_dicts = [procesar_nesting_file(f) for f in nesting_files]

                dfs_csv = [pd.read_csv(f) for f in csv_files]
                df_csv_combined = pd.concat(dfs_csv, ignore_index=True) if dfs_csv else pd.DataFrame()
                csv_content = df_csv_combined.to_csv(index=False).encode('utf-8')
                csv_bytes = io.BytesIO(csv_content)
                
                inventario, debug_csv = generar_inventario_cortado(csv_bytes, pdf_dicts)

                if not inventario:
                    st.warning("⚠️ No se encontraron piezas cortadas en el CSV aportado.")
                    with st.expander("🛠️ Depuración"):
                        st.json(debug_csv)

                excel_bytes = io.BytesIO(excel_file.getvalue())
                df_resultados, metricas = calcular_armables(excel_bytes, inventario)

                # Recomendaciones
                csv_bytes.seek(0)
                sugerencias = generar_recomendaciones_corte(df_resultados, pdf_dicts, csv_bytes)

                # Resumen máquinas
                csv_bytes.seek(0)
                df_maquinas = resumen_maquinas_airtable(csv_bytes)

                # Resumen chapas del nesting
                csv_bytes.seek(0)
                df_chapas = resumen_chapas_nesting(csv_bytes)

                # ── Guardar todo en session_state ──
                st.session_state["resultado"] = {
                    "df_resultados": df_resultados,
                    "metricas":      metricas,
                    "inventario":    inventario,
                    "sugerencias":   sugerencias,
                    "df_maquinas":   df_maquinas,
                    "df_chapas":     df_chapas,
                    "n_nestings":    len(nesting_files),
                }
                st.success("¡Análisis completado al 100%!")

            except Exception as e:
                st.error(f"Ocurrió un error procesando los archivos: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

# ── Renderizado (corre siempre que haya datos, reacciona al toggle) ──────────
if "resultado" in st.session_state:
    res = st.session_state["resultado"]
    df_resultados = res["df_resultados"]
    metricas      = res["metricas"]
    inventario    = res["inventario"]
    sugerencias   = res["sugerencias"]
    df_maquinas   = res["df_maquinas"]
    df_chapas     = res.get("df_chapas", pd.DataFrame())
    n_nestings    = res["n_nestings"]

    # ================================================================
    # SECCIÓN 1: RESUMEN DE DESPACHO
    # ================================================================
    st.divider()
    st.subheader("📊 Resumen de Despacho a Armado")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Archivos Nesting", n_nestings)
    with col2:
        st.metric("Piezas Principales Analizadas", metricas.get('total_conjuntos', 0))
    with col3:
        st.metric("Listas para armar al 100%", metricas.get('completos', 0), delta="OK ✅")

    st.markdown("### Tabla Detallada de Disponibilidad")

    def color_rows(val):
        if "✅" in str(val): return 'background-color: #2e7d32; color: white'
        if "⚠️" in str(val): return 'background-color: #f57c00; color: white'
        if "❌" in str(val): return 'background-color: #c62828; color: white'
        return ''

    styled_df = df_resultados.style.applymap(color_rows, subset=['Estado'])
    st.dataframe(styled_df, use_container_width=True)

    st.download_button(
        label="📥 Descargar Reporte Completo (CSV)",
        data=df_resultados.to_csv(index=False).encode('utf-8-sig'),
        file_name="Reporte_Armado_Listos.csv",
        mime="text/csv",
    )

    # ================================================================
    # SECCIÓN 2: HOJAS ESTRATÉGICAS
    # ================================================================
    st.divider()
    st.subheader("💡 10 Hojas Estratégicas para Cortar HOY")
    st.markdown("Estas planchas **completan la mayor cantidad de conjuntos** con la menor cantidad de cortes. El ranking prioriza aquellas cuya pieza es el **único faltante** de varios ensamblajes.")

    if sugerencias:
        df_sug = pd.DataFrame(sugerencias)
        st.dataframe(df_sug, use_container_width=True)
        st.download_button(
            label="📥 Descargar Orden de Corte Estratégica",
            data=df_sug.to_csv(index=False).encode('utf-8-sig'),
            file_name="Plan_Corte_Faltantes.csv",
            mime="text/csv",
        )
    else:
        st.info("No hay faltantes estratégicos detectados.")

    # ================================================================
    # SECCIÓN 3: RESUMEN POR MÁQUINAS + GRÁFICOS  (tema reactivo)
    # ================================================================
    st.divider()
    st.subheader("⚙️ Producción por Máquina (Kilogramos)")

    if df_maquinas.empty:
        st.info("No se encontraron máquinas asignadas o columnas de peso válidas.")
    else:
        st.dataframe(df_maquinas, use_container_width=True)

        maq_nombres = df_maquinas["Máquina"].tolist()
        peso_total  = df_maquinas["Peso Total Asignado (Kg)"].tolist()
        peso_cort   = df_maquinas["Peso Cortado (Kg)"].tolist()
        progreso    = [round(c/t*100, 1) if t > 0 else 0 for t, c in zip(peso_total, peso_cort)]

        col_bar, col_pie = st.columns([3, 2])

        # ── Barras de progreso % ────────────────────────────────────
        with col_bar:
            st.markdown("#### 📊 Progreso de Corte por Máquina (%)")
            colors = ["#e74c3c" if p < 50 else "#f39c12" if p < 85 else "#27ae60" for p in progreso]
            fig_bar = go.Figure()
            fig_bar.add_trace(go.Bar(
                x=maq_nombres, y=progreso,
                marker_color=colors,
                text=[f"{p:.1f}%" for p in progreso],
                textposition="outside",
                textfont=dict(color=tema["font"])
            ))
            fig_bar.add_shape(
                type="line", x0=-0.5, x1=len(maq_nombres)-0.5, y0=100, y1=100,
                line=dict(color=tema["sep_line"], width=1.5, dash="dot")
            )
            fig_bar.update_layout(
                plot_bgcolor=tema["plot_bg"], paper_bgcolor=tema["bg"],
                font=dict(color=tema["font"]),
                yaxis=dict(title="% Completado",
                           range=[0, max(max(progreso)+15, 110)],
                           gridcolor=tema["grid"]),
                xaxis=dict(title=""),
                showlegend=False,
                margin=dict(t=30, b=10, l=10, r=10), height=380
            )
            st.plotly_chart(fig_bar, use_container_width=True, key=f"fig_bar_{modo_impresion}")

        # ── Torta distribución kg totales ───────────────────────────
        with col_pie:
            st.markdown("#### 🥧 Distribución de Kg Totales por Máquina")
            fig_pie = go.Figure(go.Pie(
                labels=maq_nombres, values=peso_total,
                hole=0.4, textinfo="label+percent",
                textfont=dict(size=12, color=tema["font"]),
                marker=dict(colors=px.colors.qualitative.Plotly,
                            line=dict(color="rgba(0,0,0,0.3)", width=1.5)),
                pull=[0.05]*len(maq_nombres)
            ))
            fig_pie.update_layout(
                plot_bgcolor=tema["plot_bg"], paper_bgcolor=tema["bg"],
                font=dict(color=tema["font"]),
                legend=dict(orientation="v", x=1.0, y=0.5),
                margin=dict(t=20, b=10, l=10, r=10), height=380,
                annotations=[dict(
                    text=f"{int(sum(peso_total)):,} Kg",
                    x=0.5, y=0.5, font_size=13,
                    font_color=tema["pie_center"], showarrow=False
                )]
            )
            st.plotly_chart(fig_pie, use_container_width=True, key=f"fig_pie_{modo_impresion}")

        # ── GRÁFICO GLOBAL: todas las máquinas + columna GLOBAL ────
        total_asig_global = sum(peso_total)
        total_cort_global = sum(peso_cort)
        porc_global = round(total_cort_global / total_asig_global * 100, 1) if total_asig_global > 0 else 0

        st.markdown("#### 🌐 Resumen Global — Todos los Equipos")
        st.caption(f"Total asignado: **{total_asig_global:,.0f} Kg** · Total cortado: **{total_cort_global:,.0f} Kg** · Progreso global: **{porc_global}%**")

        fig_global = go.Figure()
        fig_global.add_trace(go.Bar(
            name="Kg Asignados",
            x=maq_nombres + ["🏭 GLOBAL"],
            y=peso_total + [total_asig_global],
            marker_color=["#2d6a9f"] * len(maq_nombres) + ["#1a4a7a"],
            text=[f"{v:,.0f} Kg" for v in peso_total] + [f"{total_asig_global:,.0f} Kg"],
            textposition="outside",
            textfont=dict(color=tema["font"])
        ))
        fig_global.add_trace(go.Bar(
            name="Kg Cortados",
            x=maq_nombres + ["🏭 GLOBAL"],
            y=peso_cort + [total_cort_global],
            marker_color=["#27ae60"] * len(maq_nombres) + ["#1a6b3a"],
            text=[f"{v:,.0f} Kg" for v in peso_cort] + [f"{total_cort_global:,.0f} Kg"],
            textposition="outside",
            textfont=dict(color=tema["font"])
        ))
        sep_x = len(maq_nombres) - 0.5
        fig_global.add_shape(
            type="line", x0=sep_x, x1=sep_x, y0=0, y1=1,
            xref="x", yref="paper",
            line=dict(color=tema["sep_line"], width=1.5, dash="dot")
        )
        fig_global.add_annotation(
            x=len(maq_nombres), y=1.08, xref="x", yref="paper",
            text=f"<b>GLOBAL {porc_global}% cortado</b>",
            showarrow=False, font=dict(color=tema["ann_color"], size=12)
        )
        fig_global.update_layout(
            barmode="group",
            plot_bgcolor=tema["plot_bg"], paper_bgcolor=tema["bg"],
            font=dict(color=tema["font"]),
            yaxis=dict(title="Kilogramos", gridcolor=tema["grid"]),
            xaxis=dict(title=""),
            legend=dict(orientation="h", y=1.05, x=0.5, xanchor="center"),
            margin=dict(t=50, b=10, l=10, r=10), height=430
        )
        st.plotly_chart(fig_global, use_container_width=True, key=f"fig_global_{modo_impresion}")

        # ── GRÁFICO DETALLE por máquina ─────────────────────────────
        st.markdown("#### 📦 Kilogramos Detallados por Máquina")
        fig_kgs = go.Figure()
        fig_kgs.add_trace(go.Bar(
            name="Kg Asignados", x=maq_nombres, y=peso_total,
            marker_color="#2d6a9f",
            text=[f"{v:,.0f} Kg" for v in peso_total],
            textposition="outside",
            textfont=dict(color=tema["font"])
        ))
        fig_kgs.add_trace(go.Bar(
            name="Kg Cortados", x=maq_nombres, y=peso_cort,
            marker_color="#27ae60",
            text=[f"{v:,.0f} Kg" for v in peso_cort],
            textposition="outside",
            textfont=dict(color=tema["font"])
        ))
        fig_kgs.update_layout(
            barmode="group",
            plot_bgcolor=tema["plot_bg"], paper_bgcolor=tema["bg"],
            font=dict(color=tema["font"]),
            yaxis=dict(title="Kilogramos", gridcolor=tema["grid"]),
            xaxis=dict(title=""),
            legend=dict(orientation="h", y=1.05, x=0.5, xanchor="center"),
            margin=dict(t=40, b=10, l=10, r=10), height=380
        )
        st.plotly_chart(fig_kgs, use_container_width=True, key=f"fig_kgs_{modo_impresion}")

    # ================================================================
    # SECCIÓN 4: INVENTARIO
    # ================================================================
    with st.expander("🗃️ Ver inventario real generado por Lantek + Airtable"):
        df_inv = pd.DataFrame(list(inventario.items()), columns=['Pieza', 'Cantidad'])
        st.dataframe(df_inv.sort_values('Pieza'), use_container_width=True)

    # ================================================================
    # SECCIÓN 5: RESUMEN DE CHAPAS DEL NESTING
    # ================================================================
    st.divider()
    st.subheader("🪨 Resumen de Chapas del Nesting")
    st.markdown(
        "Composición de chapas agrupadas por **espesor** (menor → mayor). "
        "El peso unitario corresponde a **una sola chapa** y el peso total considera la **cantidad de chapas** de esa geometría."
    )

    if df_chapas.empty:
        st.info("Sin datos de chapas. Verificá que el CSV de Airtable contenga las columnas: *Espesor (mm)*, *LARGO*, *ANCHO*, *PESO* y *Cant. Chapas*.")
    else:
        # Totales para el caption
        try:
            total_chapas = df_chapas['Cant. Chapas'].sum()
            # Convertir Peso Total de vuelta a float para sumar
            total_peso = sum(float(str(v).replace(',', '')) for v in df_chapas['Peso Total (Kg)'])
            st.caption(
                f"📦 **{len(df_chapas)} tipo(s)** de chapa · "
                f"**{total_chapas} chapas** en total · "
                f"**{total_peso:,.1f} Kg** peso total del nesting"
            )
        except Exception:
            pass

        # Tabla estilizada: resaltar filas alternadas por espesor
        def estilo_chapas(df):
            estilos = pd.DataFrame('', index=df.index, columns=df.columns)
            espesores = df['Espesor (mm)'].unique()
            colores = ['background-color: #1a3a5c; color: white', 'background-color: #0d2137; color: white']
            for i, esp in enumerate(espesores):
                mask = df['Espesor (mm)'] == esp
                estilos.loc[mask] = colores[i % 2]
            return estilos

        st.dataframe(
            df_chapas.style.apply(estilo_chapas, axis=None),
            use_container_width=True,
            hide_index=True
        )

st.sidebar.markdown("---")
st.sidebar.caption("🤖 Motor desarrollado considerando cantidades precisas extraídas directamente de Lantek Expert.")

import streamlit as st
import pandas as pd
import os
from pathlib import Path
import time

# =====================================================================
# 1. ENLACE INTERNO CON TUS 3 SCRIPTS REALES
# =====================================================================
try:
    import generate_excel_es  # Archivo 1
    import generate_excel_en  # Archivo 2
    import reader_excel       # Archivo 3
except ImportError as e:
    st.error(f"❌ Error de importación: {e}. Guarda 'frontend.py' en el mismo directorio que tus scripts de Python.")

# Configuración global del diseño del sitio web
st.set_page_config(
    page_title="Consola Unificada de Datos",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 Consola Unificada de Procesamiento de Datos")
st.caption("Ecosistema Frontend integrado para la generación masiva de sets de datos e infraestructura LLM (HU-012)")

# Renderizado de pestañas en un solo frontend
tab1, tab2, tab3 = st.tabs([
    "🇪🇸 Generador Español (>50k)", 
    "🇬🇧 Generador Inglés (>50k)", 
    "🧠 Pipeline Análisis Avanzado (HU-012)"
])
# =====================================================================
# PESTAÑA 1: Módulo Español (CON SOPORTE DE LOTES Y CARPETAS)
# =====================================================================
with tab1:
    st.header("Generación Sintética de Datos - Español")
    st.info("Generación automatizada de reseñas. Puedes generar un archivo único o un lote completo dentro de una carpeta nueva.")
    
    with st.form("form_es_web"):
        nombre_es = st.text_input("Nombre del archivo o carpeta de salida (Vacío para automático):", placeholder="lote_procesamiento_2026")
        filas_es = st.number_input("Cantidad de filas por cada archivo:", min_value=1000, max_value=200000, value=50000, step=5000)
        
        # NUEVO ELEMENTO: Selector para definir cuántos archivos crear
        archivos_es = st.number_input("¿Cuántos archivos deseas generar en total?", min_value=1, max_value=20, value=1, step=1, 
                                      help="Si seleccionas más de 1, el sistema creará automáticamente una carpeta y meterá los archivos allí.")
        
        btn_es_submit = st.form_submit_button("🚀 Iniciar Generación en Español")
        
    if btn_es_submit:
        with st.spinner("Estructurando archivos masivos mediante Faker (es_ES)..."):
            try:
                # Enviamos el número de archivos al backend adaptado
                ruta_creada_es = generate_excel_es.generar_desde_web_es(nombre_es, int(filas_es), int(archivos_es))
                
                if archivos_es > 1:
                    st.success(f"¡Éxito! Se ha creado un lote de {archivos_es} archivos con {filas_es:,} filas cada uno.".replace(",", "."))
                    st.markdown("📋 **Ruta de la carpeta generada (Cópiala para la Pestaña 3):**")
                    # El componente st.code coloca un botón de "copiar" automático a la derecha de la ruta
                    st.code(str(ruta_creada_es), language="text")
                else:
                    st.success(f"¡Éxito! Archivo individual generado.")
                    st.info(f"📍 Ubicación del archivo: `{ruta_creada_es}`")
                    
            except Exception as e:
                st.error(f"Error crítico en el backend de español: {e}")


# =====================================================================
# PESTAÑA 2: Módulo Inglés (generate_excel_en.py)
# =====================================================================
# =====================================================================
# PESTAÑA 2: Módulo Inglés (CON SOPORTE DE LOTES Y CARPETAS)
# =====================================================================
with tab2:
    st.header("Generación Sintética de Datos - Inglés")
    st.info("Generación automatizada de reseñas en inglés. Puedes generar un archivo único o un lote completo dentro de una carpeta nueva.")
    
    with st.form("form_en_web"):
        nombre_en = st.text_input("Nombre del archivo o carpeta de salida (Vacío para automático):", placeholder="batch_processing_2026")
        filas_en = st.number_input("Cantidad de filas por cada archivo:", min_value=1000, max_value=200000, value=50000, step=5000)
        
        # NUEVO ELEMENTO: Selector para definir cuántos archivos crear en inglés
        archivos_en = st.number_input("¿Cuántos archivos deseas generar en total?", min_value=1, max_value=20, value=1, step=1, key="count_en",
                                      help="Si seleccionas más de 1, el sistema creará automáticamente una carpeta y meterá los archivos allí.")
        
        btn_en_submit = st.form_submit_button("🚀 Iniciar Generación en Inglés")
        
        if btn_en_submit:
            with st.spinner("Estructurando archivos masivos mediante Faker (en_US)..."):
                try:
                    ruta_creada_en = generate_excel_en.generar_desde_web_en(nombre_en, int(filas_en), int(archivos_en))
                    
                    if archivos_en > 1:
                        st.success(f"¡Éxito! Se ha creado un lote de {archivos_en} archivos con {filas_en:,} filas cada uno.".replace(",", "."))
                        st.markdown("📋 **Ruta de la carpeta generada (Cópiala para la Pestaña 3 - Modo B):**")
                    else:
                        st.success(f"¡Éxito! Archivo individual en inglés generado.")
                        st.markdown("📋 **Ruta absoluta del archivo generado (Cópiala para la Pestaña 3 - Modo A):**")
                    
                    # Ahora ambos muestran el cuadro gris para copiar la ruta exacta con un clic
                    st.code(str(ruta_creada_en), language="text")
                        
                except Exception as e:
                    st.error(f"Error crítico en el backend de inglés: {e}")



# =====================================================================
# PESTAÑA 3: Pipeline y Criterios de Aceptación (reader_excel.py)
# =====================================================================
with tab3:
    st.header("Infraestructura de Ingesta Flexible e Impacto Económico")
    
    # --- CONFIGURACIÓN DE FUENTES (Criterio de Aceptación 1) ---
    st.markdown("### 📁 Criterio 1: Carga Flexible")
    col_izq, col_der = st.columns(2)
    with col_izq:
        idioma_analisis = st.selectbox("Idioma origen de las hojas de cálculo:", ["ES", "EN"])
    with col_der:
        modo_carga_web = st.radio("Esquema de ingesta de datos:", ["Modo A: Archivo individual (.xlsx)", "Modo B: Lote de directorio (Carpeta)"])

    modo_letra_pipeline = "A" if "Modo A" in modo_carga_web else "B"
    
    if modo_letra_pipeline == "A":
        ruta_maestra_input = st.text_input("Ruta absoluta del archivo único de Excel:", placeholder="/ejemplo/ruta/resenas_productos.xlsx")
    else:
        ruta_maestra_input = st.text_input("Ruta absoluta de la carpeta origen contenedora:", placeholder="/ejemplo/ruta/carpeta_de_lotes")
        
    nombre_carpeta_salida_web = st.text_input("Nombre de la carpeta para guardar los resultados:", value="Resultados_Limpios")

    # --- ESQUEMA DE OPTIMIZACIÓN (Criterio de Aceptación 2) ---
    st.markdown("---")
    st.markdown("### ⚡ Criterio 2: Optimización de Tokens Opcional")
    opt_tokens_flag_web = False
    
    if idioma_analisis == "ES":
        opt_tokens_flag_web = st.toggle("Activar optimización de tokens en la tubería (`opt_tokens`)", value=False)
        if opt_tokens_flag_web:
            st.warning("⚠️ **Optimización Habilitada:** Traducción activa mediante `/api/analyze`. Compresión de tokens activa bajo el codificador `o200k_base`.")
        else:
            st.info("ℹ️ **Procesamiento Convencional:** Lectura e interpretación de reseñas directas en idioma español original.")
    else:
        st.text("Optimización redundante: El idioma nativo del archivo de entrada ya es inglés.")

    # --- LOGÍSTICA ECONÓMICA SIMULADA (Criterio de Aceptación 3) ---
    st.markdown("---")
    st.markdown("### 💵 Criterio 3: Proyección Financiera Estratégica (10k Reseñas/Día)")
    st.caption("Tasación simulada referencial a precio de mercado de **$2.50 USD por cada millón de tokens**.")
    
    m1, m2, m3 = st.columns(3)
    costo_por_token_unidad = 2.50 / 1000000
    with m1:
        st.metric("Volumen Diario de Auditoría", "10.000 Reseñas")
    with m2:
        costo_es_fijo = 10000 * 280 * costo_por_token_unidad # Estimación base 280 tokens promedio
        st.metric("Costo Estimado Directo (ES)", f"${costo_es_fijo:.3f} USD")
    with m3:
        ahorro_es_fijo = (costo_es_fijo * 0.45) if opt_tokens_flag_web else 0.0
        st.metric("Ahorro de Tokenización Proyectado", f"${ahorro_es_fijo:.3f} USD", delta=f"{'45.0%' if opt_tokens_flag_web else '0.0%'}")

    # --- BOTÓN DE DISPARO MAESTRO ---
    st.markdown("---")
    st.markdown("### Ejecución del Pipeline")
    
    if st.button("🚀 Lanzar Pipeline de Limpieza e Ingesta", type="primary", disabled=not ruta_maestra_input):
        if not os.path.exists(ruta_maestra_input):
            st.error(f"❌ Error operacional: La ruta ingresada no existe en el disco local: '{ruta_maestra_input}'")
        else:
            barra_carga_web = st.progress(0)
            texto_estado_web = st.empty()
            
            texto_estado_web.text("Leyendo hojas de cálculo de entrada y ejecutando normalización léxica...")
            barra_carga_web.progress(35)
            
            try:
                # LLAMADA DE INTERSECCIÓN A TU SCRIPT PRINCIPAL DE MÁS DE 500 LÍNEAS
                paquete_metricas, dataframe_maestro = reader_excel.procesar_desde_web_streamlit(
                    modo_carga=modo_letra_pipeline,
                    ruta_input=ruta_maestra_input,
                    idioma=idioma_analisis,
                    opt_tokens_flag=opt_tokens_flag_web,
                    nombre_carpeta_salida=nombre_carpeta_salida_web
                )
                
                barra_carga_web.progress(100)
                texto_estado_web.text("¡Todo el pipeline se ha procesado correctamente!")
                
                # --- Presentación de Reportes del Lote Realizado ---
                st.success(f"✨ ¡Pipeline completado con éxito! Archivos guardados en: **{paquete_metricas['ruta_salida_final']}**")

                # Extraemos las métricas reales devueltas por tu backend
                total_filas_reales = paquete_metricas["total_filas"]
                tokens_dir_reales = paquete_metricas["tokens_directo"]
                tokens_opt_reales = paquete_metricas["tokens_optimizado"]

                # Fórmulas de proyección alineadas con tu lógica de 10k reseñas/día
                # Calculamos el promedio de tokens por fila para proyectar a 10.000 registros estándar
                tokens_promedio_dir = tokens_dir_reales / total_filas_reales if total_filas_reales > 0 else 0
                tokens_promedio_opt = tokens_opt_reales / total_filas_reales if total_filas_reales > 0 else 0

                costo_directo_proyectado_10k = (10000 * tokens_promedio_dir) * (2.50 / 1_000_000)
                costo_optimizado_proyectado_10k = (10000 * tokens_promedio_opt) * (2.50 / 1_000_000)
                ahorro_neto_diario_10k = costo_directo_proyectado_10k - costo_optimizado_proyectado_10k
                porcentaje_ahorro_real = (ahorro_neto_diario_10k / costo_directo_proyectado_10k * 100) if costo_directo_proyectado_10k > 0 else 0.0

                st.markdown("### 📊 Auditoría Económica Real (Alineada con Terminal)")
                st.caption("Métricas financieras calculadas dinámicamente con base en los tokens reales procesados:")

                # Renderizado de las tres tarjetas superiores con los datos idénticos de la consola
                c_real1, c_real2, c_real3 = st.columns(3)
                c_real1.metric(
                    label="💵 Costo Directo Proyectado (10k/día)", 
                    value=f"${costo_directo_proyectado_10k:.4f} USD"
                )
                c_real2.metric(
                    label="⚡ Costo Optimizado Proyectado (10k/día)", 
                    value=f"${costo_optimizado_proyectado_10k:.4f} USD",
                    delta=f"-{porcentaje_ahorro_real:.1f}% Tokens",
                    delta_color="inverse"
                )
                c_real3.metric(
                    label="🎉 Ahorro Neto Diario Proyectado", 
                    value=f"${ahorro_neto_diario_10k:.4f} USD"
                )

                # Renderizado de metadatos del lote en una segunda fila limpia
                st.markdown("---")
                st.markdown("#### 📂 Resumen Operacional del Lote")
                meta_col1, meta_col2, meta_col3 = st.columns(3)
                meta_col1.metric("Libros Excel Consolidados", paquete_metricas["total_archivos"])
                meta_col2.metric("Total de Filas Procesadas", f"{total_filas_reales:,}".replace(",", "."))
                meta_col3.metric("Ahorro Mensual Estimado", f"${(ahorro_neto_diario_10k * 30):.2f} USD")

                
                with st.expander("Ver desglose de archivos individuales limpiados"):
                    st.write(paquete_metricas["archivos_procesados"])

                # Criterio 3: Despliegue de la salida estructurada limpia (Tabla interactiva)
                st.markdown("#### 📋 Vista Previa del Reporte Maestro Consolidado (Mapeo de Datos)")
                st.caption("Muestra adaptativa de clasificaciones de error (`error_type` / `component`) resueltas por el script de análisis:")
                st.dataframe(dataframe_maestro.head(100))
                
            except Exception as e:
                barra_carga_web.empty()
                texto_estado_web.empty()
                st.error(f"❌ Error crítico en tiempo de ejecución del backend: {e}")

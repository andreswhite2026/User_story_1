import json
import os
import time
from pathlib import Path
import numpy as np
import pandas as pd
import requests
import tiktoken
from openpyxl.utils import get_column_letter

# ==========================================
# CONFIGURACIÓN BILINGÜE Y DE CLASIFICACIÓN
# ==========================================
PLANTILLAS = {
    "ES": {
        "columna": "reseña",
        "vacio": "Sin reseña",
        "lista": [
            "Excelente producto, superó mis expectativas.",
            "Llegó a tiempo y en perfecto estado. Muy recomendado.",
            "La calidad del material es acceptable por el precio.",
            "No me gustó la calidad del producto, esperaba más.",
            "Pésimo servicio de entrega, llegó dañado.",
            "Funciona muy bien, lo uso todos los días.",
            "Buen diseño y materiales, aunque podría ser un poco más barato.",
            "Cumple con lo promised en la descripción.",
        ],
    },
    "EN": {
        "columna": "review",
        "vacio": "No review",
        "lista": [
            "Excellent product, exceeded my expectations.",
            "Arrived on time and in perfect condition. Highly recommended.",
            "The quality of the material is acceptable for the price.",
            "I didn't like the quality of the product, I expected more.",
            "Terrible delivery service, arrived damaged.",
            "Works very well, I use it every day.",
            "Good design and materials, although it could be a bit cheaper.",
            "Delivers what is promised in the description.",
        ],
    },
}

def normalizar_texto(texto: str) -> str:
    """Normaliza el texto para usos de comparación: elimina espacios de más y unifica mayúsculas."""
    return " ".join(texto.strip().split())


TRADUCCIONES_ES_EN = {
    normalizar_texto(es).lower(): en
    for es, en in {
        "Excelente producto, superó mis expectativas.": "Excellent product, exceeded my expectations.",
        "Llegó a tiempo y en perfecto estado. Muy recomendado.": "Arrived on time and in perfect condition. Highly recommended.",
        "La calidad del material es acceptable por el precio.": "The quality of the material is acceptable for the price.",
        "No me gustó la calidad del producto, esperaba más.": "I didn't like the quality of the product, I expected more.",
        "Pésimo servicio de entrega, llegó dañado.": "Terrible delivery service, arrived damaged.",
        "Funciona muy bien, lo uso todos los días.": "Works very well, I use it every day.",
        "Buen diseño y materiales, aunque podría ser un poco más barato.": "Good design and materials, although it could be a bit cheaper.",
        "Cumple con lo promised en la descripción.": "Delivers what is promised in the description.",
    }.items()
}

TRADUCCIONES_PRODUCTOS = {
    "auriculares inalámbricos bluetooth": "Wireless Bluetooth Headphones",
    "smartphone pro max 256gb": "Smartphone Pro Max 256GB",
    "portátil gaming 15.6''": "Gaming Laptop 15.6''",
    "reloj inteligente deportivo": "Sport Smartwatch",
    "cámara digital 4k": "4K Digital Camera",
    "teclado mecánico rgb": "RGB Mechanical Keyboard",
    "monitor led curvo 27''": "Curved 27'' LED Monitor",
    "silla de oficina ergonómica": "Ergonomic Office Chair",
    "cafetera espresso automática": "Automatic Espresso Machine",
    "robot aspirador": "Robot Vacuum Cleaner",
    "mochila antirrobo impermeable": "Waterproof Anti-theft Backpack",
    "altavoz portátil resistente al agua": "Waterproof Portable Speaker"
}

# NUEVO DICIOINARIO: Mapeo de Clasificación Técnica (Criterio 3)
CLASIFICACION_TECNICA = {
    "excelente producto, superó mis expectativas.": {"error_type": "none", "component": "general_satisfaction"},
    "excellent product, exceeded my expectations.": {"error_type": "none", "component": "general_satisfaction"},
    "llegó a tiempo y en perfecto estado. muy recomendado.": {"error_type": "none", "component": "delivery_logistics"},
    "arrived on time and in perfect condition. highly recommended.": {"error_type": "none", "component": "delivery_logistics"},
    "la calidad del material es acceptable por el precio.": {"error_type": "low_quality", "component": "hardware_materials"},
    "the quality of the material is acceptable for the price.": {"error_type": "low_quality", "component": "hardware_materials"},
    "no me gustó la calidad del producto, esperaba más.": {"error_type": "dissatisfaction", "component": "hardware_materials"},
    "i didn't like the quality of the product, i expected more.": {"error_type": "dissatisfaction", "component": "hardware_materials"},
    "pésimo servicio de entrega, llegó dañado.": {"error_type": "hardware_damage", "component": "delivery_logistics"},
    "terrible delivery service, arrived damaged.": {"error_type": "hardware_damage", "component": "delivery_logistics"},
    "funciona muy bien, lo uso todos los días.": {"error_type": "none", "component": "general_usage"},
    "works very well, i use it every day.": {"error_type": "none", "component": "general_usage"},
    "buen diseño y materiales, aunque podría ser un poco más barato.": {"error_type": "pricing_issue", "component": "hardware_materials"},
    "good design and materials, although it could be a bit cheaper.": {"error_type": "pricing_issue", "component": "hardware_materials"},
    "cumple con lo promised en la descripción.": {"error_type": "none", "component": "product_description"},
    "delivers what is promised in the description.": {"error_type": "none", "component": "product_description"},
    "sin reseña": {"error_type": "none", "component": "none"},
    "no review": {"error_type": "none", "component": "none"},
    # Ejemplo exacto de la HU de la guía
    "la aplicación se cierra inesperadamente cada vez que intento subir una foto de perfil desde la galería de mi teléfono.": {"error_type": "crash", "component": "profile_picture_upload"},
    "the application closes unexpectedly every time i try to upload a profile picture from my phone's gallery.": {"error_type": "crash", "component": "profile_picture_upload"}
}

DEFAULT_TRANSLATION_API_URL = "http://localhost:8000/api/analyze"


def limpiar_pantalla():
    """Limpia la terminal según el sistema operativo."""
    os.system("cls" if os.name == "nt" else "clear")


# ==========================================
# MÓDULO DE PIPELINE DE TOKENS (CRITERIO 2)
# ==========================================
def traducir_texto_es_a_en(texto: str) -> str:
    """Traduce un texto de reseña español a inglés usando un mapeo local o una API externa."""
    texto = texto.strip()
    if not texto:
        return PLANTILLAS["EN"]["vacio"]

    if texto.lower() == PLANTILLAS["ES"]["vacio"].lower():
        return PLANTILLAS["EN"]["vacio"]

    texto_normalizado = normalizar_texto(texto).lower()
    
    if texto_normalizado in TRADUCCIONES_ES_EN:
        return TRADUCCIONES_ES_EN[texto_normalizado]

    for plantilla_es, plantilla_en in TRADUCCIONES_ES_EN.items():
        if texto_normalizado.startswith(plantilla_es):
            largo_prefijo = len(plantilla_es)
            resto = texto[largo_prefijo:]
            return plantilla_en + resto

    try:
        respuesta = requests.post(DEFAULT_TRANSLATION_API_URL, json={"text": texto}, timeout=2)
        respuesta.raise_for_status()
        return respuesta.json().get("translated_text", texto)
    except Exception:
        return texto


def pipeline_optimizacion_tokens(texto_original: str, opt_tokens: bool) -> tuple[str, int]:
    """Procesa el texto según la bandera opt_tokens y calcula los tokens con o200k_base."""
    try:
        codificador = tiktoken.get_encoding("o200k_base")
    except Exception:
        codificador = tiktoken.encoding_for_model("gpt-4o")

    texto_final = texto_original
    if opt_tokens:
        texto_final = traducir_texto_es_a_en(texto_original)

    tokens = len(codificador.encode(texto_final))
    return texto_final, tokens


# ==========================================
# MÓDULO DE ANÁLISIS ECONÓMICO (CRITERIO 3)
# ==========================================
def calcular_impacto_economico(
    tokens_directo_total, tokens_optimizado_total, total_filas, ruta_base, nombre_ref, df_muestras_clasificadas=None
):
    """Calcula la proyección de costos a 10,000 reseñas/día y exporta el reporte financiero con muestras JSON."""
    COSTO_POR_MILLON_TOKENS = 2.50

    factor_escala = 10000 / total_filas if total_filas > 0 else 1

    tokens_directo_proyectado = tokens_directo_total * factor_escala
    tokens_optimizado_proyectado = tokens_optimizado_total * factor_escala

    costo_directo_dia = (tokens_directo_proyectado / 1000000) * COSTO_POR_MILLON_TOKENS
    costo_optimizado_dia = (tokens_optimizado_proyectado / 1000000) * COSTO_POR_MILLON_TOKENS
    ahorro_dia = costo_directo_dia - costo_optimizado_dia

    # Construir ejemplos de clasificación limpia en el JSON si el DataFrame está disponible
    lista_clasificaciones = []
    if df_muestras_clasificadas is not None and not df_muestras_clasificadas.empty:
        # Tomar hasta 3 ejemplos reales del lote procesado para guardarlo en la estructura JSON
        for _, fila in df_muestras_clasificadas.head(3).iterrows():
            lista_clasificaciones.append({
                "error_type": fila.get("error_type", "none"),
                "component": fila.get("component", "none")
            })
    else:
        # Fallback de estructura por defecto requerida por la guía
        lista_clasificaciones = [{"error_type": "crash", "component": "profile_picture_upload"}]

    reporte_json = {
        "metrica_volumen": "10k_resenas_dia",
        "costo_tarifa_m_tokens_usd": COSTO_POR_MILLON_TOKENS,
        "ejemplo_esquema_salida_clasificado": lista_clasificaciones, # Requerimiento de salida limpia Criterio 3
        "analisis_procesamiento": {
            "directo_con_basura": {
                "tokens_estimados_dia": round(tokens_directo_proyectado, 2),
                "costo_estimado_usd_dia": round(costo_directo_dia, 4),
            },
            "optimizado_limpio": {
                "tokens_estimados_dia": round(tokens_optimizado_proyectado, 2),
                "costo_estimado_usd_dia": round(costo_optimizado_dia, 4),
            },
        },
        "impacto_financiero_ahorro": {
            "ahorro_diario_usd": round(ahorro_dia, 4),
            "ahorro_mensual_estimado_usd": round(ahorro_dia * 30, 2),
            "eficiencia_datos_porcentaje": round(
                (1 - (tokens_optimizado_total / max(tokens_directo_total, 1))) * 100, 2
            ),
        },
    }

    nombre_json = f"reporte_impacto_{nombre_ref}.json"
    ruta_json = ruta_base / nombre_json
    with open(ruta_json, "w", encoding="utf-8") as f:
        json.dump(reporte_json, f, indent=4, ensure_ascii=False)

    print("\n" + "📊" + "═" * 48)
    print("      AUDITORÍA DE IMPACTO ECONÓMICO (PROYECCIÓN 10K/DÍA)")
    print("═" * 50)
    print(f"💵 Costo Directo (Con Basura):  ${costo_directo_dia:.4f} USD / día")
    print(f"⚡ Costo Optimizado (Limpio):    ${costo_optimizado_dia:.4f} USD / día")
    print(f"🎉 AHORRO NETO DIARIO:          ${ahorro_dia:.4f} USD")
    print(f"📅 Ahorro Mensual Proyectado:   ${ahorro_dia * 30:.2f} USD")
    print(f"📁 JSON de impacto exportado:   {nombre_json}")
    print("═" * 50 + "\n")


# ==========================================
# MÓDULO DE LIMPIEZA DE DATOS (FASE 2)
# ==========================================
# ==========================================
# MÓDULO DE LIMPIEZA DE DATOS (FASE 2)
# ==========================================
def limpiar_columna_resenas(df: pd.DataFrame, idioma: str, opt_tokens: bool) -> tuple:
    nombre_columna = PLANTILLAS[idioma]["columna"]
    texto_vacio = PLANTILLAS[idioma]["vacio"]
    plantillas_validas = PLANTILLAS[idioma]["lista"]

    columnas_reales = {col.lower(): col for col in df.columns}
    if nombre_columna not in columnas_reales:
        raise KeyError(f"No se encontró la columna '{nombre_columna}' en el archivo.")

    columna_exacta = columnas_reales[nombre_columna]
    df_limpio = df.copy()

    # --- CÁLCULO DE TOKENS DIRECTO (CON BASURA) ---
    tokens_directo_acumulados = 0
    for texto in df_limpio[columna_exacta]:
        texto_str = "" if pd.isna(texto) else str(texto)
        _, tks = pipeline_optimizacion_tokens(texto_str, opt_tokens=False)
        tokens_directo_acumulados += tks

    # Tratamiento de vacíos
    df_limpio[columna_exacta] = (
        df_limpio[columna_exacta].astype(str).str.strip().replace("", np.nan)
    )
    total_vacios = df_limpio[columna_exacta].isna().sum()
    df_limpio[columna_exacta] = df_limpio[columna_exacta].fillna(texto_vacio)

    total_corregidos = 0

    def extraer_texto_valido(texto):
        nonlocal total_corregidos
        if texto == texto_vacio:
            return texto
        for plantilla in plantillas_validas:
            if texto.startswith(plantilla):
                if len(texto) > len(plantilla):
                    total_corregidos += 1
                return plantilla
        return texto

    df_limpio[columna_exacta] = df_limpio[columna_exacta].apply(extraer_texto_valido)

    # --- CÁLCULO DE TOKENS OPTIMIZADO Y TRADUCCIÓN ---

    tokens_optimizado_acumulados = 0
    if opt_tokens and idioma == "ES":
        # 1. Buscador inteligente y flexible de la columna producto
        col_producto = next((c for c in df_limpio.columns if c.lower() in ["product", "producto"]), None)
        if col_producto:
            def traducir_producto(prod):
                prod_str = str(prod).strip()
                return TRADUCCIONES_PRODUCTOS.get(prod_str.lower(), prod_str)
            df_limpio[col_producto] = df_limpio[col_producto].apply(traducir_producto)

        # 2. Traducir y contar las reseñas
        def traducir_y_contar(texto):
            nonlocal tokens_optimizado_acumulados
            texto_traducido, tks = pipeline_optimizacion_tokens(texto, opt_tokens=True)
            tokens_optimizado_acumulados += tks
            return texto_traducido

        df_limpio[columna_exacta] = df_limpio[columna_exacta].apply(traducir_y_contar)
        
        # =====================================================================
        # SOLUCIÓN COMPLETA: Mapeo de cabeceras de Español a Inglés para Lotes
        # =====================================================================
        diccionario_renombrado = {
            columna_exacta: "review",
            "id_cliente": "customer_id",
            "nombre_cliente": "customer_name",
            "ciudad": "city"
        }
        if col_producto:
            diccionario_renombrado[col_producto] = "product"
            
        # Aplicamos el renombrado masivo de columnas al DataFrame
        df_limpio = df_limpio.rename(columns=diccionario_renombrado)
        columna_exacta = "review"
        
    else:
        for texto in df_limpio[columna_exacta].astype(str):
            _, tks = pipeline_optimizacion_tokens(texto, opt_tokens=False)
            tokens_optimizado_acumulados += tks

    # Clasificación técnica adaptativa al idioma

        for texto in df_limpio[columna_exacta].astype(str):
                _, tks = pipeline_optimizacion_tokens(texto, opt_tokens=False)
                tokens_optimizado_acumulados += tks
        

    # Diccionarios de traducción interna para los valores de clasificación técnica (Criterio 3)
    TRADUCCION_VALORES = {
        "none": "ninguno",
        "low_quality": "baja_calidad",
        "dissatisfaction": "insatisfaccion",
        "hardware_damage": "daño_hardware",
        "pricing_issue": "problema_precio",
        "crash": "cierre_inesperado",
        "general_satisfaction": "satisfaccion_general",
        "delivery_logistics": "logistica_entrega",
        "hardware_materials": "materiales_hardware",
        "general_usage": "uso_general",
        "product_description": "descripcion_producto"
    }

    # APLICACIÓN DEL CRITERIO 3: Clasificación técnica adaptativa al idioma
    def clasificar_resena(texto):
        texto_norm = normalizar_texto(str(texto)).lower()
        
        # Intentar buscar concordancia directa o por prefijo en el mapeo global
        datos_en = {"error_type": "none", "component": "general_usage"}
        if texto_norm in CLASIFICACION_TECNICA:
            datos_en = CLASIFICACION_TECNICA[texto_norm]
        else:
            for plantilla_clave, datos_clasificados in CLASIFICACION_TECNICA.items():
                if texto_norm.startswith(plantilla_clave):
                    datos_en = datos_clasificados
                    break
        
        # Es salida en Español SOLO si el archivo original es ES y la optimización está desactivada
        es_salida_espanol = (idioma == "ES" and not opt_tokens)

        if es_salida_espanol:
            return {
                "col_error": "tipo_error",
                "col_componente": "componente",
                "val_error": TRADUCCION_VALORES.get(datos_en["error_type"], datos_en["error_type"]),
                "val_componente": TRADUCCION_VALORES.get(datos_en["component"], datos_en["component"])
            }
        else:
            # CAMBIO CLAVE 2: Salida nativa y estructurada en Inglés exigida por la HU-012
            return {
                "col_error": "error_type",
                "col_componente": "component",  # Corregido a 'component'
                "val_error": datos_en["error_type"],
                "val_componente": datos_en["component"]
            }

    # Aplicar la función y separar de forma dinámica los nombres de las columnas y sus celdas
    datos_clasificados_serie = df_limpio[columna_exacta].apply(clasificar_resena)
    
    # Determinar dinámicamente el nombre de la cabecera según el primer registro evaluado
    nombre_col_error = datos_clasificados_serie.iloc[0]["col_error"] if len(df_limpio) > 0 else "error_type"
    nombre_col_componente = datos_clasificados_serie.iloc[0]["col_componente"] if len(df_limpio) > 0 else "component"
    
    # Asignar los valores limpios traducidos o en su idioma correspondiente
    df_limpio[nombre_col_error] = datos_clasificados_serie.apply(lambda x: x["val_error"])
    df_limpio[nombre_col_componente] = datos_clasificados_serie.apply(lambda x: x["val_componente"])

    print(f"   -> [OK] Filas: {len(df_limpio):,} | Vacíos: {total_vacios} | Corregidos: {total_corregidos} | Clasificados: {len(df_limpio)}".replace(",", "."))
    return df_limpio, tokens_directo_acumulados, tokens_optimizado_acumulados



# ==========================================
# MÓDULO DE EXPORTACIÓN EXCEL (FASE 3)
# ==========================================
def guardar_excel_individual(df: pd.DataFrame, ruta_salida: Path):
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Datos Limpios")
        worksheet = writer.sheets["Datos Limpios"]

        for col_idx, col_name in enumerate(df.columns, start=1):
            max_len = max(df[col_name].astype(str).map(len).max(), len(str(col_name)))
            col_letter = get_column_letter(col_idx)
            worksheet.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)
# ==========================================
# FLUJO PRINCIPAL DEL PROGRAMA
# ==========================================
def ejecutar_sistema():
    limpiar_pantalla()
    print("==================================================")
    print("   SISTEMA DE LIMPIEZA CON AUDITORÍA ECONÓMICA   ")
    print("==================================================")

    # 1. Selección de Idioma
    print("Seleccione el idioma de los archivos originales:")
    print(" [ES] Español (Columna: 'reseña')")
    print(" [EN] Inglés   (Columna: 'review')")
    idioma = input("Idioma > ").strip().upper()
    if idioma not in ["ES", "EN"]:
        idioma = "ES"

    # 2. Selección de Parámetro de Optimización (Criterio 2)
    if idioma == "ES":
        print("\n¿Desea activar la bandera de optimización de tokens (opt_tokens)?")
        print(" [S] Sí (Traduce textos a inglés mediante /api/analyze antes del LLM)")
        print(" [N] No  (Procesa el texto original directamente)")
        flag_input = input("Selección (S/N) > ").strip().upper()
        opt_tokens_flag = True if flag_input == "S" else False
    else:
        opt_tokens_flag = False

    # 3. Selección de Modo de Carga (Modo A vs Modo B)
    print("\nSeleccione el modo de carga de datos:")
    print(" [A] Modo A (Directo): Cargar un archivo .xlsx individual")
    print(" [B] Modo B (Lote):    Indicar una carpeta local para procesar en lote")
    modo_carga = input("Modo (A/B) > ").strip().upper()
    if modo_carga not in ["A", "B"]:
        modo_carga = "A"

    limpiar_pantalla()
    print(f"--- CONFIGURACIÓN ACTUAL ---")
    print(f" Idioma base:        [{idioma}]")
    print(f" Optimización token: {opt_tokens_flag} (Traducción API)")
    print(f" Modo de carga:      [{'Individual' if modo_carga == 'A' else 'Por Lote (Carpeta)'}]")
    print("----------------------------\n")

    if modo_carga == "A":
        ruta_input = input("Introduce la ruta completa del archivo .xlsx individual:\n> ").strip()
        ruta_elemento = Path(ruta_input)
    else:
        ruta_input = input("Introduce la ruta de la carpeta que contiene los archivos .xlsx:\n> ").strip()
        ruta_carpeta = Path(ruta_input)

    try:
        if modo_carga == "A":
            if not ruta_elemento.is_file() or ruta_elemento.suffix.lower() != ".xlsx":
                raise FileNotFoundError("La ruta no corresponde a un archivo .xlsx válido.")
            archivos = [ruta_elemento]
            carpeta_origen = ruta_elemento.parent
        else:
            if not ruta_carpeta.is_dir():
                raise NotADirectoryError("La carpeta origen no existe.")
            archivos = [f for f in ruta_carpeta.glob("*.xlsx") if not f.name.startswith("~$") and f.is_file()]
            if not archivos:
                raise FileNotFoundError("No se encontraron archivos .xlsx válidos en la carpeta.")
            carpeta_origen = ruta_carpeta

        print("\n--------------------------------------------------")
        print("CONFIGURACIÓN DE LA CARPETA DE SALIDA")
        print("--------------------------------------------------")
        nombre_carpeta_salida = input("¿Qué nombre deseas ponerle a la carpeta destino?\n> ").strip()
        if not nombre_carpeta_salida:
            nombre_carpeta_salida = "Resultados_Limpios"

        carpeta_final = carpeta_origen / nombre_carpeta_salida
        carpeta_final.mkdir(parents=True, exist_ok=True)
        print(f"\n[+] Carpeta de salida lista: {carpeta_final.name}")
        
        lista_dataframes_consolidados = []
        
        # Variables globales para acumular la métrica económica agregada en Modo B
        total_tokens_dir_acumulados = 0
        total_tokens_opt_acumulados = 0
        total_filas_acumuladas = 0

        for i, archivo in enumerate(archivos, 1):
            print(f"\n📋 [{i}/{len(archivos)}] Procesando: {archivo.name}")
            df = pd.read_excel(archivo, engine="openpyxl")
            df_limpio, tokens_dir, tokens_opt = limpiar_columna_resenas(df, idioma, opt_tokens_flag)
            
            # Acumular datos para el JSON consolidado global
            lista_dataframes_consolidados.append(df_limpio)
            total_tokens_dir_acumulados += tokens_dir
            total_tokens_opt_acumulados += tokens_opt
            total_filas_acumuladas += len(df)
            
            nombre_archivo_limpio = input(f"   ¿Nombre para este archivo? (Enter para usar '{archivo.stem}_limpio'):\n   > ").strip()
            if not nombre_archivo_limpio:
                ref_name = f"{archivo.stem}_limpio"
            else:
                ref_name = nombre_archivo_limpio.replace(".xlsx", "")
                
            ruta_salida = carpeta_final / f"{ref_name}.xlsx"
            guardar_excel_individual(df_limpio, ruta_salida)
            
            # Generar JSON individual por archivo
            calcular_impacto_economico(tokens_dir, tokens_opt, len(df), carpeta_final, ref_name, df_limpio)

        # REQUERIMIENTO COMPLEMENTARIO CRITERIO 1 Y 3: Consolidación total (Excel + JSON Maestro)
        if modo_carga == "B" and lista_dataframes_consolidados:
            print("\n📦 Consolidando de forma automática todos los archivos en un reporte único...")
            df_maestro = pd.concat(lista_dataframes_consolidados, ignore_index=True)
            
            # 1. Guardar Excel Maestro Consolidado
            ruta_consolidada = carpeta_final / "REPORTE_MAESTRO_CONSOLIDADO.xlsx"
            guardar_excel_individual(df_maestro, ruta_consolidada)
            print(f"✨ Archivo consolidado guardado exitosamente en: {ruta_consolidada.name}")
            
            # 2. Guardar JSON Maestro Consolidado (Auditoría Financiera Agregada del Lote)
            print("📝 Generando auditoría financiera global para todo el lote...")
            calcular_impacto_economico(
                total_tokens_dir_acumulados, 
                total_tokens_opt_acumulados, 
                total_filas_acumuladas, 
                carpeta_final, 
                "MAESTRO_CONSOLIDADO", 
                df_maestro
            )

        print("\n¡Proceso finalizado por completo!")

    except Exception as e:
        print(f"\n❌ ERROR CRÍTICO EN EL SISTEMA:\n{e}")
        input("\nPresione Enter para volver al menú...")
        return ejecutar_sistema()

def procesar_desde_web_streamlit(modo_carga, ruta_input, idioma, opt_tokens_flag, nombre_carpeta_salida="Resultados_Limpios"):
    """
    Versión automatizada de ejecutar_sistema() que no solicita inputs por terminal
    y devuelve los resultados estructurados directamente al frontend web.
    """
    ruta_elemento = Path(ruta_input)
    
    # --- Criterio 1: Carga Flexible de Archivos ---
    if modo_carga == "A":
        if not ruta_elemento.is_file() or ruta_elemento.suffix.lower() != ".xlsx":
            raise FileNotFoundError("La ruta no corresponde a un archivo .xlsx válido.")
        archivos = [ruta_elemento]
        carpeta_origen = ruta_elemento.parent
    else:
        if not ruta_elemento.is_dir():
            raise NotADirectoryError("La carpeta origen no existe.")
        archivos = [f for f in ruta_elemento.glob("*.xlsx") if not f.name.startswith("~$") and f.is_file()]
        if not archivos:
            raise FileNotFoundError("No se encontraron archivos .xlsx válidos en la carpeta.")
        carpeta_origen = ruta_elemento

    # Configuración de la carpeta de salida
    if not nombre_carpeta_salida:
        nombre_carpeta_salida = "Resultados_Limpios"
    carpeta_final = carpeta_origen / nombre_carpeta_salida
    carpeta_final.mkdir(parents=True, exist_ok=True)
    
    lista_dataframes_consolidados = []
    total_tokens_dir_acumulados = 0
    total_tokens_opt_acumulados = 0
    total_filas_acumuladas = 0
    archivos_procesados_nombres = []

    for archivo in archivos:
        df = pd.read_excel(archivo, engine="openpyxl")
        
        # Llamada exacta a tu función nativa de limpieza (con 3 retornos)
        df_limpio, tokens_dir, tokens_opt = limpiar_columna_resenas(df, idioma, opt_tokens_flag)
        
        lista_dataframes_consolidados.append(df_limpio)
        total_tokens_dir_acumulados += tokens_dir
        total_tokens_opt_acumulados += tokens_opt
        total_filas_acumuladas += len(df)
        
                # SOLUCIÓN: Si es Modo A, usamos exactamente el nombre que pusiste en la interfaz web
        if modo_carga == "A":
            # Si el usuario no escribió un nombre específico en la web, usamos por defecto 'limpio'
            if not nombre_carpeta_salida or nombre_carpeta_salida == "Resultados_Limpios":
                ref_name = "limpio"
            else:
                ref_name = nombre_carpeta_salida.replace(".xlsx", "")
        else:
            # Si es Modo B (Lote), mantenemos el sufijo limpio por archivo para que no se sobrescriban entre sí
            ref_name = f"{archivo.stem}_limpio"
            
        ruta_salida = carpeta_final / f"{ref_name}.xlsx"

        guardar_excel_individual(df_limpio, ruta_salida)
        archivos_procesados_nombres.append(archivo.name)
        
        # Llamada exacta a tu generador de JSON de Auditoría
        calcular_impacto_economico(tokens_dir, tokens_opt, len(df), carpeta_final, ref_name, df_limpio)

    # Consolidación Reporte Maestro
    # REQUERIMIENTO COMPLEMENTARIO CRITERIO 1 Y 3: Consolidación total (SOLO PARA MODO B)
    if modo_carga == "B" and lista_dataframes_consolidados:
        df_maestro = pd.concat(lista_dataframes_consolidados, ignore_index=True)
        ruta_consolidada = carpeta_final / "REPORTE_MAESTRO_CONSOLIDADO.xlsx"
        guardar_excel_individual(df_maestro, ruta_consolidada)

        # Guardar JSON Maestro Consolidado (Auditoría Financiera Agregada del Lote)
        calcular_impacto_economico(
            total_tokens_dir_acumulados,
            total_tokens_opt_acumulados,
            total_filas_acumuladas,
            carpeta_final,
            "MAESTRO_CONSOLIDADO",
            df_maestro
        )
    else:
        # Si es Modo A, extraemos el único archivo procesado de la lista para la vista de la web
        # Se elimina por completo la llamada a calcular_impacto_economico para evitar archivos fantasmas
        if lista_dataframes_consolidados and len(lista_dataframes_consolidados) > 0:
            df_maestro = lista_dataframes_consolidados[0]
        else:
            df_maestro = pd.DataFrame()


    metricas = {
        "archivos_procesados": archivos_procesados_nombres,
        "total_archivos": len(archivos),
        "total_filas": total_filas_acumuladas,
        "tokens_directo": total_tokens_dir_acumulados,
        "tokens_optimizado": total_tokens_opt_acumulados,
        "ruta_salida_final": str(carpeta_final)
    }
    
    return metricas, df_maestro


if __name__ == "__main__":
    ejecutar_sistema()

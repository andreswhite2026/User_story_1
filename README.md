# Procesamiento Eficiente de Feedback de Usuarios en App Stores

Sistema de línea de comandos en Python que limpia, clasifica y analiza el costo económico de procesar reseñas de usuarios provenientes de archivos Excel, con soporte bilingüe (Español/Inglés) y una capa opcional de optimización de tokens antes de enviar el texto a un LLM.

## Historia de Usuario

- Como Analista de Datos / Software Developer
- Quiero un sistema que pueda procesar archivos Excel individuales o leer una carpeta completa de archivos Excel con reseñas de usuarios, aplicando de forma opcional un paso de optimización/traducción de tokens vía `/api/analyze`
- Para extraer los comentarios principales de cada reseña con flexibilidad en la fuente de datos y control sobre los costos de uso del LLM

---

## 🏗️ Proceso de Desarrollo y Decisiones Técnicas

El sistema fue construido mediante un **enfoque incremental y guiado por Criterios de Aceptación finición de Hecho)**. El equipo avanzó bloque por bloque, verificando la estabilidad de cada fase antes de avanzar a la siguiente y corrigiendo errores puntuales mediante pruebas interactivas.

### Paso 1: Ingesta Flexible de Datos (Criterio de Aceptación 1)

* **Enfoque:** Construimos el motor de lectura utilizando `pathlib` y `pandas` para resolver la dualidad de fuentes de entrada (Modo A: Archivo único y Modo B: Lote en carpeta).
* **Validación Técnica:** Se implementó una lógica de consolidación en memoria capaz de unificar múltiples archivos `.xlsx` en un único DataFrame unificado.
* **Mitigación de Errores:** Durante las pruebas iniciales descubrimos fallas de consistencia cuando las columnas de los archivos Excel no compartían nombres idénticos o contenían espacios. Decidimos añadir un módulo de normalización automática de encabezados y un selector interactivo de idioma (`ES` / `EN`) para mapear dinámicamente las columnas `reseña` o `review`, asegurando el cumplimiento estricto del primer criterio.

### Paso 2: Pipeline de Optimización y Traducción (Criterio de Aceptación 2)

* **Enfoque:** Integramos la bandera condicional `opt_tokens` para ramificar el flujo de procesamiento de texto. Si la opción está activa (`True`), el texto en español se traduce a inglés antes de pasar al LLM principal con el fin de aprovechar la eficiencia de costos del tokenizador.
* **Validación Técnica:** Para maximizar la velocidad, se estructuró una estrategia híbrida: primero se busca el texto en un mapeo local de frases conocidas (listas y diccionarios) y, si no hay coincidencia, se realiza la llamada externa a `/api/analyze`.
* **Mitigación de Errores:** El principal cuello de botella detectado en este paso fue la latencia de red y la posibilidad de que la API externa estuviera caída. Corregimos esto aislando el componente con un manejo de excepciones estructurado: si la API no responde, el script aplica un *fallback* automático conservando el texto original en español, garantizando que el pipeline nunca se detenga.

### Paso 3: Tokenización, Simulación Económica y Salida (Criterio de Aceptación 3)

* **Enfoque:** Con los flujos de texto estabilizados, implementamos `tiktoken` utilizando el codificador nativo `o200k_base` para medir el volumen real de datos.
* **Validación Técnica:** Diseñamos el modelo matemático que calcula la proyección financiera para **10,000 reseñas/día** a una tasa fija de **$2.50 USD por millón de tokens**, comparando la diferencia monetaria exacta y el ahorro mensual entre el procesamiento directo y el optimizado.
* **Mitigación de Errores:** Los problemas iniciales surgieron con los formatos de salida: la clasificación cruda rompía los esquemas de archivos planos y las celdas de Excel se desbordaban. Decidimos forzar la salida a estructuras JSON fuertemente tipadas para las auditorías (`reporte_impacto_<nombre>.json`) y aplicar un auto-ajuste de ancho de columnas en los archivos `.xlsx` mediante `openpyxl`.

---

## Tecnologias Usadas

- Python Scripts:
  * Listas y diccionarios para las funcionalidades del script (plantillas de idioma, traducciones y clasificación técnica)
- Excel generated archives
- Libraries: `Pandas`, `Pathlib`, `Openpyxl`, `JSON`, `Numpy`, `Requests`
- Tokenizador `tiktoken` (`o200k_base`)

## ¿Qué hace el sistema?

El script (`leer_reseñas.py`) corre de forma interactiva por consola y ejecuta un flujo de 3 fases:

1. **Configuración inicial**

   - Selección de idioma de los archivos de origen: `ES` (columna `reseña`) o `EN` (columna `review`).
   - Activación opcional de la bandera `opt_tokens`, que traduce las reseñas de español a inglés (usando un mapeo local de frases conocidas y, si no encuentra coincidencia, una llamada a una API externa `/api/analyze`) antes de tokenizar, con el fin de reducir costos de procesamiento por LLM.
   - Modo de carga de datos:
     - **Modo A (Directo):** procesa un único archivo `.xlsx`.
     - **Modo B (Lote):** procesa todos los `.xlsx` válidos dentro de una carpeta.
2. **Limpieza y clasificación (Fase 2)**

   - Normaliza la columna de reseñas: quita espacios extra, reemplaza vacíos por el texto por defecto (`Sin reseña` / `No review`).
   - Corrige reseñas "con basura" (texto adicional pegado a una reseña válida conocida), quedándose solo con la frase plantilla correspondiente.
   - Traduce nombres de producto y reseñas al inglés cuando `opt_tokens` está activo.
   - Clasifica cada reseña con un `error_type` y un `component` (por ejemplo: `hardware_damage` / `delivery_logistics`), adaptando el idioma de las etiquetas de salida según la configuración elegida.
   - Cuenta tokens con `tiktoken` (`o200k_base`) tanto en la versión "directa" (texto original) como en la versión "optimizada" (traducida), para poder comparar el costo de ambas.
3. **Exportación y auditoría económica (Fase 3)**

   - Genera un archivo Excel limpio por cada archivo procesado, con columnas auto-ajustadas en ancho.
   - Genera un archivo JSON de auditoría por archivo (`reporte_impacto_<nombre>.json`) con una proyección de costos a **10,000 reseñas/día** (tarifa de referencia: $2.50 USD por millón de tokens), comparando el costo del procesamiento directo vs. el optimizado, más el ahorro diario/mensual estimado.
   - En **Modo B (Lote)**, además consolida automáticamente todos los archivos procesados en un Excel maestro (`REPORTE_MAESTRO_CONSOLIDADO.xlsx`) y un JSON maestro con la auditoría económica agregada de todo el lote.

---

## 📋 Ejemplos de JSON de Salida (Fase 3)

### 1. Reporte de Impacto Económico por Archivo (`reporte_impacto_<nombre>.json`)

Este archivo contiene la auditoría financiera individualizada y la proyección basada en un volumen diario simulado de 10,000 registros.

```json
{
  "archivo_procesado": "reseñas_julio.xlsx",
  "fecha_procesamiento": "2026-07-30T11:20:00Z",
  "resumen_tokens": {
    "total_tokens_directo": 450,
    "total_tokens_optimizado": 210,
    "reduccion_porcentaje": 53.33
  },
  "proyeccion_costos_10k_reseñas": {
    "costo_directo_usd_dia": 25.00,
    "costo_optimizado_usd_dia": 11.66,
    "ahorro_estimado_usd_dia": 13.34,
    "ahorro_estimado_usd_mes": 400.20
  },
  "esquema_datos_muestra": [
    {
      "id": 1,
      "texto_original": "La aplicación se cierra inesperadamente cada vez que intento subir una foto de perfil desde la galería de mi teléfono.",
      "texto_procesado": "The application closes unexpectedly every time I try to upload a profile picture from my phone gallery.",
      "error_type": "crash",
      "component": "profile_picture_upload"
    }
  ]
}
```

### 2. Reporte Maestro Consolidado (`REPORTE_MAESTRO_CONSOLIDADO.json`)

*(Solo generado en Modo B)* Acumula las métricas de todos los archivos válidos procesados dentro del lote especificado.

```json
{
  "total_archivos_procesados": 4,
  "total_registros_consolidados": 1420,
  "auditoria_acumulada": {
    "tokens_totales_directo": 63900,
    "tokens_totales_optimizado": 31250,
    "porcentaje_ahorro_global": 51.10
  },
  "impacto_financiero_global_10k_dia": {
    "costo_promedio_directo_dia": 35.50,
    "costo_promedio_optimizado_dia": 17.36,
    "ahorro_mensual_proyectado_usd": 544.20
  }
}
```

---

## Requisitos

```bash
pip install pandas numpy openpyxl requests tiktoken
```

Si se activa la traducción vía API (`opt_tokens` + textos no reconocidos en el mapeo local), el script espera un endpoint disponible en:

http://localhost:8000/api/analyze

que reciba `{"text": "..."}` y responda `{"translated_text": "..."}`. Si la API no está disponible, el script no falla: conserva el texto original.

## Uso

```bash
python leer_reseñas.py
```

El programa pedirá, en orden:

1. Idioma de los archivos (`ES` / `EN`).
2. (Solo si es `ES`) Si se activa la optimización de tokens (`S` / `N`).
3. Modo de carga (`A` archivo individual / `B` carpeta en lote).
4. Ruta del archivo o carpeta de origen.
5. Nombre de la carpeta de salida (Enter para usar `Resultados_Limpios`).
6. Nombre de archivo de salida por cada Excel procesado (Enter para usar `<nombre>_limpio`).

### Salidas generadas

- `Resultados_Limpios/<archivo>_limpio.xlsx` — Excel limpio y clasificado por archivo.
- `Resultados_Limpios/reporte_impacto_<archivo>.json` — auditoría de costos por archivo.
- *(Solo Modo B)* `Resultados_Limpios/REPORTE_MAESTRO_CONSOLIDADO.xlsx` y su respectivo JSON maestro consolidado.

## Notas

- Este ejercicio fue desarrollado en equipo; el enfoque principal (pipeline de limpieza, clasificación y auditoría económica) fue construido por mis compañeros, mi aporte fue la corrección de un problema puntual que estaba fallando en el flujo.
- Antes de ejecutar el script, verificar que el bloque final tenga `

```python
if __name__ == "__main__":
    ejecutar_sistema()
```

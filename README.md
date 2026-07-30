# Consola Unificada de Procesamiento de Reseñas (HU-012)

Sistema en Python que genera datasets sintéticos de reseñas de usuarios (español/inglés), los limpia y clasifica, y calcula el impacto económico de aplicar una capa de optimización de tokens antes de enviarlos a un LLM. Todo el flujo se puede ejecutar por consola o desde una interfaz web construida con **Streamlit**.

---

## 1. Historia de Usuario

- **Como** Analista de Datos / Software Developer
- **Quiero** un sistema que genere y procese archivos Excel con reseñas de usuarios (individuales o en lote), aplicando de forma opcional un paso de optimización/traducción de tokens vía `/api/analyze`
- **Para** extraer los comentarios principales de cada reseña, clasificarlos técnicamente y proyectar el ahorro económico de optimizar el texto antes de enviarlo a un LLM.

### Criterios de aceptación cubiertos

1. **Carga flexible de datos**: Modo A (un solo archivo `.xlsx`) o Modo B (lote completo dentro de una carpeta), con selección de idioma de origen (`ES` / `EN`).
2. **Optimización de tokens opcional (`opt_tokens`)**: traducción de español a inglés antes de tokenizar, con `fallback` seguro si la API externa no responde.
3. **Auditoría económica**: cálculo de tokens con `tiktoken`, proyección de costos a 10.000 reseñas/día y exportación de reportes limpios en Excel y JSON.

---

## 2. Arquitectura del proyecto

| Archivo | Responsabilidad |
|---|---|
| `generate_excel_es.py` | Generador sintético de reseñas en español (Faker `es_ES`). Soporta modo interactivo por consola y modo automatizado para la web (archivo único o lote de archivos en una carpeta). |
| `generate_excel_en.py` | Igual que el anterior, pero en inglés (Faker `en_US`). |
| `reader_excel.py` | Núcleo del pipeline: normalización de columnas, limpieza de reseñas "con basura", traducción/optimización, clasificación técnica (`error_type` / `component`), conteo de tokens y generación de reportes Excel + JSON. Expone tanto una versión interactiva por consola (`ejecutar_sistema`) como una versión sin inputs pensada para Streamlit (`procesar_desde_web_streamlit`). |
| `frontend.py` | Interfaz web en Streamlit con 3 pestañas: generador en español, generador en inglés y el pipeline de análisis (HU-012), que reutiliza directamente las funciones de los tres scripts anteriores. |

La idea central es que **el frontend no reimplementa lógica**: solo importa y llama a las funciones ya existentes (`generar_desde_web_es`, `generar_desde_web_en`, `procesar_desde_web_streamlit`), de modo que el comportamiento en consola y en web sea idéntico.

---

## 3. Tecnologías usadas

- **Python 3** como lenguaje principal.
- **Pandas** y **NumPy** para manipulación de datos y limpieza de columnas.
- **Faker** para la generación sintética de nombres, ciudades y datos de clientes.
- **Openpyxl** para leer/escribir `.xlsx` y ajustar automáticamente el ancho de columnas.
- **tiktoken** (`o200k_base`) para el conteo real de tokens en las reseñas.
- **Requests** para la llamada opcional al endpoint externo de traducción `/api/analyze`.
- **JSON / diccionarios y listas en memoria** para las plantillas de reseñas, traducciones y clasificación técnica (sin depender de servicios externos salvo cuando `opt_tokens` no encuentra coincidencia local).
- **Streamlit** como capa de interfaz web sobre los scripts de consola ya existentes.

---

## 4. Decisiones técnicas

### 4.1 Ingesta flexible (Criterio 1)
Se usó `pathlib` para resolver de forma uniforme tanto un archivo único (Modo A) como una carpeta completa (Modo B), filtrando archivos temporales de Excel (`~$...`) y validando que la extensión sea `.xlsx`. Para evitar errores por nombres de columnas inconsistentes, se normalizan encabezados a minúsculas antes de buscar `reseña` o `review`, según el idioma seleccionado.

### 4.2 Optimización y traducción (Criterio 2)
La bandera `opt_tokens` solo está disponible cuando el idioma de origen es español. La estrategia es híbrida:
1. Se busca primero coincidencia exacta o por prefijo en un mapeo local de frases conocidas (`TRADUCCIONES_ES_EN`).
2. Si no hay coincidencia, se llama a `http://localhost:8000/api/analyze`.
3. Si la API falla o no está disponible, el sistema conserva el texto original en español (nunca detiene el pipeline).

Esto se decidió así para minimizar la latencia (la mayoría de las reseñas son plantillas conocidas) y para que el sistema sea resiliente ante caídas de un servicio externo.

### 4.3 Tokenización y proyección económica (Criterio 3)
Se usa `tiktoken` con el codificador `o200k_base` para contar tokens tanto en la versión "directa" (texto original) como en la "optimizada" (traducida/limpia). Con ese conteo se proyecta matemáticamente el costo a 10.000 reseñas/día a una tarifa de referencia de $2.50 USD por millón de tokens, comparando el costo directo vs. el optimizado y el ahorro estimado diario/mensual. La salida limpia se fuerza a estructuras JSON tipadas (`reporte_impacto_<nombre>.json`) para evitar romper esquemas de archivos planos, y las columnas de los `.xlsx` de salida se autoajustan en ancho con `openpyxl`.

### 4.4 Integración con el frontend (Streamlit)
`frontend.py` no reescribe lógica de negocio: importa los tres scripts como módulos y llama directamente a sus funciones "web-friendly" (`generar_desde_web_es`, `generar_desde_web_en`, `procesar_desde_web_streamlit`), que son versiones de las funciones de consola sin `input()`. Esto asegura que cualquier corrección hecha en los scripts base se refleje automáticamente en la web sin duplicar mantenimiento.

En Modo B (lote), el pipeline además consolida automáticamente todos los archivos procesados en un `REPORTE_MAESTRO_CONSOLIDADO.xlsx` y su respectivo JSON con la auditoría económica agregada de todo el lote.

---

## 5. Requisitos e instalación

Instala las dependencias desde `requirements.txt`:

```bash
pip install -r requirements.txt
```

Contenido de `requirements.txt`:

```
faker
pandas
tiktoken
numpy
openpyxl
requests
streamlit
```

> Si se activa la optimización de tokens (`opt_tokens`) y el texto no coincide con el mapeo local, el script intentará llamar a `http://localhost:8000/api/analyze` (debe recibir `{"text": "..."}` y responder `{"translated_text": "..."}`). Si esa API no está disponible, el sistema no falla: conserva el texto original en español.

---

## 6. Cómo ejecutar

### 6.1 Interfaz web (Streamlit) — recomendado

```bash
streamlit run frontend.py
```

Al abrirse en el navegador, Streamlit puede pedir un email de registro la primera vez: solo presiona **Enter** dejando el campo vacío para continuar. La app se organiza en 3 pestañas:

1. **🇪🇸 Generador Español (>50k)**: crea reseñas sintéticas en español; permite generar un archivo único o un lote de varios archivos dentro de una carpeta nueva.
2. **🇬🇧 Generador Inglés (>50k)**: igual que el anterior, en inglés.
3. **🧠 Pipeline Análisis Avanzado (HU-012)**: toma la ruta de un archivo o carpeta generada en las pestañas anteriores, permite activar `opt_tokens` (solo para español), ejecuta la limpieza/clasificación y muestra las métricas económicas y la vista previa del reporte maestro.

### 6.2 Modo consola (independiente)

Cada script también puede ejecutarse por separado:

```bash
python generate_excel_es.py     # Genera un Excel de reseñas en español
python generate_excel_en.py     # Genera un Excel de reseñas en inglés
python reader_excel.py          # Ejecuta el pipeline interactivo de limpieza y auditoría
```

`reader_excel.py` en modo consola pedirá, en orden:

1. Idioma de los archivos (`ES` / `EN`).
2. (Solo si es `ES`) Si se activa la optimización de tokens (`S` / `N`).
3. Modo de carga (`A` archivo individual / `B` carpeta en lote).
4. Ruta del archivo o carpeta de origen.
5. Nombre de la carpeta de salida (Enter para usar `Resultados_Limpios`).
6. Nombre de archivo de salida por cada Excel procesado (Enter para usar `<nombre>_limpio`).

### Salidas generadas

- `Resultados_Limpios/<archivo>_limpio.xlsx` — Excel limpio y clasificado por archivo.
- `Resultados_Limpios/reporte_impacto_<archivo>.json` — auditoría de costos por archivo.
- *(Solo Modo B)* `Resultados_Limpios/REPORTE_MAESTRO_CONSOLIDADO.xlsx` y su JSON maestro consolidado.

---

## 7. Ejemplo de JSON de auditoría económica

```json
{
  "metrica_volumen": "10k_resenas_dia",
  "costo_tarifa_m_tokens_usd": 2.50,
  "analisis_procesamiento": {
    "directo_con_basura": {
      "tokens_estimados_dia": 45000,
      "costo_estimado_usd_dia": 112.50
    },
    "optimizado_limpio": {
      "tokens_estimados_dia": 21000,
      "costo_estimado_usd_dia": 52.50
    }
  },
  "impacto_financiero_ahorro": {
    "ahorro_diario_usd": 60.00,
    "ahorro_mensual_estimado_usd": 1800.00,
    "eficiencia_datos_porcentaje": 53.33
  }
}
```

---

## 8. Notas del equipo

Este ejercicio fue desarrollado en equipo, avanzando de forma incremental por cada Criterio de Aceptación y verificando la estabilidad de cada fase antes de continuar con la siguiente. El pipeline principal de limpieza, clasificación y auditoría económica (`reader_excel.py`) fue construido colaborativamente; sobre esa base se integró la capa de interfaz web (`frontend.py`) para exponer los tres scripts como una consola unificada.

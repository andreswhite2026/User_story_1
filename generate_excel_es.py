import random
import time
from pathlib import Path
import pandas as pd
from faker import Faker


def generar_resenas_excel_es(num_filas=50000):
    print("==================================================")
    print("   GENERADOR DE DATOS DE PRUEBA EN ESPAÑOL (50K)  ")
    print("==================================================")

    # Preguntar al usuario el nombre del archivo
    nombre_entrada = (
        input(
            "Ingrese el nombre para el archivo (Presione Enter para usar nombre automático):\n> "
        )
        .strip()
    )

    if not nombre_entrada:
        # Genera un nombre único usando la fecha y hora: product_reviews_20260729_084512.xlsx
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        nombre_archivo = f"resenas_productos_{timestamp}.xlsx"
    else:
        # Asegura que tenga la extensión correcta
        if not nombre_entrada.endswith(".xlsx"):
            nombre_entrada += ".xlsx"
        nombre_archivo = nombre_entrada

    print(f"\n[+] Generando {num_filas} registros ficticios en español...")

    # Inicializar Faker en español (Nombres de personas y ciudades hispanas)
    fake = Faker("es_ES")

    # Lista de productos de ejemplo traducida al español
    productos_cat = [
        "Auriculares Inalámbricos Bluetooth",
        "Smartphone Pro Max 256GB",
        "Portátil Gaming 15.6''",
        "Reloj Inteligente Deportivo",
        "Cámara Digital 4K",
        "Teclado Mecánico RGB",
        "Monitor LED Curvo 27''",
        "Silla de Oficina Ergonómica",
        "Cafetera Espresso Automática",
        "Robot Aspirador",
        "Mochila Antirrobo Impermeable",
        "Altavoz Portátil Resistente al Agua",
    ]

    # Lista de plantillas de reseñas exactas en español
    plantillas_resenas = [
        "Excelente producto, superó mis expectativas.",
        "Llegó a tiempo y en perfecto estado. Muy recomendado.",
        "La calidad del material es acceptable por el precio.",
        "No me gustó la calidad del producto, esperaba más.",
        "Pésimo servicio de entrega, llegó dañado.",
        "Funciona muy bien, lo uso todos los días.",
        "Buen diseño y materiales, aunque podría ser un poco más barato.",
        "Cumple con lo promised en la descripción.",
    ]

    data = []

    # Generación eficiente de datos
    for _ in range(num_filas):
        id_cliente = fake.uuid4()[:8].upper()
        cliente = fake.name()
        ciudad = fake.city()
        producto = random.choice(productos_cat)

        if random.random() < 0.25:
            resena = ""
        else:
            # Combina una plantilla exacta con una frase aleatoria en español generada por Faker
            resena = f"{random.choice(plantillas_resenas)} {fake.sentence(nb_words=6)}"

        data.append(
            {
                "customer_id": id_cliente,
                "customer_name": cliente,
                "city": ciudad,
                "product": producto,
                "reseña": resena,  # Cabecera cambiada a español para que coincida con tus plantillas
            }
        )

    print("Creando DataFrame de Pandas...")
    df = pd.DataFrame(data)

    print(f"Guardando archivo Excel: {nombre_archivo}...")
    # Exportar a Excel
    df.to_excel(nombre_archivo, index=False, engine="openpyxl")

    print(
        f"¡Proceso completado con éxito! Archivo creado: {Path(nombre_archivo).resolve()}"
    )


if __name__ == "__main__":
    generar_resenas_excel_es()

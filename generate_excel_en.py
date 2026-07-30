import random
import time
from pathlib import Path
import pandas as pd
from faker import Faker


def generar_resenas_excel_en(num_filas=50000):
    print("==================================================")
    print("   GENERADOR DE DATOS DE PRUEBA EN INGLÉS (50K)   ")
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
        nombre_archivo = f"product_reviews_{timestamp}.xlsx"
    else:
        # Asegura que tenga la extensión correcta
        if not nombre_entrada.endswith(".xlsx"):
            nombre_entrada += ".xlsx"
        nombre_archivo = nombre_entrada

    print(f"\n[+] Generating {num_filas} fake records in English...")

    # Inicializar Faker en inglés
    fake = Faker("en_US")

    # Lista de productos de ejemplo traducida al inglés
    productos_cat = [
        "Wireless Bluetooth Headphones",
        "Smartphone Pro Max 256GB",
        "Gaming Laptop 15.6''",
        "Sport Smartwatch",
        "4K Digital Camera",
        "RGB Mechanical Keyboard",
        "Curved 27'' LED Monitor",
        "Ergonomic Office Chair",
        "Automatic Espresso Machine",
        "Robot Vacuum Cleaner",
        "Waterproof Anti-theft Backpack",
        "Waterproof Portable Speaker",
    ]

    # Lista de plantillas de reseñas exactas en inglés
    plantillas_resenas = [
        "Excellent product, exceeded my expectations.",
        "Arrived on time and in perfect condition. Highly recommended.",
        "The quality of the material is acceptable for the price.",
        "I didn't like the quality of the product, I expected more.",
        "Terrible delivery service, arrived damaged.",
        "Works very well, I use it every day.",
        "Good design and materials, although it could be a bit cheaper.",
        "Delivers what is promised in the description.",
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
            resena = f"{random.choice(plantillas_resenas)} {fake.sentence(nb_words=6)}"

        data.append(
            {
                "customer_id": id_cliente,
                "customer_name": cliente,
                "city": ciudad,
                "product": producto,
                "review": resena,
            }
        )

    print("Creating Pandas DataFrame...")
    df = pd.DataFrame(data)

    print(f"Saving Excel file: {nombre_archivo}...")
    # Exportar a Excel
    df.to_excel(nombre_archivo, index=False, engine="openpyxl")

    print(
        f"¡Process completed successfully! File created: {Path(nombre_archivo).resolve()}"
    )


if __name__ == "__main__":
    generar_resenas_excel_en()

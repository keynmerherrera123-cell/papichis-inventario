import sqlite3

# Nos conectamos a tu base de datos
conn = sqlite3.connect('inventario_papichis.db')
cursor = conn.cursor()

try:
    # Agregamos la columna de precio de venta (por defecto en 0)
    cursor.execute("ALTER TABLE productos ADD COLUMN precio_venta REAL DEFAULT 0.0")
    # Agregamos la columna de precio de compra (por defecto en 0)
    cursor.execute("ALTER TABLE productos ADD COLUMN precio_compra REAL DEFAULT 0.0")
    
    print("¡Éxito! Las columnas de precios se agregaron a la base de datos.")
except sqlite3.OperationalError as e:
    print(f"Aviso: {e}. (Probablemente las columnas ya existen).")

conn.commit()
conn.close()
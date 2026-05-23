import sqlite3
from datetime import datetime

def iniciar_sistema_papichis():
    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    
    # Creamos la tabla
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            categoria TEXT NOT NULL,
            subcategoria TEXT,
            nombre_producto TEXT NOT NULL UNIQUE,
            cantidad_actual INTEGER NOT NULL,
            cantidad_minima_alerta INTEGER DEFAULT 5,
            ultima_actualizacion TEXT
        )
    ''')
    conn.commit()
    conn.close()

def cargar_inventario_inicial():
    # El inventario EXACTO de Papichis
    inventario = [
        # --- PLÁSTICOS ---
        ('Plástico', 'Consumibles', 'Servilletas', 10),
        ('Plástico', 'Consumibles', 'Bolsas de papel mediana', 5),
        ('Plástico', 'Consumibles', 'Bolsas de papel grandes', 4),
        ('Plástico', 'Consumibles', 'Vasos plástico', 6),
        ('Plástico', 'Consumibles', 'Vasos nestea', 2),
        ('Plástico', 'Consumibles', 'Vaso Nescafé', 1),
        ('Plástico', 'Consumibles', 'Papel antigraso', 4),
        ('Plástico', 'Consumibles', 'Bolsas plástica medianas', 7),
        ('Plástico', 'Consumibles', 'Bolsas plástica grandes', 5),
        ('Plástico', 'Consumibles', 'Bolsas negras', 4),
        ('Plástico', 'Consumibles', 'Cucharas', 4),
        ('Plástico', 'Consumibles', 'Tenedores', 2),
        ('Plástico', 'Consumibles', 'Pitillos', 5),
        ('Plástico', 'Consumibles', 'Cajas de pasapalos grandes', 18),
        ('Plástico', 'Consumibles', 'Cajas de pasapalos pequeñas', 50),
        ('Plástico', 'Consumibles', 'CT2', 3),
        ('Plástico', 'Consumibles', 'CT1', 0),

        # --- BEBIDAS: REFRESCO ---
        ('Bebidas', 'Refresco', 'Refresco 1lt', 127),
        ('Bebidas', 'Refresco', 'Refresco 1.5lt', 45),
        ('Bebidas', 'Refresco', 'Refresco 2lt', 30),
        ('Bebidas', 'Refresco', 'Refresco 350', 413),
        ('Bebidas', 'Refresco', 'Refresco de lata', 0),
        ('Bebidas', 'Refresco', 'Refresco bombonita', 0),
        ('Bebidas', 'Refresco', 'Refresco retornable', 41),

        # --- BEBIDAS: MALTA ---
        ('Bebidas', 'Malta', 'Malta retornable', 63),
        ('Bebidas', 'Malta', 'Malta desechable', 48),
        ('Bebidas', 'Malta', 'Malta de lata', 48),
        ('Bebidas', 'Malta', 'Malta 1.5lt', 1),

        # --- BEBIDAS: AGUA ---
        ('Bebidas', 'Agua', 'Agua 600ml', 38),
        ('Bebidas', 'Agua', 'Agua 355ml', 14),
        ('Bebidas', 'Agua', 'Agua gasificada 355ml', 1),
        ('Bebidas', 'Agua', 'Agua Saborizada 355ml', 1),
        ('Bebidas', 'Agua', 'Agua 1.5lt', 11),
        ('Bebidas', 'Agua', 'Agua gasificada 1.5 lt', 6),
        ('Bebidas', 'Agua', 'Agua saborizada 1.5lt', 2),

        # --- BEBIDAS: JUGOS Y ENERGIZANTES ---
        ('Bebidas', 'Jugos', 'Jugo del valle 500', 0),
        ('Bebidas', 'Jugos', 'Jugo del valle 1.5lt', 6),
        ('Bebidas', 'Jugos', 'Yukery de botella', 48),
        ('Bebidas', 'Jugos', 'Yuky pak', 56),
        ('Bebidas', 'Jugos', 'Cifrut', 0),
        ('Bebidas', 'Jugos', 'Yukery 1.5lt', 1),
        ('Bebidas', 'Energizantes', 'Speed Max', 27),
        ('Bebidas', 'Energizantes', 'Gatorade', 11),
        ('Bebidas', 'Energizantes', 'Power', 4)
    ]

    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for cat, subcat, nombre, cant in inventario:
        try:
            cursor.execute('''
                INSERT INTO productos (categoria, subcategoria, nombre_producto, cantidad_actual, ultima_actualizacion)
                VALUES (?, ?, ?, ?, ?)
            ''', (cat, subcat, nombre, cant, fecha_actual))
        except sqlite3.IntegrityError:
            pass

    conn.commit()
    conn.close()
    print("¡Inventario oficial de Papichis cargado con éxito!")

# Ejecutar
iniciar_sistema_papichis()
cargar_inventario_inicial()
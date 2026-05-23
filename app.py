from flask import Flask, render_template_string, request, redirect, url_for, session, send_file
import sqlite3
import io
from datetime import datetime

# Librerías de OpenPyXL para ponerle estilo al Excel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

app = Flask(__name__)
# La clave secreta es necesaria para que las sesiones funcionen de forma segura
app.secret_key = 'clave_secreta_super_segura_papichis'


# --- ZONA DE TRABAJADORES (Vista Móvil) ---

@app.route('/')
def mostrar_inventario():
    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, nombre_producto, cantidad_actual, categoria FROM productos")
    datos = cursor.fetchall()
    conn.close()

    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Control Papichis</title>
        <style>
            body { font-family: Arial; padding: 15px; background: #f4f4f4; margin: 0; }
            h2 { color: #333; text-align: center; }
            .tarjeta { background: white; padding: 15px; margin-bottom: 12px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); display: flex; justify-content: space-between; align-items: center;}
            .info { flex-grow: 1; }
            .formulario-edicion { display: flex; gap: 8px; }
            input[type="number"] { width: 60px; padding: 8px; font-size: 16px; text-align: center; border: 1px solid #ccc; border-radius: 4px; }
            button { background: #007bff; color: white; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; cursor: pointer; }
        </style>
    </head>
    <body>
        <h2>📱 Inventario Papichis</h2>
        {% for producto in datos %}
        <div class="tarjeta">
            <div class="info">
                <strong>{{ producto[1] }}</strong><br>
                <small style="color: gray;">{{ producto[3] }}</small>
            </div>
            <form action="/actualizar" method="POST" class="formulario-edicion">
                <input type="hidden" name="id_producto" value="{{ producto[0] }}">
                <input type="number" name="nueva_cantidad" value="{{ producto[2] }}" min="0">
                <button type="submit">Guardar</button>
            </form>
        </div>
        {% endfor %}
        <br>
        <div style="text-align: center;">
            <a href="/admin" style="color: gray; text-decoration: none; font-size: 12px;">Acceso Administrador</a>
        </div>
    </body>
    </html>
    """
    return render_template_string(html, datos=datos)

@app.route('/actualizar', methods=['POST'])
def actualizar_inventario():
    producto_id = request.form['id_producto']
    nueva_cantidad = request.form['nueva_cantidad']
    fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE productos SET cantidad_actual = ?, ultima_actualizacion = ? WHERE id = ?
    ''', (nueva_cantidad, fecha_actual, producto_id))
    conn.commit()
    conn.close()
    return redirect(url_for('mostrar_inventario'))


# --- ZONA DE ADMINISTRACIÓN (Dashboard y Seguridad) ---

@app.route('/admin', methods=['GET', 'POST'])
def admin_login():
    if request.method == 'POST':
        clave = request.form.get('clave')
        if clave == 'papichis2026':
            session['admin_autenticado'] = True  # Guardamos el estado en la sesión
            return redirect(url_for('admin_dashboard'))
        else:
            return "<h2 style='color:red; text-align:center;'>Contraseña incorrecta. <a href='/admin'>Volver</a></h2>"
    
    # Si es GET, mostramos el login
    html = """
    <!DOCTYPE html>
    <html>
    <head><title>Login Papichis</title></head>
    <body style="font-family: Arial; text-align: center; padding-top: 100px; background: #333; color: white;">
        <h2>Acceso Exclusivo Administrador</h2>
        <form action="/admin" method="POST">
            <input type="password" name="clave" placeholder="Ingresa la contraseña" required style="padding: 10px; font-size: 16px;">
            <button type="submit" style="padding: 10px 20px; font-size: 16px; cursor: pointer;">Entrar</button>
        </form>
    </body>
    </html>
    """
    return html

@app.route('/admin/dashboard')
def admin_dashboard():
    # Verificación de seguridad por sesión
    if not session.get('admin_autenticado'):
        return redirect(url_for('admin_login'))

    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    cursor.execute("SELECT nombre_producto, categoria, cantidad_actual, ultima_actualizacion FROM productos ORDER BY cantidad_actual ASC")
    datos = cursor.fetchall()
    conn.close()

    html = """
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="UTF-8">
        <title>Dashboard Papichis</title>
        <style>
            body { font-family: Arial; padding: 20px; background: #fff; }
            .encabezado { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #333; padding-bottom: 10px; }
            h2 { margin: 0; color: #333; }
            .btn-excel { background-color: #1f497d; color: white; text-decoration: none; padding: 10px 20px; border-radius: 5px; font-weight: bold; font-size: 14px; }
            .btn-excel:hover { background-color: #102a4a; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; }
            th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
            th { background-color: #f4f4f4; }
            .agotado { background-color: #ffe6e6; color: #cc0000; font-weight: bold; }
        </style>
    </head>
    <body>
        <div class="encabezado">
            <h2>📊 Panel de Reportes - Papichis</h2>
            <a href="/descargar_excel" class="btn-excel">📥 Descargar Reporte (Excel)</a>
        </div>
        <p>Vista de solo lectura para auditoría. Los productos en rojo necesitan reposición urgente.</p>
        
        <table>
            <tr>
                <th>Producto</th>
                <th>Categoría</th>
                <th>Stock Actual</th>
                <th>Última Revisión</th>
            </tr>
            {% for producto in datos %}
            <tr class="{% if producto[2] <= 5 %}agotado{% endif %}">
                <td>{{ producto[0] }}</td>
                <td>{{ producto[1] }}</td>
                <td>{{ producto[2] }}</td>
                <td>{{ producto[3] or 'No revisado aún' }}</td>
            </tr>
            {% endfor %}
        </table>
        <br>
        <a href="/admin/logout" style="color: red; text-decoration: none; font-size: 14px;">Cerrar Sesión Administrador</a>
    </body>
    </html>
    """
    return render_template_string(html, datos=datos)


# --- RUTA DE GENERACIÓN EN TIEMPO REAL DEL EXCEL ---

@app.route('/descargar_excel')
def descargar_excel():
    if not session.get('admin_autenticado'):
        return redirect(url_for('admin_login'))

    # 1. Leer los datos más frescos de la BD
    conn = sqlite3.connect('inventario_papichis.db')
    cursor = conn.cursor()
    cursor.execute("SELECT id, categoria, subcategoria, nombre_producto, cantidad_actual, ultima_actualizacion FROM productos")
    productos_bd = cursor.fetchall()
    conn.close()

    # 2. Inicializar el libro de Excel en memoria RAM (No creamos archivos basura en el disco)
    wb = openpyxl.Workbook()
    
    ws_summary = wb.active
    ws_summary.title = "Resumen Ejecutivo"
    ws_data = wb.create_sheet(title="Inventario Detallado")
    
    ws_summary.views.sheetView[0].showGridLines = True
    ws_data.views.sheetView[0].showGridLines = True

    # Estilos del Diseño Corporativo
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    fill_sub_header = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    fill_alert = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
    
    font_title = Font(name="Calibri", size=16, bold=True, color="1F497D")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    font_alert = Font(name="Calibri", size=11, bold=True, color="C00000")

    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'), top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    border_total = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))

    # --- PESTAÑA 2: INVENTARIO DETALLADO ---
    ws_data.append([])
    ws_data.append(["REPORTE DE INVENTARIO DETALLADO - PAPICHIS"])
    ws_data.cell(row=2, column=1).font = font_title
    ws_data.append([])
    
    headers = ["ID", "Categoría", "Subcategoría", "Producto", "Stock Actual", "Estado Alerta", "Última Revisión"]
    ws_data.append(headers)
    
    for col_num, h in enumerate(headers, 1):
        c = ws_data.cell(row=4, column=col_num)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal='center' if col_num in [1,5,6,7] else 'left', vertical='center')

    start_row = 5
    for idx, fila in enumerate(productos_bd, 1):
        r_num = start_row + idx - 1
        _, cat, subcat, name, qty, last_rev = fila
        
        # Truco: Calculamos los datos estáticos directamente con Python
        estado_alerta = "CRÍTICO" if qty <= 5 else "OK"
        revision_texto = last_rev if last_rev else datetime.now().strftime("%Y-%m-%d %H:%M")
        
        ws_data.append([idx, cat, subcat, name, qty, estado_alerta, revision_texto])
        
        for col_idx in range(1, 8):
            cell = ws_data.cell(row=r_num, column=col_idx)
            cell.font = font_regular
            cell.border = border_thin
            if r_num % 2 == 0:
                cell.fill = fill_zebra
            
            cell.alignment = Alignment(horizontal='center' if col_idx in [1,5,6,7] else 'left', vertical='center')
            if col_idx == 5:
                cell.number_format = '#,##0'

    tot_row = start_row + len(productos_bd)
    ws_data.cell(row=tot_row, column=4, value="Total Productos en Stock").font = font_bold
    ws_data.cell(row=tot_row, column=4).alignment = Alignment(horizontal='right')
    
    # Aquí calculamos la suma total en Python para saltarnos el bloqueo de la Vista Protegida de Excel
    suma_total_stock = sum(f[4] for f in productos_bd)
    items_criticos = sum(1 for f in productos_bd if f[4] <= 5)
    
    total_cell = ws_data.cell(row=tot_row, column=5, value=suma_total_stock)
    total_cell.font = font_bold
    total_cell.alignment = Alignment(horizontal='center')
    total_cell.border = border_total

    # Formato condicional para la columna de alertas
    rule = CellIsRule(operator='equal', formula=['"CRÍTICO"'], fill=fill_alert, font=font_alert)
    ws_data.conditional_formatting.add(f"F5:F{tot_row-1}", rule)

    # --- PESTAÑA 1: RESUMEN EJECUTIVO ---
    ws_summary.append([])
    ws_summary.append(["SISTEMA DE CONTROL DE INVENTARIO - PAPICHIS"])
    ws_summary.cell(row=2, column=2).font = font_title
    ws_summary.append([])
    ws_summary.append(["", "PANEL DE CONTROL GENERAL (KPIs)"])
    ws_summary.cell(row=4, column=2).font = Font(name="Calibri", size=11, bold=True)
    ws_summary.cell(row=4, column=2).fill = PatternFill(start_color="EAEAEA", end_color="EAEAEA", fill_type="solid")
    ws_summary.append([])
    
    ws_summary.cell(row=6, column=2, value="Métrica de Control").font = font_bold
    ws_summary.cell(row=6, column=2).fill = fill_sub_header
    ws_summary.cell(row=6, column=2).border = border_thin
    ws_summary.cell(row=6, column=3, value="Valor").font = font_bold
    ws_summary.cell(row=6, column=3).fill = fill_sub_header
    ws_summary.cell(row=6, column=3).border = border_thin
    ws_summary.cell(row=6, column=3).alignment = Alignment(horizontal='center')

    metrics = [
        ("Total de Items Registrados", len(productos_bd)),
        ("Total Unidades en Almacén", suma_total_stock),
        ("Productos con Stock Crítico (<=5)", items_criticos)
    ]

    for r_idx, (m_name, m_val) in enumerate(metrics, 7):
        c_name = ws_summary.cell(row=r_idx, column=2, value=m_name)
        c_name.font = font_regular
        c_name.border = border_thin
        
        c_val = ws_summary.cell(row=r_idx, column=3, value=m_val)
        c_val.font = font_bold
        c_val.border = border_thin
        c_val.alignment = Alignment(horizontal='center')
        
        if r_idx == 9 and items_criticos > 0:
            c_name.fill = fill_alert
            c_name.font = font_alert
            c_val.fill = fill_alert
            c_val.font = font_alert

    # Auto-ajustar columnas
    for ws in [ws_summary, ws_data]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)
            
    ws_summary.column_dimensions['A'].width = 3
    ws_summary.column_dimensions['B'].width = 35
    ws_summary.column_dimensions['C'].width = 15

    # Guardar en memoria de intercambio y enviar el archivo al navegador
    excel_stream = io.BytesIO()
    wb.save(excel_stream)
    excel_stream.seek(0)
    
    return send_file(
        excel_stream,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=f'Reporte_Inventario_Papichis_{datetime.now().strftime("%d_%m_%Y")}.xlsx'
    )

@app.route('/admin/logout')
def admin_logout():
    session.pop('admin_autenticado', None)
    return redirect(url_for('mostrar_inventario'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
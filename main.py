from flask import Flask, request, send_file
import openpyxl
from io import BytesIO
from flask_cors import CORS

app = Flask(__name__)
CORS(app) # Esto permite que tu HTML de Claude se conecte al servidor

@app.route('/generar-protocolo', methods=['POST'])
def generar_protocolo():
    try:
        # 1. Recibir los datos desde tu página web
        datos = request.json
        
        # 2. Seleccionar la plantilla según lo que pida la web
        # Por defecto usaremos el de Uniones si no se especifica
        nombre_plantilla = datos.get('plantilla', 'GM-RU-HDPE-INF-RAI-001.xlsx')
        
        # 3. Cargar el libro de Excel original
        wb = openpyxl.load_workbook(nombre_plantilla)
        ws = wb.active

        # --- MAPEO DE DATOS (Encabezado) ---
        # Basado en tus archivos de Membrantec
        ws['Y6'] = datos.get('fecha', '')            # Celda de Fecha
        ws['U6'] = datos.get('numero_protocolo', '') # Celda N° Protocolo
        ws['A10'] = datos.get('tag_piscina', '')     # Celda TAG Piscina/Proyecto

        # --- MAPEO DE TABLA (Soldaduras/Uniones) ---
        # En tus archivos, los datos suelen empezar en la fila 12
        fila_inicio = 12
        uniones = datos.get('lista_datos', [])

        for i, item in enumerate(uniones):
            fila_actual = fila_inicio + i
            # Columna B: Unión N° / Columna C: Distancia / Columna E: Operador
            ws[f'B{fila_actual}'] = item.get('id', '')
            ws[f'C{fila_actual}'] = item.get('distancia', '')
            ws[f'E{fila_actual}'] = item.get('operador', '')
            ws[f'F{fila_actual}'] = item.get('maquina', '')
            
            # Resultado: Si está aprobado, pone una X en la columna de 'Línea Completa SI'
            if item.get('estado') == 'Aprobado':
                ws[f'I{fila_actual}'] = "X"
            else:
                ws[f'J{fila_actual}'] = "X"

        # 4. Preparar el archivo para descarga
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=f"Protocolo_{datos.get('numero_protocolo', 'Generado')}.xlsx"
        )

    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == '__main__':
    # Render usa la variable de entorno PORT
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
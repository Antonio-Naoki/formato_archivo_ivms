from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Función para procesar el archivo
def procesar_archivo(df, nombre_archivo):
    columnas_a_eliminar = ["Data Source", "Handling Type", "Temperature", "Abnormal", "Attendance Check Point"]
    df = df.drop(columns=columnas_a_eliminar, errors='ignore')
    
    df = df.rename(columns={
        "Person ID": "ID Persona",
        "Name": "Nombre",
        "Time": "Hora",
        "Attendance Status": "Estado de Asistencia",
        "Custom Name": "Tipo de Evento"
    })
    
    if "Hora" in df.columns:
        df["Hora"] = pd.to_datetime(df["Hora"], errors='coerce')
        df = df.sort_values(by="Hora", ascending=True) 
        df["Hora"] = df["Hora"].dt.strftime('%I:%M %p')
    
    if "Tipo de Evento" in df.columns:
        df["Tipo de Evento"] = df["Tipo de Evento"].replace({"Entrada": "Entrada", "Salida": "Salida"})
    
    nombre_salida = os.path.join(PROCESSED_FOLDER, f"{nombre_archivo}_formateado.xlsx")
    
    with pd.ExcelWriter(nombre_salida, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Datos")
        workbook = writer.book
        worksheet = writer.sheets["Datos"]
        
        # Estilos para encabezados
        header_fill = PatternFill(start_color="3CB371", end_color="3CB371", fill_type="solid")  # Verde profesional
        header_font = Font(color="FFFFFF", bold=True, size=12)
        header_alignment = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # Estilos para filas (Fondo blanco, bordes definidos)
        for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column):
            for cell in row:
                cell.fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="left", vertical="center")
        
        # Ajustar el tamaño de las columnas automáticamente
        for col in worksheet.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[col_letter].width = adjusted_width
    
    return nombre_salida

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/procesar', methods=['POST'])
def procesar_archivo_route():
    if 'file' not in request.files:
        return "No se ha subido ningún archivo"
    
    archivo = request.files['file']
    if archivo.filename == '':
        return "No se ha seleccionado ningún archivo"
    
    nombre_archivo = os.path.splitext(archivo.filename)[0]
    ruta_archivo = os.path.join(UPLOAD_FOLDER, archivo.filename)
    archivo.save(ruta_archivo)
    
    try:
        if archivo.filename.endswith('.csv'):
            df = pd.read_csv(ruta_archivo)
        elif archivo.filename.endswith('.xlsx'):
            df = pd.read_excel(ruta_archivo, engine='openpyxl')
        else:
            return "Formato de archivo no soportado. Sube un archivo .csv o .xlsx."
        
        archivo_salida = procesar_archivo(df, nombre_archivo)
        return send_file(archivo_salida, as_attachment=True)
    except Exception as e:
        return f"Error procesando el archivo: {str(e)}"

if __name__ == '__main__':
    app.run(debug=True)


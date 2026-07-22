import os
import smtplib
import traceback
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from flask import Flask, render_template, request, jsonify

app = Flask(__name__)

# ==========================================
# CONFIGURACIÓN Y DESTINATARIOS
# ==========================================
DESTINATARIOS_DICT = {
    "Antioquia - Costa": ["correo1@empresa.com"],
    "Centro - Ori": ["correo2@empresa.com"],
    "Cuentas Claves": ["correo3@empresa.com"],
    "Eje - Occidente": ["correo4@empresa.com"],
    "TODOS": ["equipo_general@empresa.com"]
}

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "tu_correo@empresa.com")
SMTP_PASS = os.getenv("SMTP_PASS", "tu_contraseña")

MESES_ESPANOL = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# MAPEO EXACTO DE SECCIONES A TUS COLUMNAS EN EXCEL
MAPEO_SECCIONES = {
    "Competencia": {"meta": "Metas Competencia", "real": "SP. Competencia", "efe": "Efe Competencia"},
    "Leads": {"meta": "Metas Leads", "real": "Leads.Leads", "efe": "Efe Leads"},
    "Prospecciones": {"meta": "Metas Prospecciones", "real": "Prospecciones", "efe": "Efe Prospecciones"},
    "Food": {"meta": "Metas Food", "real": "SP. Food", "efe": "Efe Food"},
    "Ingredientes": {"meta": "Metas Ingredientes", "real": "SP. Ingredientes", "efe": "Efe Ingredientes"},
    "Puntos Claves": {"meta": "Puntos Programados Claves", "real": "Poc Visitados Claves", "efe": "Efe Puntos Claves"},
    "Puntos Normal": {"meta": "Puntos Programados Normal", "real": "Poc Visitados Normal", "efe": "Efe Puntos Normales"},
    "Logueos": {"meta": "Metas Logueos", "real": "Faltas Logueos", "efe": "Efe Logueos"},
    "Deslogueos": {"meta": "Metas Deslogueos", "real": "Faltas Deslogueos", "efe": "Efe Deslogueos"},
    "Sell Out": {"meta": "Meta Sell Out", "real": "Ejecutado Sell Out", "efe": "efe Sell Out"},
    "Sell In": {"meta": "Meta Sell In", "real": "Ejecutado Sell In", "efe": "Efe Sell In"},
    "Instalaciones": {"meta": "Meta instalaciones", "real": "Ejecutado Instalaciones", "efe": "Efe instalaciones"},
    "Retiros": {"meta": "Meta Retiros", "real": "Ejecutado Retiros", "efe": "Efe Retiros"},
    "Netas": {"meta": "Meta Netas", "real": "Ejecutado Netas", "efe": "Efe Netas"}
}

def obtener_fecha_corte_actual():
    ahora = datetime.now()
    nombre_mes = MESES_ESPANOL.get(ahora.month, "")
    return f"{ahora.day} de {nombre_mes} – {ahora.year}"

def estilo_cumplimiento(porcentaje, meta_esperada):
    if porcentaje >= meta_esperada:
        return 'background-color: #C6EFCE; color: #006100; font-weight: bold;', '✔️'
    elif porcentaje >= (meta_esperada - 10):
        return 'background-color: #FFEB9C; color: #9C6500; font-weight: bold;', '⚠️'
    else:
        return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;', '❌'

def generar_html_reporte(df, secciones_seleccionadas, meta_esperada):
    # Limpiar espacios en los nombres de las columnas
    df.columns = df.columns.str.strip()

    # Detectar la columna del nombre
    col_persona = None
    for posible in ['NOMBRE COMPLETO', 'NOMBRE', 'DESARROLLADOR', 'CARGO']:
        if posible in df.columns:
            col_persona = posible
            break

    if not col_persona:
        return "<h3 style='color:red;'>No se encontró la columna 'NOMBRE COMPLETO' o 'DESARROLLADOR' en el Excel.</h3>"

    # Filtrar solo secciones que existan realmente en las columnas
    columnas_a_sumar = []
    secciones_validas = []

    for sec in secciones_seleccionadas:
        if sec in MAPEO_SECCIONES:
            meta_col = MAPEO_SECCIONES[sec]['meta'].strip()
            real_col = MAPEO_SECCIONES[sec]['real'].strip()
            
            if meta_col in df.columns and real_col in df.columns:
                df[meta_col] = pd.to_numeric(df[meta_col], errors='coerce').fillna(0)
                df[real_col] = pd.to_numeric(df[real_col], errors='coerce').fillna(0)
                columnas_a_sumar.extend([meta_col, real_col])
                secciones_validas.append((sec, meta_col, real_col))

    if not secciones_validas:
        return "<h3 style='color:orange;'>No se seleccionaron secciones válidas o no coinciden con el Excel.</h3>"

    # AGRUPAR DE MANERA ÚNICA POR PERSONA
    df_persona = df.groupby(col_persona)[columnas_a_sumar].sum().reset_index()

    # CONSTRUIR ENCABEZADOS DE LA TABLA HTML
    th_secciones = ""
    th_sub = ""
    for sec, _, _ in secciones_validas:
        th_secciones += f"<th colspan='3' class='head-sec'>{sec}</th>"
        th_sub += "<th class='head-sub'>Meta</th><th class='head-sub'>Real</th><th class='head-sub'>% Cumpl.</th>"

    html_tabla = f"""
    <table class="data-table">
        <thead>
            <tr>
                <th rowspan="2" class="head-main">Nombre Completo</th>
                {th_secciones}
            </tr>
            <tr>
                {th_sub}
            </tr>
        </thead>
        <tbody>
    """

    # CONSTRUIR FILAS POR CADA PERSONA
    for _, row in df_persona.iterrows():
        html_tabla += f"<tr><td style='text-align:left; font-weight:bold;'>{row[col_persona]}</td>"
        
        for sec, meta_col, real_col in secciones_validas:
            m = int(row[meta_col])
            r = int(row[real_col])
            pct = round((r / m * 100), 1) if m > 0 else 0
            est, ico = estilo_cumplimiento(pct, meta_esperada)

            html_tabla += f"<td>{m}</td><td>{r}</td><td style='{est}'>{pct:.0f}% {ico}</td>"
        
        html_tabla += "</tr>"

    html_tabla += "</tbody></table>"

    return f"""
    <html>
    <head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #333; }}
        .data-table {{ border-collapse: collapse; width: 100%; margin-top: 15px; font-size: 11px; }}
        .data-table th, .data-table td {{ border: 1px solid #b0c4de; padding: 6px 8px; text-align: center; white-space: nowrap; }}
        .head-main {{ background-color: #1F497D; color: white; vertical-align: middle; }}
        .head-sec {{ background-color: #2F5597; color: white; font-size: 12px; }}
        .head-sub {{ background-color: #DDEBF7; color: #000; font-weight: bold; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
    </style>
    </head>
    <body>
        <p>Buenos días Equipo,</p>
        <p>Comparto el consolidado individual al <b>{obtener_fecha_corte_actual()}</b>:</p>
        <p><i>Meta de cumplimiento esperada: <b>{meta_esperada}%</b></i></p>

        {html_tabla}

        <p style="margin-top:20px;">Cordialmente,<br><b>Automatización de Reportes</b></p>
    </body>
    </html>
    """

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get-filters', methods=['POST'])
def get_filters():
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "No se subió archivo"}), 400
    try:
        df = pd.read_excel(file, sheet_name="BASE")
        df.columns = df.columns.str.strip()
        meses = df['MES'].dropna().unique().tolist() if 'MES' in df.columns else []
        regionales = df['REGIONAL'].dropna().unique().tolist() if 'REGIONAL' in df.columns else []
        return jsonify({"meses": meses, "regionales": regionales})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/preview', methods=['POST'])
def preview():
    try:
        file = request.files.get('file')
        if not file:
            return "Por favor selecciona el archivo Excel desde tu equipo.", 400
        
        df = pd.read_excel(file, sheet_name="BASE")
        df.columns = df.columns.str.strip()

        mes = request.form.get('mes')
        regional = request.form.get('regional')
        meta_esperada = float(request.form.get('meta_esperada', 67))
        secciones_sel = request.form.getlist('columnas_sel')

        # Aplicar Filtros
        if mes and mes != 'TODOS' and 'MES' in df.columns:
            if pd.api.types.is_numeric_dtype(df['MES']):
                try: mes = float(mes)
                except ValueError: pass
            df = df[df['MES'] == mes]
            
        if regional and regional != 'TODOS' and 'REGIONAL' in df.columns:
            df = df[df['REGIONAL'] == regional]

        return generar_html_reporte(df, secciones_sel, meta_esperada)
    except Exception as e:
        return f"<div style='color:red; padding:15px;'><b>Error al procesar:</b><pre>{traceback.format_exc()}</pre></div>"

@app.route('/send', methods=['POST'])
def send():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({"message": "Falta el archivo Excel"}), 400

        df = pd.read_excel(file, sheet_name="BASE")
        df.columns = df.columns.str.strip()

        mes = request.form.get('mes')
        regional = request.form.get('regional')
        meta_esperada = float(request.form.get('meta_esperada', 67))
        secciones_sel = request.form.getlist('columnas_sel')

        if mes and mes != 'TODOS' and 'MES' in df.columns:
            if pd.api.types.is_numeric_dtype(df['MES']):
                try: mes = float(mes)
                except ValueError: pass
            df = df[df['MES'] == mes]
            
        if regional and regional != 'TODOS' and 'REGIONAL' in df.columns:
            df = df[df['REGIONAL'] == regional]

        html_body = generar_html_reporte(df, secciones_sel, meta_esperada)
        correos = DESTINATARIOS_DICT.get(regional, DESTINATARIOS_DICT.get('TODOS', []))

        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Reporte Consolidado por Persona - {obtener_fecha_corte_actual()}"
        msg['From'] = SMTP_USER
        msg['To'] = ", ".join(correos)
        msg.attach(MIMEText(html_body, 'html'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, correos, msg.as_string())

        return jsonify({"message": f"Correo enviado a: {', '.join(correos)}"})
    except Exception as e:
        return jsonify({"message": f"Error al enviar correo: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)

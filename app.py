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
# 1. DICCIONARIO DE DESTINATARIOS POR REGIONAL
# ==========================================
DESTINATARIOS_DICT = {
    "Antioquia - Costa": ["correo1@empresa.com", "correo2@empresa.com"],
    "Centro - Ori": ["correo3@empresa.com"],
    "Cuentas Claves": ["correo4@empresa.com"],
    "Eje - Occidente": ["correo5@empresa.com"],
    "TODOS": ["equipo_general@empresa.com"]
}

# CONFIGURACIÓN SMTP
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "tu_correo@empresa.com")
SMTP_PASS = os.getenv("SMTP_PASS", "tu_contraseña_o_token")

MESES_ESPANOL = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

# MAPEO DE INDICADORES A COLUMNAS EN TU EXCEL
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

def obtener_estilo_cumplimiento(porcentaje, meta_esperada):
    if porcentaje >= meta_esperada:
        return 'background-color: #C6EFCE; color: #006100; font-weight: bold;', '✔️'
    elif porcentaje >= (meta_esperada - 10):
        return 'background-color: #FFEB9C; color: #9C6500; font-weight: bold;', '⚠️'
    else:
        return 'background-color: #FFC7CE; color: #9C0006; font-weight: bold;', '❌'

def generar_html_reporte(df, secciones_seleccionadas, meta_esperada):
    fecha_corte = obtener_fecha_corte_actual()
    html_bloques = ""

    # Limpiar nombres de columnas eliminando espacios antes/después
    df.columns = df.columns.str.strip()

    col_nombre = 'NOMBRE COMPLETO' if 'NOMBRE COMPLETO' in df.columns else 'DESARROLLADOR'
    col_reg = 'REGIONAL' if 'REGIONAL' in df.columns else 'REGIONAL'

    for sec in secciones_seleccionadas:
        if sec not in MAPEO_SECCIONES:
            continue
        
        info = MAPEO_SECCIONES[sec]
        col_meta = info['meta'].strip()
        col_real = info['real'].strip()

        if col_meta not in df.columns or col_real not in df.columns:
            continue

        # Convertir a numérico por seguridad
        df[col_meta] = pd.to_numeric(df[col_meta], errors='coerce').fillna(0)
        df[col_real] = pd.to_numeric(df[col_real], errors='coerce').fillna(0)

        # 1. Total Indicador
        tot_m = int(df[col_meta].sum())
        tot_r = int(df[col_real].sum())
        pct_tot = round((tot_r / tot_m * 100), 1) if tot_m > 0 else 0
        est_tot, ico_tot = obtener_estilo_cumplimiento(pct_tot, meta_esperada)

        # 2. Agrupado por Regional
        filas_reg = ""
        if col_reg in df.columns:
            df_reg = df.groupby(col_reg)[[col_meta, col_real]].sum().reset_index()
            for _, r in df_reg.iterrows():
                m, real_val = int(r[col_meta]), int(r[col_real])
                pct = round((real_val / m * 100), 1) if m > 0 else 0
                est, ico = obtener_estilo_cumplimiento(pct, meta_esperada)
                filas_reg += f"""
                <tr>
                    <td style="text-align:left; background-color: #f9f9f9;"><b>{r[col_reg]}</b></td>
                    <td>{m}</td><td>{real_val}</td>
                    <td style="{est}">{pct:.0f}% {ico}</td>
                </tr>"""

        # 3. Agrupado por Persona
        filas_dev = ""
        if col_nombre in df.columns:
            df_dev = df.groupby(col_nombre)[[col_meta, col_real]].sum().reset_index().sort_values(by=col_real, ascending=False)
            for _, r in df_dev.iterrows():
                m, real_val = int(r[col_meta]), int(r[col_real])
                pct = round((real_val / m * 100), 1) if m > 0 else 0
                est, ico = obtener_estilo_cumplimiento(pct, meta_esperada)
                filas_dev += f"""
                <tr>
                    <td style="text-align:left;">{r[col_nombre]}</td>
                    <td>{m}</td><td>{real_val}</td>
                    <td style="{est}">{pct:.0f}% {ico}</td>
                </tr>"""

        # Construir HTML del bloque
        html_bloques += f"""
        <div style="margin-top: 25px; border-top: 2px solid #2F5597; padding-top: 10px;">
            <h3 style="color: #2F5597; margin-bottom: 8px;">📊 Reporte: {sec}</h3>
            
            <table class="data-table">
                <tr><th colspan="2" class="head-comp">{sec} Totales</th></tr>
                <tr><td>Meta</td><td><b>{tot_m}</b></td></tr>
                <tr><td>Real / Ejecutado</td><td><b>{tot_r}</b></td></tr>
                <tr style="{est_tot}"><td>% Cumpl.</td><td>{pct_tot:.0f}% {ico_tot}</td></tr>
            </table>

            <div class="seccion">■ Por Regional ({sec}):</div>
            <table class="data-table">
                <tr>
                    <th class="head-base">Regional</th>
                    <th class="head-comp">Meta</th><th class="head-comp">Real</th>
                    <th class="head-comp">% Cumplimiento</th>
                </tr>
                {filas_reg}
            </table>

            <div class="seccion">■ Por Persona / Desarrollador ({sec}):</div>
            <table class="data-table">
                <tr>
                    <th class="head-base">Nombre Completo</th>
                    <th class="head-comp">Meta</th><th class="head-comp">Real</th>
                    <th class="head-comp">% Cumplimiento</th>
                </tr>
                {filas_dev}
            </table>
        </div>
        """

    return f"""
    <html>
    <head>
    <style>
        body {{ font-family: 'Bahnschrift SemiCondensed', 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #333; }}
        table {{ border-collapse: collapse; }}
        .data-table {{ width: auto; margin-bottom: 15px; }}
        .data-table th, .data-table td {{ border: 1px solid #ccc; padding: 4px 10px; text-align: center; white-space: nowrap; }}
        .head-base {{ background-color: #f2f2f2; color: #000; font-weight: bold; }}
        .head-comp {{ background-color: #DDEBF7; color: #000; }} 
        .seccion {{ font-weight: bold; color: #2F5597; font-size: 13px; margin: 10px 0 5px 0; }}
    </style>
    </head>
    <body>
        <p>Buenos días Equipo,</p>
        <p>Comparto reporte consolidado al <b>{fecha_corte}</b>.</p>
        <p><i>Nota: Meta esperada de cumplimiento a hoy: <b>{meta_esperada}%</b></i></p>

        {html_bloques if html_bloques else '<p><b>No se seleccionaron secciones válidas o existentes en el Excel.</b></p>'}

        <p>Cordialmente,</p>
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
            return "Por favor selecciona el archivo Excel desde tu escritorio.", 400
        
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
        return f"<div style='color:red; padding:20px;'><b>Error al generar vista previa:</b><pre>{traceback.format_exc()}</pre></div>"

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
        msg['Subject'] = f"Reporte Consolidado - {obtener_fecha_corte_actual()}"
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

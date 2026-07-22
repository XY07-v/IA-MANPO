import os
import smtplib
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

# CONFIGURACIÓN SMTP (Usa variables de entorno o configura tus credenciales aquí)
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.office365.com")  # o smtp.gmail.com
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SMTP_USER = os.getenv("SMTP_USER", "tu_correo@empresa.com")
SMTP_PASS = os.getenv("SMTP_PASS", "tu_contraseña_o_token")

MESES_ESPANOL = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def obtener_fecha_corte_actual():
    """Genera la fecha actual automáticamente formato: '21 de Julio – 2026'"""
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

def generar_html_reporte(df_filtrado, meta_esperada):
    fecha_corte = obtener_fecha_corte_actual()
    
    total_meta = int(df_filtrado['META'].sum()) if 'META' in df_filtrado.columns else 0
    total_real = int(df_filtrado['REAL'].sum()) if 'REAL' in df_filtrado.columns else 0
    pct_total = round((total_real / total_meta * 100), 1) if total_meta > 0 else 0
    estilo_tot, icono_tot = obtener_estilo_cumplimiento(pct_total, meta_esperada)

    # Por Regional
    filas_regional = ""
    if 'REGIONAL' in df_filtrado.columns:
        df_reg = df_filtrado.groupby('REGIONAL')[['META', 'REAL']].sum().reset_index()
        for _, row in df_reg.iterrows():
            m, r = int(row['META']), int(row['REAL'])
            pct = round((r / m * 100), 1) if m > 0 else 0
            estilo, icono = obtener_estilo_cumplimiento(pct, meta_esperada)
            filas_regional += f"""
            <tr>
                <td style="text-align:left; background-color: #f9f9f9;"><b>{row['REGIONAL']}</b></td>
                <td>{m}</td><td>{r}</td>
                <td style="{estilo}">{pct:.0f}% {icono}</td>
            </tr>"""

    # Por Desarrollador
    filas_dev = ""
    if 'DESARROLLADOR' in df_filtrado.columns:
        df_dev = df_filtrado.groupby('DESARROLLADOR')[['META', 'REAL']].sum().reset_index().sort_values(by='REAL', ascending=False)
        for _, row in df_dev.iterrows():
            m, r = int(row['META']), int(row['REAL'])
            pct = round((r / m * 100), 1) if m > 0 else 0
            estilo, icono = obtener_estilo_cumplimiento(pct, meta_esperada)
            filas_dev += f"""
            <tr>
                <td style="text-align:left;">{row['DESARROLLADOR']}</td>
                <td>{m}</td><td>{r}</td>
                <td style="{estilo}">{pct:.0f}% {icono}</td>
            </tr>"""

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
        .seccion {{ font-weight: bold; color: #2F5597; font-size: 14px; margin: 15px 0 5px 0; }}
    </style>
    </head>
    <body>
        <p>Buenos días Equipo,</p>
        <p>Comparto reporte de <b>Competencia</b> al <b>{fecha_corte}</b>.</p>
        <p><i>Nota: Meta esperada a hoy: <b>{meta_esperada}%</b></i></p>

        <table class="data-table">
            <tr><th colspan="2" class="head-comp">Competencia</th></tr>
            <tr><td>Meta</td><td><b>{total_meta}</b></td></tr>
            <tr><td>Real</td><td><b>{total_real}</b></td></tr>
            <tr style="{estilo_tot}"><td>% Cumpl.</td><td>{pct_total:.0f}% {icono_tot}</td></tr>
        </table>

        <div class="seccion">■ Por Regional:</div>
        <table class="data-table">
            <tr>
                <th class="head-base">Regional</th>
                <th class="head-comp">Meta Comp.</th><th class="head-comp">Real Comp.</th>
                <th class="head-comp">% Cumplimiento</th>
            </tr>
            {filas_regional}
        </table>

        <div class="seccion">■ Por Desarrollador:</div>
        <table class="data-table">
            <tr>
                <th class="head-base">Nombres</th>
                <th class="head-comp">Meta Comp.</th><th class="head-comp">Real Comp.</th>
                <th class="head-comp">% Cumplimiento</th>
            </tr>
            {filas_dev}
        </table>
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
        meses = df['MES'].dropna().unique().tolist() if 'MES' in df.columns else []
        regionales = df['REGIONAL'].dropna().unique().tolist() if 'REGIONAL' in df.columns else []
        return jsonify({"meses": meses, "regionales": regionales})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/preview', methods=['POST'])
def preview():
    file = request.files.get('file')
    if not file:
        return "Por favor selecciona el archivo Excel desde tu escritorio.", 400
    
    df = pd.read_excel(file, sheet_name="BASE")
    mes = request.form.get('mes')
    regional = request.form.get('regional')
    meta_esperada = float(request.form.get('meta_esperada', 67))

    if mes and mes != 'TODOS' and 'MES' in df.columns:
        df = df[df['MES'] == mes]
    if regional and regional != 'TODOS' and 'REGIONAL' in df.columns:
        df = df[df['REGIONAL'] == regional]

    return generar_html_reporte(df, meta_esperada)

@app.route('/send', methods=['POST'])
def send():
    file = request.files.get('file')
    if not file:
        return jsonify({"message": "Falta el archivo Excel"}), 400

    try:
        df = pd.read_excel(file, sheet_name="BASE")
        mes = request.form.get('mes')
        regional = request.form.get('regional')
        meta_esperada = float(request.form.get('meta_esperada', 67))

        if mes and mes != 'TODOS' and 'MES' in df.columns:
            df = df[df['MES'] == mes]
        if regional and regional != 'TODOS' and 'REGIONAL' in df.columns:
            df = df[df['REGIONAL'] == regional]

        html_body = generar_html_reporte(df, meta_esperada)
        correos = DESTINATARIOS_DICT.get(regional, DESTINATARIOS_DICT.get('TODOS', []))

        # Envío por SMTP
        msg = MIMEMultipart('alternative')
        msg['Subject'] = f"Reporte de Competencia - {obtener_fecha_corte_actual()}"
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

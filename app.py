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

# MAPEO DE SECCIONES A TUS COLUMNAS EN EXCEL
MAPEO_SECCIONES = {
    "Competencia": {"meta": "Metas Competencia", "real": "SP. Competencia", "efe": "Efe Competencia", "label": "Comp."},
    "Leads": {"meta": "Metas Leads", "real": "Leads.Leads", "efe": "Efe Leads", "label": "Leads"},
    "Prospecciones": {"meta": "Metas Prospecciones", "real": "Prospecciones", "efe": "Efe Prospecciones", "label": "Prosp."},
    "Food": {"meta": "Metas Food", "real": "SP. Food", "efe": "Efe Food", "label": "Food"},
    "Ingredientes": {"meta": "Metas Ingredientes", "real": "SP. Ingredientes", "efe": "Efe Ingredientes", "label": "Ingred."},
    "Puntos Claves": {"meta": "Puntos Programados Claves", "real": "Poc Visitados Claves", "efe": "Efe Puntos Claves", "label": "Pts. Claves"},
    "Puntos Normal": {"meta": "Puntos Programados Normal", "real": "Poc Visitados Normal", "efe": "Efe Puntos Normales", "label": "Pts. Normal"},
    "Logueos": {"meta": "Metas Logueos", "real": "Faltas Logueos", "efe": "Efe Logueos", "label": "Logueos"},
    "Deslogueos": {"meta": "Metas Deslogueos", "real": "Faltas Deslogueos", "efe": "Efe Deslogueos", "label": "Deslogueos"},
    "Sell Out": {"meta": "Meta Sell Out", "real": "Ejecutado Sell Out", "efe": "efe Sell Out", "label": "Sell Out"},
    "Sell In": {"meta": "Meta Sell In", "real": "Ejecutado Sell In", "efe": "Efe Sell In", "label": "Sell In"},
    "Instalaciones": {"meta": "Meta instalaciones", "real": "Ejecutado Instalaciones", "efe": "Efe instalaciones", "label": "Instal."},
    "Retiros": {"meta": "Meta Retiros", "real": "Ejecutado Retiros", "efe": "Efe Retiros", "label": "Retiros"},
    "Netas": {"meta": "Meta Netas", "real": "Ejecutado Netas", "efe": "Efe Netas", "label": "Netas"}
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
    df.columns = df.columns.str.strip()

    col_persona = None
    for posible in ['NOMBRE COMPLETO', 'NOMBRE', 'DESARROLLADOR', 'CARGO']:
        if posible in df.columns:
            col_persona = posible
            break

    col_regional = 'REGIONAL' if 'REGIONAL' in df.columns else None

    html_bloques = ""

    for sec in secciones_seleccionadas:
        if sec not in MAPEO_SECCIONES:
            continue

        info = MAPEO_SECCIONES[sec]
        meta_col = info['meta'].strip()
        real_col = info['real'].strip()
        lbl = info['label']

        if meta_col not in df.columns or real_col not in df.columns:
            continue

        # Asegurar números
        df[meta_col] = pd.to_numeric(df[meta_col], errors='coerce').fillna(0)
        df[real_col] = pd.to_numeric(df[real_col], errors='coerce').fillna(0)

        # ----------------------------------------------------
        # 1. TABLA GLOBAL DEL INDICADOR
        # ----------------------------------------------------
        tot_meta = int(df[meta_col].sum())
        tot_real = int(df[real_col].sum())
        
        # Si la meta total global es 0, omitir la sección por completo
        if tot_meta == 0:
            continue

        pct_global = round((tot_real / tot_meta * 100), 1)
        est_g, ico_g = estilo_cumplimiento(pct_global, meta_esperada)

        tabla_global = f"""
        <table class="data-table table-global">
            <tr><th colspan="2" class="head-blue">{sec}</th></tr>
            <tr><td>Meta</td><td><b>{tot_meta}</b></td></tr>
            <tr><td>Real</td><td><b>{tot_real}</b></td></tr>
            <tr style="{est_g}"><td>% CumpL.</td><td>{pct_global:.0f}% {ico_g}</td></tr>
        </table>
        """

        # ----------------------------------------------------
        # 2. TABLA POR REGIONAL (Meta > 0, Ordenado por % CumpL desc)
        # ----------------------------------------------------
        tabla_regional_html = ""
        if col_regional:
            df_reg = df.groupby(col_regional)[[meta_col, real_col]].sum().reset_index()
            # OMITIR METAS = 0
            df_reg = df_reg[df_reg[meta_col] > 0].copy()
            df_reg['PCT'] = (df_reg[real_col] / df_reg[meta_col]) * 100
            # ORDENAR DE MAYOR A MENOR EFECTIVIDAD
            df_reg = df_reg.sort_values(by='PCT', ascending=False)

            filas_reg = ""
            for _, r in df_reg.iterrows():
                m, real_v = int(r[meta_col]), int(r[real_col])
                pct = r['PCT']
                est, ico = estilo_cumplimiento(pct, meta_esperada)
                filas_reg += f"""
                <tr>
                    <td style="text-align:left; font-weight:bold;">{r[col_regional]}</td>
                    <td>{m}</td>
                    <td>{real_v}</td>
                    <td style="{est}">{pct:.0f}% {ico}</td>
                </tr>"""

            if filas_reg:
                tabla_regional_html = f"""
                <div class="sub-title">■ Por Regional:</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th class="head-gray">Regional</th>
                            <th class="head-blue">Meta {lbl}</th>
                            <th class="head-blue">Real {lbl}</th>
                            <th class="head-blue">% Cumplimiento</th>
                        </tr>
                    </thead>
                    <tbody>{filas_reg}</tbody>
                </table>"""

        # ----------------------------------------------------
        # 3. TABLA POR DESARROLLADOR/PERSONA (Meta > 0, Ordenado desc)
        # ----------------------------------------------------
        tabla_dev_html = ""
        if col_persona:
            df_dev = df.groupby(col_persona)[[meta_col, real_col]].sum().reset_index()
            # OMITIR METAS = 0 O VACÍAS
            df_dev = df_dev[df_dev[meta_col] > 0].copy()
            df_dev['PCT'] = (df_dev[real_col] / df_dev[meta_col]) * 100
            # ORDENAR DE MAYOR A MENOR EFECTIVIDAD
            df_dev = df_dev.sort_values(by='PCT', ascending=False)

            filas_dev = ""
            for _, r in df_dev.iterrows():
                m, real_v = int(r[meta_col]), int(r[real_v] if real_v in r else r[real_col])
                pct = r['PCT']
                est, ico = estilo_cumplimiento(pct, meta_esperada)
                filas_dev += f"""
                <tr>
                    <td style="text-align:left;">{r[col_persona]}</td>
                    <td>{m}</td>
                    <td>{real_v}</td>
                    <td style="{est}">{pct:.0f}% {ico}</td>
                </tr>"""

            if filas_dev:
                tabla_dev_html = f"""
                <div class="sub-title">■ Por Desarrollador:</div>
                <table class="data-table">
                    <thead>
                        <tr>
                            <th class="head-gray">Nombres</th>
                            <th class="head-blue">Meta {lbl}</th>
                            <th class="head-blue">Real {lbl}</th>
                            <th class="head-blue">% Cumplimiento</th>
                        </tr>
                    </thead>
                    <tbody>{filas_dev}</tbody>
                </table>"""

        # UNIR LAS 3 TABLAS EN EL ORDEN SOLICITADO
        html_bloques += f"""
        <div class="bloque-indicador">
            {tabla_global}
            {tabla_regional_html}
            {tabla_dev_html}
        </div>
        <br>
        """

    return f"""
    <html>
    <head>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size: 12px; color: #333; margin: 15px; }}
        .bloque-indicador {{ margin-bottom: 25px; }}
        .data-table {{ border-collapse: collapse; margin-top: 5px; font-size: 11px; }}
        .data-table th, .data-table td {{ border: 1px solid #b0c4de; padding: 5px 12px; text-align: center; white-space: nowrap; }}
        
        .table-global {{ width: 180px; margin-bottom: 12px; }}
        .table-global td {{ font-size: 12px; }}
        
        .head-blue {{ background-color: #DDEBF7; color: #000; font-weight: bold; border: 1px solid #b0c4de; }}
        .head-gray {{ background-color: #EAEAEA; color: #000; font-weight: bold; border: 1px solid #b0c4de; }}
        
        .sub-title {{ font-weight: bold; color: #1F497D; font-size: 12px; margin-top: 12px; margin-bottom: 4px; }}
        tr:nth-child(even) {{ background-color: #FAFAFA; }}
    </style>
    </head>
    <body>
        <p>Buenos días Equipo,</p>
        <p>Comparto el reporte detallado al <b>{obtener_fecha_corte_actual()}</b>.</p>
        <p><i>Meta esperada de cumplimiento: <b>{meta_esperada}%</b></i></p>

        {html_bloques if html_bloques else '<p><b>No hay datos con meta mayor a 0 para las secciones seleccionadas.</b></p>'}

        <p style="margin-top:20px;">Cordialmente,<br><b>Sistema de Reportes</b></p>
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

        # Filtro de Mes
        if mes and mes != 'TODOS' and 'MES' in df.columns:
            if pd.api.types.is_numeric_dtype(df['MES']):
                try: mes = float(mes)
                except ValueError: pass
            df = df[df['MES'] == mes]
            
        # Filtro de Regional
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
        msg['Subject'] = f"Reporte de Competencia y Gestión - {obtener_fecha_corte_actual()}"
        msg['From'] = SMTP_USER
        msg['To'] = ", ".join(correos)
        msg.attach(MIMEText(html_body, 'html'))

        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.sendmail(SMTP_USER, correos, msg.as_string())

        return jsonify({"message": f"Correo enviado con éxito a: {', '.join(correos)}"})
    except Exception as e:
        return jsonify({"message": f"Error al enviar correo: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(port=5000, debug=True)

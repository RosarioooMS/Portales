import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import streamlit as st
import io
import base64
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ===================== FUNCIONES =====================
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

# ===================== CONFIGURACIÓN DE IMÁGENES =====================
BASE_DIR = os.path.dirname(__file__)
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")
FONDO_PATH = os.path.join(BASE_DIR, "depa.png")

logo_base64 = get_base64_of_bin_file(LOGO_PATH)
fondo_base64 = get_base64_of_bin_file(FONDO_PATH)

st.markdown(f"""
<style>
    .dashboard-container {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 100%;
        background-image: url("data:image/png;base64,{fondo_base64}");
        background-size: cover;
        background-repeat: no-repeat;
        background-position: center;
        z-index: -1;
    }}
    .reportview-container .main .block-container {{
        position: relative;
        z-index: 1;
        padding: 20px;
    }}
    .header img {{
        height: 200px;
    }}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="header" style="display:flex;align-items:center;gap:15px;">
  <img src="data:image/png;base64,{logo_base64}" alt="Los Portales">
</div>
""", unsafe_allow_html=True)

# ===================== FUNCIONES DE TABLAS =====================
def crear_tabla_financiamiento(filtrados, grupo):
    grupo_df = filtrados[filtrados['GRUPO'] == grupo]
    if grupo_df.empty:
        return None

    max_ventas = grupo_df['NUM DE VENTAS'].max()
    columnas_ventas = [f"Si logra {i} venta{'s' if i > 1 else ''} del {grupo}" if i < max_ventas 
                       else f"Si logra {i} ventas a más del {grupo}" for i in range(1, max_ventas + 1)]

    if grupo == "Grupo 1":
        filas = ["Contado","Cierre de Proforma CO","Crédito Hipotecario cuota inicial completa",
                 "Cierre de Proforma CIC","Desembolso CIC","Crédito Hipotecario cuota inicial fraccionado",
                 "Cierre de Proforma CIF","Fin de pago de cuota inicial CIF","Desembolso CIF",
                 "Crédito Directo","Cierre de Proforma CD","50% cuota del Credito directo CD"]
        porcentajes_logro = {"Contado": "", "Cierre de Proforma CO": "100%", "Crédito Hipotecario cuota inicial completa": "",
                             "Cierre de Proforma CIC": "80%", "Desembolso CIC": "20%", "Crédito Hipotecario cuota inicial fraccionado": "",
                             "Cierre de Proforma CIF": "70%", "Fin de pago de cuota inicial CIF": "10%", "Desembolso CIF": "20%",
                             "Crédito Directo": "", "Cierre de Proforma CD": "80%", "50% cuota del Credito directo CD": "20%"}
        mapping = {"Cierre de Proforma":{"CO":"Cierre de Proforma CO","CIC":"Cierre de Proforma CIC",
                                        "CIF":"Cierre de Proforma CIF","CD":"Cierre de Proforma CD"},
                   "Desembolso":{"CIC":"Desembolso CIC","CIF":"Desembolso CIF"},
                   "Fin de pago de cuota inicial":"Fin de pago de cuota inicial CIF",
                   "50% cuota del Credito directo":"50% cuota del Credito directo CD"}
        columnas_usar = ["Cierre de Proforma","Fin de pago de cuota inicial","50% cuota del Credito directo","Desembolso"]

    elif grupo == "Grupo 2":
        filas = ["Plan Ahorro - Hasta 9 Meses","Cierre Plan Ahorro - Mixto","Fin de pago de cuotas de plan Ahorro","Desembolso"]
        porcentajes_logro = {"Plan Ahorro - Hasta 9 Meses": "", "Cierre Plan Ahorro - Mixto": "60%",
                             "Fin de pago de cuotas de plan Ahorro": "10%", "Desembolso": "30%"}
        mapping = {"Cierre de Proforma":{"CO":"Plan Ahorro - Hasta 9 Meses","PA":"Cierre Plan Ahorro - Mixto"},
                   "Fin de pago de cuota inicial":"Fin de pago de cuotas de plan Ahorro",
                   "Desembolso":"Desembolso"}
        columnas_usar = ["Cierre de Proforma","Fin de pago de cuota inicial","Desembolso"]

    elif grupo == "Grupo 3":
        filas = ["Plan Ahorro - Hasta 6 Meses","Cierre Plan Ahorro - Mixto","Fin de pago de cuotas de plan Ahorro","Desembolso"]
        porcentajes_logro = {"Plan Ahorro - Hasta 6 Meses": "", "Cierre Plan Ahorro - Mixto": "60%",
                             "Fin de pago de cuotas de plan Ahorro": "10%", "Desembolso": "30%"}
        mapping = {"Cierre de Proforma":{"CO":"Plan Ahorro - Hasta 6 Meses","PA":"Cierre Plan Ahorro - Mixto"},
                   "Fin de pago de cuota inicial":"Fin de pago de cuotas de plan Ahorro",
                   "Desembolso":"Desembolso"}
        columnas_usar = ["Cierre de Proforma","Fin de pago de cuota inicial","Desembolso"]
    else:
        return None

    columnas_finales = ["% logro"] + columnas_ventas
    tabla = pd.DataFrame("", index=filas, columns=columnas_finales)
    for fila, val in porcentajes_logro.items():
        tabla.loc[fila, "% logro"] = val

    for idx, row in grupo_df.iterrows():
        num_ventas = int(row['NUM DE VENTAS'])
        for col in columnas_usar:
            valor = row.get(col)
            if pd.isna(valor): continue
            valor_formato = f"{round(valor*100,2):.2f}%"
            fila_nombre = mapping[col][row['FORMA DE PAGO']] if isinstance(mapping[col], dict) else mapping[col]
            fila_idx = filas.index(fila_nombre)
            col_idx = min(num_ventas-1,len(columnas_ventas)-1)
            tabla.iloc[fila_idx, col_idx+1] = valor_formato

    return tabla

def crear_tabla_tipo_unid(filtrados, proyecto):
    if filtrados.empty: return None
    es_df = filtrados[filtrados['TIPO UNID']=='ES']
    dp_df = filtrados[filtrados['TIPO UNID']=='DP']
    if es_df.empty and dp_df.empty: return None

    n_es = len(es_df)
    columnas_finales = [f"Vende {i}" for i in range(1,n_es+1)] if n_es>0 else []
    tabla = pd.DataFrame(index=["Pago spot por venta de Estacionamiento","Pago spot por cada venta de Depósito"],
                         columns=columnas_finales)
    for i, (_, row) in enumerate(es_df.iterrows()):
        if i < len(tabla.columns):
            tabla.iloc[0,i] = row.get('SPOT',"")
    if not dp_df.empty and 'SPOT' in dp_df.columns and n_es>0:
        mid_col = n_es//2
        tabla.iloc[1,:] = ""
        tabla.iloc[1,mid_col] = dp_df.iloc[0]['SPOT']
    tabla.attrs["titulo"] = proyecto
    return tabla

# ===================== STREAMLIT APP =====================
st.set_page_config(page_title="Los Portales - Ficha Departamentos", layout="wide")
st.title("📊 Generar Ficha Departamentos")
st.markdown('<div class="dashboard-container">', unsafe_allow_html=True)

archivo = st.file_uploader("📂 Sube el archivo Excel (hoja 'Ficha' y 'Dim Grupos')", type=["xlsx","xls"])

if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name='Ficha')
        df_grupos = pd.read_excel(archivo, sheet_name='Dim Grupos')
        df['PERIODO'] = pd.to_datetime(df['PERIODO'], dayfirst=True)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        st.stop()

    periodos = []
    fecha = datetime(2023,6,1)
    fecha_fin = datetime(2025,6,1)
    while fecha<=fecha_fin:
        periodos.append(f"{fecha.day}/{fecha.month}/{fecha.year}")
        fecha += relativedelta(months=1)

    proyectos = ['GRAN GUARDIA PERUANA - ETAPA 1','MF GRAN PLAZA LORETO','GRAN CENTRAL COLONIAL 2 -  ETAPA I',
                 'TICINO','MANDRAGORA','MF LA MAR','MF GRAN CENTRAL COLONIAL','GRAN TOMAS VALLE',
                 'MF LIMA NORTE CARABAYLLO','VILLA RIVIERA','MF GRAN COLONIAL','MF VILLA RIVIERA - ETAPA 2',
                 'GRAN JARDIN TICINO 2 - ETAPA 1','GRAN GUARDIA PERUANA - ETAPA 2']

    periodo = st.selectbox("Seleccione PERIODO:", periodos)
    proyecto = st.selectbox("Seleccione PROYECTO:", proyectos)

    if st.button("📊 Crear Ficha Departamentos"):
        dia, mes, anio = map(int, periodo.split('/'))
        periodo_fecha = datetime(anio, mes, dia)
        filtrados = df[(df['PERIODO']==periodo_fecha)&(df['PROYECTO']==proyecto)]
        comentarios = df_grupos[(df_grupos['PERIODO']==periodo_fecha)&(df_grupos['PROYECTO']==proyecto)]

        if filtrados.empty:
            st.warning("⚠️ No se encontraron datos para ese PERIODO y PROYECTO.")
        else:
            tabla1 = crear_tabla_financiamiento(filtrados,"Grupo 1")
            tabla2 = crear_tabla_financiamiento(filtrados,"Grupo 2")
            tabla3 = crear_tabla_financiamiento(filtrados,"Grupo 3")
            tabla_tipo_unid = crear_tabla_tipo_unid(filtrados, proyecto)

            for i, t in enumerate([tabla1, tabla2, tabla3, tabla_tipo_unid], start=1):
                if t is not None:
                    st.subheader(f"📌 Ficha Departamentos - Grupo {i}" if i<=3 else f"📌 Ficha Departamentos - {proyecto} (ES / DP)")
                    st.dataframe(t)

            if all(x is None for x in [tabla1, tabla2, tabla3, tabla_tipo_unid]):
                st.info("No hay datos para ninguno de los grupos.")

            # ================== EXPORTAR A PDF ==================
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()
            style_cell = ParagraphStyle(name="cell", fontSize=8, alignment=TA_CENTER, leading=10)
            style_index = ParagraphStyle(name="index", fontSize=8, alignment=TA_LEFT, leading=10)
            def safe_str(val): return "" if pd.isna(val) else str(val)

            # --- Logo ---
            try:
                img = Image(LOGO_PATH, width=80, height=50)
                img.hAlign = 'CENTER'
                elements.append(img)
                elements.append(Spacer(1,12))
            except Exception as e:
                st.warning(f"No se pudo cargar el logo en el PDF: {e}")

            # --- Título proyecto ---
            elements.append(Paragraph(f"<b>{proyecto}</b>", ParagraphStyle(
                name="titulo_proyecto", fontSize=14, alignment=TA_CENTER, spaceAfter=20)))
            elements.append(Spacer(1,12))

            # --- Comentarios generales ---
            for grupo_general in ["A) CONSIDERACIONES","B) OBJETIVO"]:
                comentarios_general = df_grupos[(df_grupos['PERIODO']==periodo_fecha)&
                                                (df_grupos['PROYECTO']==proyecto)&
                                                (df_grupos['Grupo']==grupo_general)]
                for _, row in comentarios_general.iterrows():
                    elements.append(Paragraph(f"<b>{grupo_general}</b>: {safe_str(row['Comentario'])}", styles["Normal"]))
                    elements.append(Spacer(1,8))

            # --- Primer comentario de comisión antes de tablas ---
            comision_total = df_grupos[(df_grupos['PERIODO']==periodo_fecha)&
                                       (df_grupos['PROYECTO']==proyecto)&
                                       (df_grupos['Grupo']=="1. Comision por unidad Vendida***")]
            if not comision_total.empty:
                primer_com = comision_total.iloc[0]
                elements.append(Paragraph(f"1. Comision por unidad Vendida***: {safe_str(primer_com['Comentario'])}", styles["Normal"]))
                elements.append(Spacer(1,8))
                resto_comisiones = comision_total.iloc[1:]
            else:
                resto_comisiones = pd.DataFrame()

            # --- Función para DataFrame a tabla PDF ---
            def df_to_table(df, titulo):
                header = [Paragraph(df.index.name or "", style_index)] + [Paragraph(str(c), style_cell) for c in df.columns]
                data = [header]
                for idx, row in df.iterrows():
                    data.append([Paragraph(str(idx), style_index)] + [Paragraph(safe_str(v), style_cell) for v in row.values])
                col_widths = [(A4[0]-2*cm)/len(header)]*len(header) if len(header)>0 else []
                table = Table(data, colWidths=col_widths, repeatRows=1)
                table.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,0),colors.lightgrey),
                    ('GRID',(0,0),(-1,-1),0.5,colors.black),
                    ('ALIGN',(1,0),(-1,-1),'CENTER'),
                    ('ALIGN',(0,0),(0,-1),'LEFT'),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                    ('FONTSIZE',(0,0),(-1,-1),8),
                    ('LEFTPADDING',(0,0),(-1,-1),4),
                    ('RIGHTPADDING',(0,0),(-1,-1),4),
                    ('TOPPADDING',(0,0),(-1,-1),4),
                    ('BOTTOMPADDING',(0,0),(-1,-1),4),
                ]))
                elements.append(Paragraph(f"<b>{titulo}</b>", styles["Heading3"]))
                elements.append(Spacer(1,8))
                elements.append(table)
                elements.append(Spacer(1,16))

            # --- Insertar tablas ---
            for idx, (tabla, titulo, grupo) in enumerate([(tabla1,"Ficha Departamentos - Grupo 1","Grupo 1"),
                                                          (tabla2,"Ficha Departamentos - Grupo 2","Grupo 2"),
                                                          (tabla3,"Ficha Departamentos - Grupo 3","Grupo 3"),
                                                          (tabla_tipo_unid,f"Ficha Departamentos - {proyecto} (ES / DP)",None)], start=1):
                if tabla is not None:
                    df_to_table(tabla, titulo)
                    if grupo=="Grupo 3" and not resto_comisiones.empty:
                        for _, row in resto_comisiones.iterrows():
                            elements.append(Paragraph(f"{safe_str(row['Comentario'])}", styles["Normal"]))
                            elements.append(Spacer(1,8))

            # --- Comentarios finales ---
            for g_final in ["E) RESOLUCIONES","F) PENALIDAD","F) PENALIDAD (Esto no es de penalidad pero va al último)"]:
                comentarios_final = df_grupos[(df_grupos['PERIODO']==periodo_fecha)&
                                              (df_grupos['PROYECTO']==proyecto)&
                                              (df_grupos['Grupo']==g_final)]
                for _, row in comentarios_final.iterrows():
                    elements.append(Paragraph(f"<b>{g_final}</b>: {safe_str(row['Comentario'])}", styles["Normal"]))
                    elements.append(Spacer(1,8))

            doc.build(elements)
            buffer.seek(0)
            st.download_button(label="⬇️ Descargar Ficha Departamentos en PDF", data=buffer.getvalue(),
                               file_name="Ficha_Departamentos.pdf", mime="application/pdf")

st.markdown('</div>', unsafe_allow_html=True)

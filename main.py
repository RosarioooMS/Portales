import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import streamlit as st
import io
import base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ===================== FUNCIONES =====================
# Función para convertir imagen a base64
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

# Logo (ajustado a la ruta dentro de Prueba/)
logo_base64 = get_base64_of_bin_file("Prueba/logo.png")

# Header con logo y título
st.markdown(f"""
<div class="header" style="display:flex;align-items:center;gap:15px;">
  <img src="data:image/png;base64,{logo_base64}" alt="Los Portales" style="height:200px;">
</div>
""", unsafe_allow_html=True)

# ===================== TABLA GENERICA =====================
def crear_tabla_financiamiento(filtrados, grupo):
    grupo_df = filtrados[filtrados['GRUPO'] == grupo]

    if grupo_df.empty:
        return None

    max_ventas = grupo_df['NUM DE VENTAS'].max()
    columnas_ventas = []
    for i in range(1, max_ventas + 1):
        if i == max_ventas:
            columnas_ventas.append(f"Si logra {i} ventas a más del {grupo}")
        else:
            columnas_ventas.append(f"Si logra {i} venta{'s' if i > 1 else ''} del {grupo}")

    # Configuración según grupo
    if grupo == "Grupo 1":
        filas = [
            "Contado",
            "Cierre de Proforma CO",
            "Crédito Hipotecario cuota inicial completa",
            "Cierre de Proforma CIC",
            "Desembolso CIC",
            "Crédito Hipotecario cuota inicial fraccionado",
            "Cierre de Proforma CIF",
            "Fin de pago de cuota inicial CIF",
            "Desembolso CIF",
            "Crédito Directo",
            "Cierre de Proforma CD",
            "50% cuota del Credito directo CD"
        ]
        porcentajes_logro = {
            "Contado": "",
            "Cierre de Proforma CO": "100%",
            "Crédito Hipotecario cuota inicial completa": "" ,
            "Cierre de Proforma CIC": "80%",
            "Desembolso CIC": "20%",
            "Crédito Hipotecario cuota inicial fraccionado": "",
            "Cierre de Proforma CIF": "70%",
            "Fin de pago de cuota inicial CIF": "10%",
            "Desembolso CIF": "20%",
            "Crédito Directo": "",
            "Cierre de Proforma CD": "80%",
            "50% cuota del Credito directo CD": "20%"
        }
        mapping = {
            "Cierre de Proforma": {
                "CO": "Cierre de Proforma CO",
                "CIC": "Cierre de Proforma CIC",
                "CIF": "Cierre de Proforma CIF",
                "CD": "Cierre de Proforma CD"
            },
            "Desembolso": {
                "CIC": "Desembolso CIC",
                "CIF": "Desembolso CIF"
            },
            "Fin de pago de cuota inicial": "Fin de pago de cuota inicial CIF",
            "50% cuota del Credito directo": "50% cuota del Credito directo CD"
        }
        columnas_usar = ["Cierre de Proforma","Fin de pago de cuota inicial","50% cuota del Credito directo","Desembolso"]

    elif grupo == "Grupo 2":
        filas = [
            "Plan Ahorro - Hasta 9 Meses",
            "Cierre Plan Ahorro - Mixto",
            "Fin de pago de cuotas de plan Ahorro",
            "Desembolso"
        ]
        porcentajes_logro = {
            "Plan Ahorro - Hasta 9 Meses": "",
            "Cierre Plan Ahorro - Mixto": "60%",
            "Fin de pago de cuotas de plan Ahorro": "10%",
            "Desembolso": "30%"
        }
        mapping = {
            "Cierre de Proforma": {
                "CO": "Plan Ahorro - Hasta 9 Meses",
                "PA": "Cierre Plan Ahorro - Mixto"
            },
            "Fin de pago de cuota inicial": "Fin de pago de cuotas de plan Ahorro",
            "Desembolso": "Desembolso"
        }
        columnas_usar = ["Cierre de Proforma","Fin de pago de cuota inicial","Desembolso"]

    elif grupo == "Grupo 3":
        filas = [
            "Plan Ahorro - Hasta 6 Meses",
            "Cierre Plan Ahorro - Mixto",
            "Fin de pago de cuotas de plan Ahorro",
            "Desembolso"
        ]
        porcentajes_logro = {
            "Plan Ahorro - Hasta 6 Meses": "",
            "Cierre Plan Ahorro - Mixto": "60%",
            "Fin de pago de cuotas de plan Ahorro": "10%",
            "Desembolso": "30%"
        }
        mapping = {
            "Cierre de Proforma": {
                "CO": "Plan Ahorro - Hasta 6 Meses",
                "PA": "Cierre Plan Ahorro - Mixto"
            },
            "Fin de pago de cuota inicial": "Fin de pago de cuotas de plan Ahorro",
            "Desembolso": "Desembolso"
        }
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
            if pd.isna(valor):
                continue
            valor_formato = f"{round(valor*100,2):.2f}%"

            if isinstance(mapping[col], dict):
                forma = row['FORMA DE PAGO']
                if forma in mapping[col]:
                    fila_nombre = mapping[col][forma]
                else:
                    continue
            else:
                fila_nombre = mapping[col]

            fila_idx = filas.index(fila_nombre)
            col_idx = num_ventas - 1
            if col_idx >= len(columnas_ventas):
                col_idx = len(columnas_ventas) - 1
            tabla.iloc[fila_idx, col_idx + 1] = valor_formato

    return tabla

# ===================== NUEVA TABLA POR TIPO UNID =====================
def crear_tabla_tipo_unid(filtrados, proyecto):
    if filtrados.empty:
        return None

    es_df = filtrados[filtrados['TIPO UNID'] == 'ES']
    dp_df = filtrados[filtrados['TIPO UNID'] == 'DP']

    if es_df.empty and dp_df.empty:
        return None

    n_es = len(es_df)
    columnas_finales = [f"Vende {i}" for i in range(1, n_es+1)] if n_es > 0 else []

    tabla = pd.DataFrame(index=[
        "Pago spot por venta de Estacionamiento",
        "Pago spot por cada venta de Depósito"
    ], columns=columnas_finales)

    for i, (_, row) in enumerate(es_df.iterrows()):
        if i < len(tabla.columns):
            tabla.iloc[0, i] = row['SPOT'] if 'SPOT' in row else ""

    if not dp_df.empty and 'SPOT' in dp_df.columns and n_es > 0:
        dp_valor = dp_df.iloc[0]['SPOT']
        mid_col = (n_es // 2)
        tabla.iloc[1, :] = ""  
        tabla.iloc[1, mid_col] = dp_valor

    tabla.attrs["titulo"] = proyecto
    return tabla

# ===================== STREAMLIT APP =====================
st.set_page_config(page_title="Los Portales - Tabla Financiamiento", layout="wide")
st.title("📊 Generar Tablas de Financiamiento")

st.markdown('<div class="dashboard-container">', unsafe_allow_html=True)

archivo = st.file_uploader("📂 Sube el archivo Excel (hoja 'Ficha')", type=["xlsx", "xls"])

if archivo:
    try:
        df = pd.read_excel(archivo, sheet_name='Ficha')
        df['PERIODO'] = pd.to_datetime(df['PERIODO'], dayfirst=True)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        st.stop()

    # Periodos
    periodos = []
    fecha = datetime(2023, 6, 1)
    fecha_fin = datetime(2025, 6, 1)
    while fecha <= fecha_fin:
        periodos.append(f"{fecha.day}/{fecha.month}/{fecha.year}")
        fecha += relativedelta(months=1)

    proyectos = [
        'GRAN GUARDIA PERUANA - ETAPA 1',
        'MF GRAN PLAZA LORETO',
        'GRAN CENTRAL COLONIAL 2 -  ETAPA I',
        'TICINO',
        'MANDRAGORA',
        'MF LA MAR',
        'MF GRAN CENTRAL COLONIAL',
        'GRAN TOMAS VALLE',
        'MF LIMA NORTE CARABAYLLO',
        'VILLA RIVIERA',
        'MF GRAN COLONIAL',
        'MF VILLA RIVIERA - ETAPA 2',
        'GRAN JARDIN TICINO 2 - ETAPA 1',
        'GRAN GUARDIA PERUANA - ETAPA 2'
    ]

    periodo = st.selectbox("Seleccione PERIODO:", periodos)
    proyecto = st.selectbox("Seleccione PROYECTO:", proyectos)

    if st.button("📊 Crear Tablas Financiamiento"):
        dia, mes, anio = map(int, periodo.split('/'))
        periodo_fecha = datetime(anio, mes, dia)

        filtrados = df[(df['PERIODO'] == periodo_fecha) & (df['PROYECTO'] == proyecto)]

        if filtrados.empty:
            st.warning("⚠️ No se encontraron datos para ese PERIODO y PROYECTO.")
        else:
            tabla1 = crear_tabla_financiamiento(filtrados,"Grupo 1")
            tabla2 = crear_tabla_financiamiento(filtrados,"Grupo 2")
            tabla3 = crear_tabla_financiamiento(filtrados,"Grupo 3")
            tabla_tipo_unid = crear_tabla_tipo_unid(filtrados, proyecto)

            if tabla1 is not None:
                st.subheader("📌 Tabla Financiamiento - Grupo 1")
                st.dataframe(tabla1)

            if tabla2 is not None:
                st.subheader("📌 Tabla Financiamiento - Grupo 2")
                st.dataframe(tabla2)

            if tabla3 is not None:
                st.subheader("📌 Tabla Financiamiento - Grupo 3")
                st.dataframe(tabla3)

            if tabla_tipo_unid is not None:
                st.subheader(f"📌 Tabla Financiamiento - {proyecto} (ES / DP)")
                st.dataframe(tabla_tipo_unid)

            if tabla1 is None and tabla2 is None and tabla3 is None and tabla_tipo_unid is None:
                st.info("No hay datos para ninguno de los grupos.")

            # ================== EXPORTAR A PDF ==================
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=A4)
            elements = []
            styles = getSampleStyleSheet()

            style_cell = ParagraphStyle(
                name="cell",
                fontSize=8,
                alignment=TA_CENTER,
                leading=10
            )
            style_index = ParagraphStyle(
                name="index",
                fontSize=8,
                alignment=TA_LEFT,
                leading=10
            )

            def safe_str(val):
                if pd.isna(val):
                    return ""
                if isinstance(val, float):
                    return str(val)
                return str(val)

            def df_to_table(df, titulo):
                header = [Paragraph(df.index.name or "", style_index)] + [Paragraph(str(c), style_cell) for c in df.columns]
                data = [header]

                for idx, row in df.iterrows():
                    row_cells = [Paragraph(str(idx), style_index)]
                    for val in row.values:
                        row_cells.append(Paragraph(safe_str(val), style_cell))
                    data.append(row_cells)

                ancho_total = A4[0] - 2*cm
                num_cols = len(data[0])
                if num_cols == 0:
                    return
                col_widths = [ancho_total/num_cols] * num_cols

                table = Table(data, colWidths=col_widths, repeatRows=1)
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.black),
                    ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (0,0), (0,-1), 'LEFT'),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('FONTSIZE', (0,0), (-1,-1), 8),
                    ('LEFTPADDING', (0,0), (-1,-1), 4),
                    ('RIGHTPADDING', (0,0), (-1,-1), 4),
                    ('TOPPADDING', (0,0), (-1,-1), 4),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 4),
                ]))

                elements.append(Paragraph(f"<b>{titulo}</b>", styles["Heading3"]))
                elements.append(Spacer(1, 8))
                elements.append(table)
                elements.append(Spacer(1, 16))

            if tabla1 is not None:
                df_to_table(tabla1, "Tabla Financiamiento - Grupo 1")
            if tabla2 is not None:
                df_to_table(tabla2, "Tabla Financiamiento - Grupo 2")
            if tabla3 is not None:
                df_to_table(tabla3, "Tabla Financiamiento - Grupo 3")
            if tabla_tipo_unid is not None:
                df_to_table(tabla_tipo_unid, f"Tabla Financiamiento - {proyecto} (ES / DP)")

            doc.build(elements)
            buffer.seek(0)

            st.download_button(
                label="⬇️ Descargar Tablas en PDF",
                data=buffer.getvalue(),
                file_name="Tablas_Financiamiento.pdf",
                mime="application/pdf"
            )

st.markdown('</div>', unsafe_allow_html=True)

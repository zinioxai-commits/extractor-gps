import streamlit as st
import pytesseract
import numpy as np
import re
import io
import os
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── Tesseract path para Windows ──────────────────────────────────────────────
if os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ── Funciones de extracción ───────────────────────────────────────────────────

def aislar_texto_azul(imagen_pil, franja=0.07):
    w, h = imagen_pil.size
    recorte = imagen_pil.crop((0, int(h * (1 - franja)), w, h))
    arr = np.array(recorte)
    mascara = (
        (arr[:,:,2] > 80)  &
        (arr[:,:,0] < 150) &
        (arr[:,:,1] < 200) &
        (arr[:,:,2] > arr[:,:,0])
    )
    resultado = np.ones_like(arr) * 255
    resultado[mascara] = [0, 0, 0]
    img = Image.fromarray(resultado.astype(np.uint8))
    return img.resize((img.width * 4, img.height * 4), Image.LANCZOS)

def extraer_texto(imagen_pil):
    img_proc = aislar_texto_azul(imagen_pil)
    return pytesseract.image_to_string(img_proc, lang='eng', config='--psm 6 --oem 3')

def parsear_coordenadas(texto):
    tu = texto.upper()
    # Formato limpio: 14.062602S 69.203552W
    m = re.search(r'(\d{1,3})[:\.](\d{6})\s*([NS])\s*(\d{1,3})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        lat = float(f"{m.group(1)}.{m.group(2)}")
        lon = float(f"{m.group(4)}.{m.group(5)}")
        if m.group(3) == 'S': lat = -lat
        if m.group(6) in ('W','O'): lon = -lon
        return lat, lon
    # Formato fusionado OCR: 14:0626028369:203552W
    nums = re.findall(r'\d+', tu)
    hem = re.search(r'([WEO])\s*$', tu.strip())
    if len(nums) >= 3 and hem and len(nums[1]) >= 8:
        lat = -float(f"{nums[0]}.{nums[1][:6]}")
        lon =  float(f"{nums[1][-2:]}.{nums[2][:6]}")
        if hem.group(1) in ('W','O'): lon = -lon
        return lat, lon
    return None, None

def parsear_fecha_hora(texto):
    m = re.search(
        r'(\d{1,2}[\s\-]\w{2,4}[\s\-:]\d{4})\s+(\d{1,2}:\d{2}:\d{2}\s*[ap]\.?m\.?)',
        texto, re.IGNORECASE
    )
    if m:
        fecha = re.sub(r'[\-:]', ' ', m.group(1)).strip()
        return fecha, m.group(2).strip()
    return None, None

def generar_excel(datos):
    wb = Workbook()
    ws = wb.active
    ws.title = "Coordenadas GPS"

    f_titulo = Font(name='Arial', bold=True, color='FFFFFF', size=11)
    f_normal = Font(name='Arial', size=10)
    fill_azul = PatternFill('solid', start_color='2E75B6')
    fill_alt  = PatternFill('solid', start_color='DDEEFF')
    centro = Alignment(horizontal='center', vertical='center')
    izq    = Alignment(horizontal='left',   vertical='center')
    borde  = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'),  bottom=Side(style='thin'))

    encabezados = ['#', 'Archivo', 'Fecha', 'Hora', 'Latitud', 'Longitud', 'Estado']
    for col, enc in enumerate(encabezados, 1):
        c = ws.cell(row=1, column=col, value=enc)
        c.font = f_titulo; c.fill = fill_azul
        c.alignment = centro; c.border = borde
    ws.row_dimensions[1].height = 22

    for i, (archivo, fecha, hora, lat, lon, estado) in enumerate(datos, 1):
        fila = i + 1
        for col, val in enumerate([i, archivo, fecha, hora, lat, lon, estado], 1):
            c = ws.cell(row=fila, column=col, value=val)
            c.font = f_normal
            c.alignment = izq if col == 2 else centro
            c.border = borde
            if i % 2 == 0: c.fill = fill_alt
        ws.cell(row=fila, column=5).number_format = '0.000000'
        ws.cell(row=fila, column=6).number_format = '0.000000'

    for col, ancho in zip('ABCDEFG', [5, 38, 18, 16, 14, 14, 16]):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = 'A2'

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ── UI Streamlit ──────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Extractor GPS",
    page_icon="📍",
    layout="wide"
)

st.title("📍 Extractor de Coordenadas GPS")
st.caption("Sube tus fotos con marca de agua Garmin y descarga las coordenadas en Excel")

st.divider()

archivos = st.file_uploader(
    "Arrastra o selecciona tus fotos",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    help="Puedes subir varias fotos a la vez"
)

if archivos:
    st.divider()
    st.subheader(f"📷 {len(archivos)} foto(s) cargada(s)")

    if st.button("🔍 Extraer coordenadas", type="primary", use_container_width=True):
        datos = []
        resultados_ui = []

        progress = st.progress(0, text="Procesando fotos...")
        cols_header = st.columns([2, 1.5, 1.5, 1.5, 1.5, 1])
        cols_header[0].markdown("**Archivo**")
        cols_header[1].markdown("**Fecha**")
        cols_header[2].markdown("**Hora**")
        cols_header[3].markdown("**Latitud**")
        cols_header[4].markdown("**Longitud**")
        cols_header[5].markdown("**Estado**")
        st.divider()

        for idx, archivo in enumerate(archivos):
            try:
                img = Image.open(archivo)
                texto = extraer_texto(img)
                lat, lon = parsear_coordenadas(texto)
                fecha, hora = parsear_fecha_hora(texto)
                estado = "✅" if lat is not None else "⚠️"

                cols = st.columns([2, 1.5, 1.5, 1.5, 1.5, 1])
                cols[0].write(archivo.name)
                cols[1].write(fecha or "—")
                cols[2].write(hora or "—")
                cols[3].write(f"{lat:.6f}" if lat else "—")
                cols[4].write(f"{lon:.6f}" if lon else "—")
                cols[5].write(estado)

                datos.append((archivo.name, fecha, hora, lat, lon,
                               "OK" if lat else "Sin coordenadas"))

            except Exception as e:
                cols = st.columns([2, 1.5, 1.5, 1.5, 1.5, 1])
                cols[0].write(archivo.name)
                cols[5].write("❌")
                datos.append((archivo.name, None, None, None, None, f"Error: {e}"))

            progress.progress((idx + 1) / len(archivos),
                              text=f"Procesando {idx+1}/{len(archivos)}...")

        progress.empty()
        st.divider()

        ok = sum(1 for d in datos if d[4] is not None)
        if ok > 0:
            st.success(f"✅ {ok} de {len(archivos)} fotos con coordenadas extraídas")
            excel_buf = generar_excel(datos)
            st.download_button(
                label="📥 Descargar Excel",
                data=excel_buf,
                file_name="coordenadas.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        else:
            st.warning("⚠️ No se pudieron extraer coordenadas de ninguna foto.")

else:
    st.info("👆 Sube tus fotos para comenzar")

with st.sidebar:
    st.markdown("### ℹ️ Instrucciones")
    st.markdown("""
    1. Sube una o varias fotos
    2. Haz clic en **Extraer coordenadas**
    3. Revisa los resultados
    4. Descarga el **Excel**

    ---
    **Formato soportado:**
    Marca de agua azul con coordenadas decimales
    `14.062602S 69.203552W`

    **Dispositivos compatibles:**
    Garmin GPSmap y similares
    """)

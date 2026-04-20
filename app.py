import streamlit as st
import pytesseract
import numpy as np
import re
import io
import os
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

if os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ── Extracción ────────────────────────────────────────────────────────────────

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

def extraer_texto_multi(imagen_pil):
    """Prueba varias franjas y devuelve el texto con más información."""
    mejores = ""
    for franja in [0.05, 0.07, 0.08, 0.10]:
        img_proc = aislar_texto_azul(imagen_pil, franja)
        texto = pytesseract.image_to_string(img_proc, lang='eng', config='--psm 6 --oem 3')
        if len(texto.strip()) > len(mejores.strip()):
            mejores = texto
    return mejores

def parsear_coordenadas(texto):
    tu = texto.upper()

    # Caso 1: formato limpio  →  14.062602S 69.203552W
    m = re.search(r'(\d{1,3})[:\.](\d{6})\s*([NS])\s*(\d{1,3})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        lat = float(f"{m.group(1)}.{m.group(2)}")
        lon = float(f"{m.group(4)}.{m.group(5)}")
        if m.group(3) == 'S': lat = -lat
        if m.group(6) in ('W','O'): lon = -lon
        return lat, lon

    # Caso 2: separador no-dígito entre bloques  →  14:0626028'69.203552W
    m = re.search(r'(\d{1,3})[:\.](\d{6})[^0-9]{1,4}(\d{2})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        lat = -float(f"{m.group(1)}.{m.group(2)}")
        lon =  float(f"{m.group(3)}.{m.group(4)}")
        if m.group(5) in ('W','O'): lon = -lon
        return lat, lon

    # Caso 3: todo fusionado  →  14:0626028369:203552W
    m = re.search(r'(\d{1,3})[:\.](\d{6}).{0,4}?(\d{2})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        lat = -float(f"{m.group(1)}.{m.group(2)}")
        lon =  float(f"{m.group(3)}.{m.group(4)}")
        if m.group(5) in ('W','O'): lon = -lon
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

# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Extractor GPS Garmin", page_icon="📍", layout="wide")

st.markdown("""
<style>
.main-title { font-size:2.2rem; font-weight:700; color:#1a3a5c; }
.subtitle   { color:#666; font-size:1rem; margin-bottom:1.5rem; }
.stat-box   { background:#f0f4ff; border-radius:10px; padding:1rem; text-align:center; }
.stat-num   { font-size:2rem; font-weight:700; color:#2E75B6; }
.stat-lbl   { color:#666; font-size:0.85rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<p class="main-title">📍 Extractor de Coordenadas GPS</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Sube tus fotos con marca de agua Garmin · Obtén las coordenadas en Excel al instante</p>', unsafe_allow_html=True)
st.divider()

with st.sidebar:
    st.markdown("### ℹ️ Cómo usar")
    st.markdown("1. 📤 Sube una o varias fotos\n2. 🔍 Clic en **Extraer**\n3. 📊 Revisa la tabla\n4. 📥 Descarga el **Excel**")
    st.divider()
    st.markdown("### 📋 Formato soportado")
    st.code("14.062602S 69.203552W")
    st.caption("Garmin GPSmap y similares")
    st.divider()
    st.caption("v1.1 · Extractor GPS Garmin")

archivos = st.file_uploader(
    "📤 Arrastra tus fotos aquí o haz clic para seleccionar",
    type=["jpg","jpeg","png"],
    accept_multiple_files=True
)

if not archivos:
    st.info("👆 Sube tus fotos para comenzar.")
    st.stop()

st.divider()

# Vista previa
with st.expander(f"🖼️ Vista previa — {len(archivos)} foto(s)", expanded=False):
    cols = st.columns(min(len(archivos), 4))
    for i, f in enumerate(archivos):
        with cols[i % 4]:
            f.seek(0)
            st.image(Image.open(f), caption=f.name, use_container_width=True)

st.divider()

if st.button("🔍 Extraer coordenadas", type="primary", use_container_width=True):
    datos = []
    ok_count = 0

    progress  = st.progress(0, text="Iniciando...")
    status    = st.empty()

    # Encabezados
    h = st.columns([0.4, 2.2, 1.4, 1.4, 1.4, 1.4, 0.5])
    for col, lbl in zip(h, ['#','Archivo','Fecha','Hora','Latitud','Longitud','✓']):
        col.markdown(f"**{lbl}**")
    st.markdown("---")

    for idx, archivo in enumerate(archivos):
        status.info(f"⏳ Procesando **{archivo.name}** ({idx+1}/{len(archivos)})")
        try:
            archivo.seek(0)
            img          = Image.open(archivo)
            texto        = extraer_texto_multi(img)
            lat, lon     = parsear_coordenadas(texto)
            fecha, hora  = parsear_fecha_hora(texto)
            if lat is not None:
                ok_count += 1
                icono = "✅"
            else:
                icono = "⚠️"
            row = st.columns([0.4, 2.2, 1.4, 1.4, 1.4, 1.4, 0.5])
            row[0].write(idx + 1)
            row[1].write(archivo.name)
            row[2].write(fecha or "—")
            row[3].write(hora  or "—")
            row[4].write(f"{lat:.6f}" if lat is not None else "—")
            row[5].write(f"{lon:.6f}" if lon is not None else "—")
            row[6].write(icono)
            datos.append((archivo.name, fecha, hora, lat, lon,
                          "OK" if lat is not None else "Sin coordenadas"))
        except Exception as e:
            row = st.columns([0.4, 2.2, 1.4, 1.4, 1.4, 1.4, 0.5])
            row[0].write(idx + 1); row[1].write(archivo.name); row[6].write("❌")
            datos.append((archivo.name, None, None, None, None, f"Error: {e}"))

        progress.progress((idx+1)/len(archivos), text=f"Procesando {idx+1}/{len(archivos)}...")

    status.empty(); progress.empty()
    st.divider()

    if ok_count > 0:
        fail = len(archivos) - ok_count
        st.success(f"✅ {ok_count} foto(s) con coordenadas · {fail} sin detectar")
        st.download_button(
            label="📥 Descargar Excel",
            data=generar_excel(datos),
            file_name="coordenadas_gps.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.error("❌ No se detectaron coordenadas. Verifica que las fotos tengan la marca de agua azul.")

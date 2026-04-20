import streamlit as st
import pytesseract
import numpy as np
import re
import io
import os
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from collections import Counter

if os.name == 'nt':
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# ── Extracción ────────────────────────────────────────────────────────────────

def aislar_texto_azul(img_pil, franja):
    w, h = img_pil.size
    crop = img_pil.crop((0, int(h * (1 - franja)), w, h))
    arr = np.array(crop)
    mask = (arr[:,:,2] > 80) & (arr[:,:,0] < 150) & (arr[:,:,1] < 200) & (arr[:,:,2] > arr[:,:,0])
    res = np.ones_like(arr) * 255
    res[mask] = [0, 0, 0]
    img2 = Image.fromarray(res.astype(np.uint8))
    return img2.resize((img2.width * 4, img2.height * 4), Image.LANCZOS)

def extraer_lon(texto):
    tu = texto.upper()
    # Con punto decimal
    m = re.search(r'(\d{2,3})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        v = float(f"{m.group(1)}.{m.group(2)}")
        return -v if m.group(3) in ('W', 'O') else v
    # Sin punto (8 dígitos + W)
    m = re.search(r'(\d{2})(\d{6})\s*([WEO])', tu)
    if m:
        v = float(f"{m.group(1)}.{m.group(2)}")
        return -v if m.group(3) in ('W', 'O') else v
    return None

def extraer_lat(texto):
    tu = texto.upper()
    # Caso 1: normal con separador  →  14.060653S
    m = re.search(r'(\d{2})[^0-9A-Z]{0,2}(\d{6})\s*([NS])', tu)
    if m:
        v = float(f"{m.group(1)}.{m.group(2)}")
        return -v if m.group(3) == 'S' else v
    # Caso 2: sin separador  →  14060653S
    m = re.search(r'(\d{2})(\d{6})\s*([NS])', tu)
    if m:
        v = float(f"{m.group(1)}.{m.group(2)}")
        return -v if m.group(3) == 'S' else v
    # Caso 3: dígitos + S + espacio + lon (identifica S-hemisferio real)
    m = re.search(r'(\d+)\s*([NS])\s+\d', tu)
    if m:
        bloque, hem = m.group(1), m.group(2)
        if len(bloque) >= 6:
            dec = bloque[-6:]
            entero = bloque[-8:-6] if len(bloque) >= 8 else "14"
            if not entero or not entero.isdigit() or int(entero) == 0:
                entero = "14"
            v = float(f"{entero}.{dec}")
            if 0 < v < 25:
                return -v if hem == 'S' else v
    # Caso 4: lat fusionada con lon  →  14:0626028369:203552W
    m = re.search(r'(\d{1,3})[:\.](\d{6}).{0,4}?(\d{2})[:\.](\d{6})\s*([WEO])', tu)
    if m:
        return -float(f"{m.group(1)}.{m.group(2)}")
    return None

def votar(valores):
    vals = [round(v, 4) for v in valores if v is not None]
    if not vals:
        return None
    ganador = Counter(vals).most_common(1)[0][0]
    # Devolver el valor original más cercano al ganador (con más precisión)
    originales = [v for v in valores if v is not None and abs(round(v,4) - ganador) < 0.0001]
    return originales[0] if originales else ganador

def parsear_fecha_hora(texto):
    m = re.search(
        r'(\d{1,2}[\s\-]\w{2,4}[\s\-:]\d{4})\s+(\d{1,2}:\d{2}:\d{2}\s*[ap]\.?m\.?)',
        texto, re.IGNORECASE
    )
    if m:
        fecha = re.sub(r'[\-:]', ' ', m.group(1)).strip()
        return fecha, m.group(2).strip()
    return None, None

def procesar_imagen(img_pil):
    """Prueba múltiples franjas y vota por la coordenada más frecuente."""
    lats, lons, fechas, horas = [], [], [], []
    for franja in [0.05, 0.06, 0.07, 0.08, 0.09, 0.10, 0.12]:
        proc = aislar_texto_azul(img_pil, franja)
        texto = pytesseract.image_to_string(proc, lang='eng', config='--psm 6 --oem 3')
        for linea in texto.split('\n'):
            lats.append(extraer_lat(linea))
            lons.append(extraer_lon(linea))
        f, h = parsear_fecha_hora(texto)
        if f: fechas.append(f)
        if h: horas.append(h)
    lat = votar(lats)
    lon = votar(lons)
    fecha = Counter(fechas).most_common(1)[0][0] if fechas else None
    hora  = Counter(horas).most_common(1)[0][0]  if horas  else None
    return lat, lon, fecha, hora

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
    st.caption("v1.2 · Extractor GPS Garmin")

archivos = st.file_uploader(
    "📤 Arrastra tus fotos aquí o haz clic para seleccionar",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if not archivos:
    st.info("👆 Sube tus fotos para comenzar.")
    st.stop()

st.divider()

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
    progress = st.progress(0, text="Iniciando...")
    status   = st.empty()

    encabezados_cols = st.columns([0.4, 2.2, 1.4, 1.4, 1.4, 1.4, 0.5])
    for col, lbl in zip(encabezados_cols, ['#', 'Archivo', 'Fecha', 'Hora', 'Latitud', 'Longitud', '✓']):
        col.markdown(f"**{lbl}**")
    st.markdown("---")

    for idx, archivo in enumerate(archivos):
        status.info(f"⏳ Procesando **{archivo.name}** ({idx+1}/{len(archivos)})")
        try:
            archivo.seek(0)
            img = Image.open(archivo)
            lat, lon, fecha, hora = procesar_imagen(img)
            if lat is not None and lon is not None:
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

        progress.progress((idx + 1) / len(archivos), text=f"Procesando {idx+1}/{len(archivos)}...")

    status.empty(); progress.empty()
    st.divider()

    if ok_count > 0:
        fail = len(archivos) - ok_count
        st.success(f"✅ {ok_count} foto(s) procesadas · {fail} sin detectar")
        st.download_button(
            label="📥 Descargar Excel",
            data=generar_excel(datos),
            file_name="coordenadas_gps.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.error("❌ No se detectaron coordenadas. Verifica que las fotos tengan marca de agua azul.")

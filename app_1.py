"""
app.py  —  Procesador de Pedidos Darnel
========================================
Para correr localmente:
    pip install streamlit pandas openpyxl
    streamlit run app.py

Para deployar gratis:
    1. Subir app.py y requirements.txt a un repositorio de GitHub
    2. Ir a https://share.streamlit.io y conectar el repo
"""

import io
from pathlib import Path

import pandas as pd
import openpyxl
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Configuración ──────────────────────────────────────────────────────────────
HOJA_ABM     = "ABM ART. DEPURADO 23.02.2025"
HOJA_RESUMEN = "RESUMEN"
COL_CODIGO_DARNEL = 0
COL_CANT_EMPAQUE  = 17

ZONE_BG = {1: '1F4E79', 2: '2E75B6', 3: '0070C0',
           4: '00B0F0', 5: '9DC3E6', 6: 'BDD7EE'}
ZONE_FG = {1: 'FFFFFF', 2: 'FFFFFF', 3: 'FFFFFF',
           4: '1F3864', 5: '1F3864', 6: '1F3864'}

HEADERS = ['Zona', 'Código Darnel', 'Código Pilarica', 'Descripción',
           'Cant x Un Empaque', 'Uds/Pallet', 'Pallets']
COL_W   = [8, 18, 17, 38, 18, 13, 13]


# ══════════════════════════════════════════════════════════════════════════════
# LÓGICA
# ══════════════════════════════════════════════════════════════════════════════

def leer_pedido(file) -> list:
    """
    Detecta automáticamente todos los bloques de productos en el pedido.
    Cada bloque empieza con 'ID ARTICULO' en col 0 y termina con 'TOTAL UM:'.
    Funciona tanto para pedidos de un solo bloque como de múltiples bloques.
    """
    wb   = load_workbook(file, read_only=True, data_only=True)
    rows = list(wb.active.iter_rows(values_only=True))
    wb.close()

    header_rows = [i for i, r in enumerate(rows) if r[0] == 'ID ARTICULO']
    total_rows  = [i for i, r in enumerate(rows)
                   if r[0] and str(r[0]).startswith('TOTAL UM:')]

    if not header_rows or not total_rows:
        raise ValueError("No se encontraron bloques de productos en el pedido.")

    pedido = []
    for h, t in zip(header_rows, total_rows):
        for row in rows[h + 2 : t]:
            cod = row[COL_CODIGO_DARNEL]
            if not cod or not str(cod).strip():
                continue
            try:
                cant = int(str(row[COL_CANT_EMPAQUE]).replace(',', '').replace('.', '').strip())
            except (ValueError, TypeError, AttributeError):
                cant = None
            pedido.append({'cod_darnel': str(cod).strip(), 'cant_empaque': cant})

    return pedido


def leer_catalogo(file) -> tuple:
    df_abm = pd.read_excel(file, sheet_name=HOJA_ABM, header=2)
    df_abm['cd'] = df_abm['Articulo Formularios'].astype(str).str.strip()
    df_abm['cp'] = df_abm['Articulo'].astype(str).str.strip()
    df_abm['nm'] = df_abm['Nombre Articulo'].astype(str).str.strip()
    mapa = (df_abm.drop_duplicates('cd')
                  .set_index('cd')[['cp', 'nm']]
                  .to_dict('index'))

    df_res = pd.read_excel(file, sheet_name=HOJA_RESUMEN, header=2)
    df_res['ci']   = df_res['Código'].astype(str).str.strip()
    df_res['zona'] = pd.to_numeric(df_res['Zona'], errors='coerce')
    df_res['cpal'] = pd.to_numeric(df_res['Cant por Pallet'], errors='coerce')
    catalogo = (df_res.drop_duplicates('ci')
                      .set_index('ci')[['zona', 'cpal']]
                      .to_dict('index'))
    return mapa, catalogo


def cruzar(pedido, mapa, catalogo) -> tuple:
    resultado, sin_match = [], []
    for p in pedido:
        cd  = p['cod_darnel']
        abm = mapa.get(cd)
        if not abm:
            sin_match.append(cd); continue
        cod_pilarica = abm['cp']
        nombre       = abm['nm']
        cat          = catalogo.get(cod_pilarica)
        if not cat:
            sin_match.append(cd); continue
        zona     = cat['zona']
        cant_pal = cat['cpal']
        cant_emp = p['cant_empaque']
        pallets  = round(cant_emp / cant_pal, 2) if (cant_pal and cant_pal > 0 and cant_emp) else None
        resultado.append({
            'zona': zona, 'cod_darnel': cd, 'cod_pilarica': cod_pilarica,
            'descripcion': nombre, 'cant_empaque': cant_emp,
            'cant_pallet': int(cant_pal) if cant_pal else None, 'pallets': pallets,
        })
    resultado.sort(key=lambda x: (x['zona'] or 0, x['cod_darnel']))
    return resultado, sin_match


# ══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════════════════════

_thin = Side(style='thin', color='B0C4DE')
_BRD  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

def _s(cell, bold=False, italic=False, bg=None, fg='1F3864', align='left', size=11):
    cell.font      = Font(bold=bold, italic=italic, color=fg, name='Arial', size=size)
    cell.fill      = PatternFill('solid', start_color=bg) if bg else PatternFill()
    cell.alignment = Alignment(horizontal=align, vertical='center')
    cell.border    = _BRD

def _set_widths(ws):
    for i, w in enumerate(COL_W, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

def _escribir_encabezados(ws, row, bgz, fgz):
    ws.row_dimensions[row].height = 22
    for col, h in enumerate(HEADERS, 1):
        _s(ws.cell(row, col, h), bold=True, bg=bgz, fg=fgz, align='center')

def _escribir_fila(ws, ri, r, i):
    bg = 'EBF3FB' if i % 2 == 0 else 'FFFFFF'
    ws.row_dimensions[ri].height = 17
    _s(ws.cell(ri, 1, int(r['zona'])),       bg=bg, align='center')
    _s(ws.cell(ri, 2, r['cod_darnel']),       bg=bg)
    _s(ws.cell(ri, 3, r['cod_pilarica']),     bg=bg)
    _s(ws.cell(ri, 4, r['descripcion']),      bg=bg)
    c5 = ws.cell(ri, 5, r['cant_empaque']);   _s(c5, bg=bg, align='right'); c5.number_format = '#,##0'
    c6 = ws.cell(ri, 6, r['cant_pallet']);    _s(c6, bg=bg, align='right'); c6.number_format = '#,##0'
    c7 = ws.cell(ri, 7, r['pallets']);        _s(c7, bg=bg, align='right'); c7.number_format = '#,##0.00'

def _escribir_total(ws, ri, first, last, label='TOTAL'):
    ws.row_dimensions[ri].height = 19
    for col in range(1, 8):
        _s(ws.cell(ri, col, ''), bold=True, bg='DDEBF7')
    _s(ws.cell(ri, 4, label), bold=True, bg='DDEBF7', align='right')
    ct = ws.cell(ri, 5, f'=SUM(E{first}:E{last})')
    _s(ct, bold=True, bg='DDEBF7', align='right'); ct.number_format = '#,##0'
    cp = ws.cell(ri, 7, f'=SUM(G{first}:G{last})')
    _s(cp, bold=True, bg='DDEBF7', align='right'); cp.number_format = '#,##0.00'

def generar_excel(resultado, nombre_pedido) -> bytes:
    wb    = openpyxl.Workbook()
    zonas = sorted(set(r['zona'] for r in resultado))

    ws = wb.active
    ws.title = 'Resumen General'
    ws.sheet_view.showGridLines = False
    _set_widths(ws)

    ws.row_dimensions[1].height = 36
    ws.merge_cells('A1:G1')
    _s(ws['A1'], bold=True, bg='1F4E79', fg='FFFFFF', align='center', size=15)
    ws['A1'] = f'PEDIDO {nombre_pedido}  ·  RECEPCIÓN POR ZONA'

    ws.row_dimensions[2].height = 18
    ws.merge_cells('A2:G2')
    _s(ws['A2'], italic=True, bg='2E75B6', fg='FFFFFF', align='center', size=10)
    ws['A2'] = 'Agrupado por Zona'

    _escribir_encabezados(ws, 3, '2E75B6', 'FFFFFF')

    ri = 4
    for zona in zonas:
        zi    = int(zona)
        items = [r for r in resultado if r['zona'] == zona]
        bgz   = ZONE_BG.get(zi, '2E75B6')
        fgz   = ZONE_FG.get(zi, 'FFFFFF')

        ws.row_dimensions[ri].height = 20
        ws.merge_cells(f'A{ri}:G{ri}')
        _s(ws.cell(ri, 1, f'  ▶  ZONA  {zi}'), bold=True, bg=bgz, fg=fgz, align='left', size=12)
        ri += 1

        first = ri
        for i, r in enumerate(items):
            _escribir_fila(ws, ri, r, i); ri += 1

        _escribir_total(ws, ri, first, ri - 1, f'TOTAL ZONA {zi}')
        ri += 1

    ws.row_dimensions[ri].height = 24
    for col in range(1, 8):
        _s(ws.cell(ri, col, ''), bold=True, bg='1F4E79', fg='FFFFFF')
    c = ws.cell(ri, 4, 'TOTAL GENERAL')
    c.font = Font(bold=True, color='FFFFFF', name='Arial', size=13)
    c.alignment = Alignment(horizontal='right', vertical='center'); c.border = _BRD
    for col, val, fmt in [
        (5, sum(r['cant_empaque'] for r in resultado if r['cant_empaque']), '#,##0'),
        (7, round(sum(r['pallets'] for r in resultado if r['pallets']), 2), '#,##0.00'),
    ]:
        cx = ws.cell(ri, col, val)
        cx.font = Font(bold=True, color='FFFFFF', name='Arial', size=12)
        cx.alignment = Alignment(horizontal='right', vertical='center')
        cx.border = _BRD; cx.number_format = fmt

    for zona in zonas:
        zi    = int(zona)
        items = [r for r in resultado if r['zona'] == zona]
        bgz   = ZONE_BG.get(zi, '2E75B6')
        fgz   = ZONE_FG.get(zi, 'FFFFFF')
        wz    = wb.create_sheet(f'Zona {zi}')
        wz.sheet_view.showGridLines = False
        _set_widths(wz)

        wz.row_dimensions[1].height = 34
        wz.merge_cells('A1:G1')
        _s(wz['A1'], bold=True, bg=bgz, fg=fgz, align='center', size=14)
        wz['A1'] = f'ZONA {zi}  —  PEDIDO {nombre_pedido}'

        _escribir_encabezados(wz, 2, bgz, fgz)
        for i, r in enumerate(items):
            _escribir_fila(wz, 3 + i, r, i)

        tr = 3 + len(items)
        _escribir_total(wz, tr, 3, tr - 1)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title="Procesador de Pedidos Darnel", page_icon="📦", layout="centered")

st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
        html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }
        .block-container { max-width: 780px; padding-top: 2.5rem; }
        h1 { font-family: 'IBM Plex Mono', monospace !important; font-size: 1.6rem !important; letter-spacing: -0.5px; }
        h3 { font-family: 'IBM Plex Mono', monospace !important; font-size: 1rem !important; color: #2E75B6; }
        .upload-label { font-family: 'IBM Plex Mono', monospace; font-size: 0.75rem; font-weight: 600;
            color: #1F4E79; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 0.4rem; }
        .stat-box { background: #1F4E79; color: white; border-radius: 8px; padding: 1rem 1.2rem;
            text-align: center; font-family: 'IBM Plex Mono', monospace; }
        .stat-num  { font-size: 2rem; font-weight: 600; line-height: 1; }
        .stat-label { font-size: 0.7rem; opacity: 0.7; margin-top: 4px; letter-spacing: 1px; text-transform: uppercase; }
        .zone-row { display: flex; align-items: center; gap: 12px; padding: 8px 12px; border-radius: 6px;
            margin-bottom: 6px; background: #f8fbff; border-left: 4px solid #2E75B6;
            font-family: 'IBM Plex Mono', monospace; font-size: 0.85rem; }
        .zone-badge { background: #1F4E79; color: white; border-radius: 4px; padding: 2px 8px;
            font-weight: 600; font-size: 0.75rem; }
        .divider { height: 1px; background: #e0eaf5; margin: 1.5rem 0; }
        .warn-box { background: #fff8e1; border-left: 4px solid #f0a500; padding: 0.8rem 1rem;
            border-radius: 6px; font-size: 0.85rem; margin-top: 0.5rem; }
        .info-pill { display:inline-block; background:#e8f4fd; color:#1F4E79; border-radius:4px;
            padding:2px 8px; font-family:'IBM Plex Mono',monospace; font-size:0.75rem; margin-left:8px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("# 📦 Procesador de Pedidos Darnel")
st.markdown("Cruzá el pedido con el catálogo y descargá el reporte agrupado por zona.")
st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="upload-label">① Pedido Darnel</div>', unsafe_allow_html=True)
    pedido_file = st.file_uploader("", type=["xlsx"], key="pedido", label_visibility="collapsed")
    if pedido_file:
        st.success(f"✓ {pedido_file.name}")

with col2:
    st.markdown('<div class="upload-label">② Catálogo Pilarica (Stock/Zonas)</div>', unsafe_allow_html=True)
    catalogo_file = st.file_uploader("", type=["xlsx"], key="catalogo", label_visibility="collapsed")
    if catalogo_file:
        st.success(f"✓ {catalogo_file.name}")

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

if pedido_file and catalogo_file:
    if st.button("⚡ Procesar", use_container_width=True, type="primary"):
        with st.spinner("Procesando..."):
            try:
                pedido               = leer_pedido(pedido_file)
                mapa, catalogo       = leer_catalogo(catalogo_file)
                resultado, sin_match = cruzar(pedido, mapa, catalogo)

                zonas     = sorted(set(r['zona'] for r in resultado))
                total_u   = sum(r['cant_empaque'] for r in resultado if r['cant_empaque'])
                total_pal = round(sum(r['pallets'] for r in resultado if r['pallets']), 2)
                nombre_ped = Path(pedido_file.name).stem

                # Métricas
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(resultado)}</div><div class="stat-label">Productos</div></div>', unsafe_allow_html=True)
                with c2:
                    st.markdown(f'<div class="stat-box"><div class="stat-num">{len(zonas)}</div><div class="stat-label">Zonas</div></div>', unsafe_allow_html=True)
                with c3:
                    st.markdown(f'<div class="stat-box"><div class="stat-num">{total_pal:,.1f}</div><div class="stat-label">Pallets totales</div></div>', unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown("### Resumen por zona")

                for zona in zonas:
                    items  = [r for r in resultado if r['zona'] == zona]
                    zona_u = sum(r['cant_empaque'] for r in items if r['cant_empaque'])
                    zona_p = round(sum(r['pallets'] for r in items if r['pallets']), 2)
                    st.markdown(
                        f'<div class="zone-row">'
                        f'<span class="zone-badge">ZONA {int(zona)}</span>'
                        f'<span>{len(items)} productos</span>'
                        f'<span style="margin-left:auto;opacity:.6">{zona_u:,} uds · {zona_p} pallets</span>'
                        f'</div>',
                        unsafe_allow_html=True
                    )

                if sin_match:
                    st.markdown(
                        f'<div class="warn-box">⚠️ <b>{len(sin_match)} producto(s)</b> no encontrados en el catálogo: '
                        + ", ".join(sin_match) + "</div>",
                        unsafe_allow_html=True
                    )

                st.markdown("<br>", unsafe_allow_html=True)
                excel_bytes = generar_excel(resultado, nombre_ped)
                st.download_button(
                    label="📥 Descargar Excel",
                    data=excel_bytes,
                    file_name=f"reporte_por_zona_{nombre_ped}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"❌ Error al procesar: {e}")
                st.exception(e)
else:
    st.info("Subí los dos archivos para habilitar el procesamiento.")

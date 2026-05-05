"""
Servidor Flask para generar protocolos de calidad Membrantec Chile.

Rellena las plantillas XLSX oficiales modificando SOLO el XML de la hoja
(sheet1.xml) dentro del ZIP del XLSX. Esto preserva al 100%:
  - Logos (PNG y EMF/WMF)
  - Bordes, fuentes, colores
  - Celdas fusionadas
  - Firmas, encabezados, pies de página
  - Todo el formato original

Plantillas esperadas (junto al script, o en subcarpeta "Aplicación Membrantec"):
  - GM-PI-HDPE-INF-RAI-001.xlsx    (Pruebas Iniciales)
  - GM-PD-HDPE-INF-RAI-001.xlsx    (Pruebas Destructivas)
  - GM-RU-HDPE-INF-RAI-001.xlsx    (Uniones y Prueba de Aire)
  - GM-REP-HDPE-SUP-RAI-001_1.xlsx (Reparaciones y Vacío)
  - GM-RI-HDPE-SUP-RAI-001.xlsx    (Instalación)
"""
import os
import re
import shutil
import tempfile
import zipfile
from io import BytesIO
from xml.etree import ElementTree as ET

from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

# ═══════════════════════════════════════════════════════════════════
# CONFIG DE PLANTILLAS
# ═══════════════════════════════════════════════════════════════════

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))

PLANTILLAS = {
    "pi":  "GM-PI-HDPE-INF-RAI-001.xlsx",
    "pd":  "GM-PD-HDPE-INF-RAI-001.xlsx",
    "ru":  "GM-RU-HDPE-INF-RAI-001.xlsx",
    "rep": "GM-REP-HDPE-SUP-RAI-001_1.xlsx",
    "ri":  "GM-RI-HDPE-SUP-RAI-001.xlsx",
}


def _encontrar_plantillas_dir():
    test_file = PLANTILLAS["pi"]
    if os.path.isfile(os.path.join(_BASE_DIR, test_file)):
        return _BASE_DIR
    for nombre in ["Aplicación Membrantec", "Aplicacion Membrantec",
                   "App Membrantec", "plantillas", "templates"]:
        d = os.path.join(_BASE_DIR, nombre)
        if os.path.isfile(os.path.join(d, test_file)):
            return d
    try:
        for sub in os.listdir(_BASE_DIR):
            d = os.path.join(_BASE_DIR, sub)
            if os.path.isdir(d) and os.path.isfile(os.path.join(d, test_file)):
                return d
    except Exception:
        pass
    return _BASE_DIR


PLANTILLAS_DIR = _encontrar_plantillas_dir()
print(f"[membrantec] Plantillas en: {PLANTILLAS_DIR}")


# ═══════════════════════════════════════════════════════════════════
# MOTOR DE PARCHE XLSX  (modifica sheet1.xml dentro del ZIP)
# ═══════════════════════════════════════════════════════════════════

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', NS)
# Registrar prefijos OOXML para que Excel no los renombre a ns1/ns2/ns3 y
# rompa la referencia mc:Ignorable="x14ac xr xr2 xr3"
ET.register_namespace('r',     'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
ET.register_namespace('mc',    'http://schemas.openxmlformats.org/markup-compatibility/2006')
ET.register_namespace('x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac')
ET.register_namespace('xr',    'http://schemas.microsoft.com/office/spreadsheetml/2014/revision')
ET.register_namespace('xr2',   'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2')
ET.register_namespace('xr3',   'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3')


def _col_to_num(col):
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n


def _parse_coord(coord):
    m = re.match(r'([A-Z]+)(\d+)', coord)
    return _col_to_num(m.group(1)), int(m.group(2))


def _parsear_merges(root):
    merges = []
    mc = root.find(f'{{{NS}}}mergeCells')
    if mc is None:
        return merges
    for m in mc.findall(f'{{{NS}}}mergeCell'):
        ref = m.get('ref')
        a, b = ref.split(':')
        ac, ar = _parse_coord(a)
        bc, br = _parse_coord(b)
        merges.append((ac, ar, bc, br, a))
    return merges


def _top_left_de_merge(coord, merges):
    cc, cr = _parse_coord(coord)
    for ac, ar, bc, br, top in merges:
        if ac <= cc <= bc and ar <= cr <= br:
            return top
    return None


def _escribir_valor(cell_el, value):
    for child in list(cell_el):
        cell_el.remove(child)
    if 't' in cell_el.attrib:
        del cell_el.attrib['t']
    if isinstance(value, bool):
        value = str(value)
    if isinstance(value, (int, float)):
        cell_el.set('t', 'n')
        v = ET.SubElement(cell_el, f'{{{NS}}}v')
        v.text = str(value)
    else:
        cell_el.set('t', 'inlineStr')
        is_el = ET.SubElement(cell_el, f'{{{NS}}}is')
        t_el = ET.SubElement(is_el, f'{{{NS}}}t')
        t_el.text = str(value)
        t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')


class XlsxPatcher:
    def __init__(self, src_path, sheet_xml_path='xl/worksheets/sheet1.xml'):
        self.src_path = src_path
        self.sheet_xml_path = sheet_xml_path
        self._patches = {}

    def set(self, coord, value):
        if value is None or value == "":
            return
        self._patches[coord] = value

    def set_map(self, mapa, fuente):
        for campo, coord in mapa.items():
            val = fuente.get(campo, "")
            self.set(coord, val)

    def save(self, dst_path):
        shutil.copy(self.src_path, dst_path)
        with zipfile.ZipFile(dst_path, 'r') as z:
            sheet_xml = z.read(self.sheet_xml_path).decode('utf-8')
        root = ET.fromstring(sheet_xml)
        sheetData = root.find(f'{{{NS}}}sheetData')
        merges = _parsear_merges(root)
        rows_by_num = {}
        for row_el in sheetData.findall(f'{{{NS}}}row'):
            rows_by_num[int(row_el.get('r'))] = row_el
        for coord, value in self._patches.items():
            top = _top_left_de_merge(coord, merges)
            if top:
                coord = top
            m = re.match(r'([A-Z]+)(\d+)', coord)
            rownum = int(m.group(2))
            row_el = rows_by_num.get(rownum)
            if row_el is None:
                row_el = ET.SubElement(sheetData, f'{{{NS}}}row')
                row_el.set('r', str(rownum))
                rows_by_num[rownum] = row_el
            cell_el = None
            for c in row_el.findall(f'{{{NS}}}c'):
                if c.get('r') == coord:
                    cell_el = c
                    break
            if cell_el is None:
                cell_el = ET.SubElement(row_el, f'{{{NS}}}c')
                cell_el.set('r', coord)
            _escribir_valor(cell_el, value)

        # Reordenar para que Excel no muestre warning de "contenido no valido".
        # Excel exige <row> en orden creciente por numero de fila, y dentro de
        # cada fila <c> en orden creciente por columna. Como ET.SubElement los
        # agrega al final, hay que reordenar antes de serializar.
        rows_sorted = sorted(
            sheetData.findall(f'{{{NS}}}row'),
            key=lambda r: int(r.get('r'))
        )
        for r in list(sheetData):
            sheetData.remove(r)
        for r in rows_sorted:
            cells_sorted = sorted(
                r.findall(f'{{{NS}}}c'),
                key=lambda c: _col_to_num(re.match(r'([A-Z]+)', c.get('r')).group(1))
            )
            for c in list(r):
                r.remove(c)
            for c in cells_sorted:
                r.append(c)
            if cells_sorted:
                first_col = _col_to_num(re.match(r'([A-Z]+)', cells_sorted[0].get('r')).group(1))
                last_col  = _col_to_num(re.match(r'([A-Z]+)', cells_sorted[-1].get('r')).group(1))
                r.set('spans', f'{first_col}:{last_col}')
            sheetData.append(r)

        new_xml = ET.tostring(root, encoding='utf-8', xml_declaration=True)
        tmp = dst_path + '.tmp'
        with zipfile.ZipFile(dst_path, 'r') as zin:
            with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = new_xml if item.filename == self.sheet_xml_path else zin.read(item.filename)
                    zout.writestr(item, data)
        os.replace(tmp, dst_path)


# ═══════════════════════════════════════════════════════════════════
# MARCAS DE ESCALA / TIPO  (helpers)
# ═══════════════════════════════════════════════════════════════════

def _marcar_escala(patch, escala, lb_cell="D26", kg_cell="E26"):
    """Marca con (X) la escala usada en el tensiómetro y deja la otra plana."""
    e = (escala or "Lb/in").strip().lower()
    if e.startswith("lb"):
        patch.set(lb_cell, "(X) Lb/in"); patch.set(kg_cell, "Kg/in")
    else:
        patch.set(lb_cell, "Lb/in"); patch.set(kg_cell, "(X) Kg/in")


def _marcar_tipo_soldadura(patch, tipo, fus_cell="F22", ext_cell="F23"):
    t = (tipo or "fusion").lower()
    if t.startswith("fus"):
        patch.set(fus_cell, "X"); patch.set(ext_cell, "-")
    else:
        patch.set(fus_cell, "-"); patch.set(ext_cell, "X")


# ═══════════════════════════════════════════════════════════════════
# GENERADORES POR TIPO DE PROTOCOLO
# ═══════════════════════════════════════════════════════════════════

def generar_pi(patch, proyecto, registros):
    """Pruebas Iniciales. Tabla en filas 36+."""
    patch.set_map({
        "contrato":   "F4",
        "proyecto":   "F5",
        "area_cod":   "D7",
        "plano":      "J7",
        "protocolo":  "S7",
        "fecha":      "AB7",
        "tag":        "A

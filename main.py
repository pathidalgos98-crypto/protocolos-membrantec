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
        "tag":        "A9",
        "area_cod2":  "C10",
        "area_desc":  "E10",
        "sis_cod":    "H10",
        "sis_desc":   "K10",
        "sub_cod":    "M10",
        "sub_desc":   "P10",
        "pie":        "S10",
        "aconex":     "Y10",
        "lugar":      "F12",
        "espesor":    "B17",
        "min_des":    "C25",
        "min_cor":    "E25",
        "tens_eq":    "C28",
        "tens_cert":  "C29",
    }, proyecto)

    _marcar_tipo_soldadura(patch, proyecto.get("tipo_prueba"), "F22", "F23")
    _marcar_escala(patch, proyecto.get("escala"), "D26", "E26")

    fila = 36
    for r in registros:
        patch.set(f"A{fila}",  r.get("fecha"))
        patch.set(f"C{fila}",  r.get("hora"))
        patch.set(f"D{fila}",  r.get("operador"))
        patch.set(f"E{fila}",  r.get("tamb"))
        patch.set(f"F{fila}",  r.get("maq"))
        patch.set(f"H{fila}",  r.get("tmaq"))
        patch.set(f"J{fila}",  r.get("tleister") or "-")
        patch.set(f"K{fila}",  r.get("veloc"))
        patch.set(f"L{fila}",  r.get("inspector"))
        for letra, key in zip(["N","O","P","Q","R","S","T","U","V","W"],
                              ["d1a","d1b","d2a","d2b","d3a","d3b","d4a","d4b","d5a","d5b"]):
            patch.set(f"{letra}{fila}", r.get(key))
        for letra, key in zip(["X","Y","Z","AA","AB"], ["c1","c2","c3","c4","c5"]):
            patch.set(f"{letra}{fila}", r.get(key))
        patch.set(f"AC{fila}", r.get("resultado") or "Aprobado")
        patch.set(f"AD{fila}", r.get("falla_prob") or "-")
        patch.set(f"AE{fila}", r.get("falla_tipo") or "-")
        fila += 1


def generar_pd(patch, proyecto, registros):
    """Pruebas Destructivas. Tabla en filas 35+."""
    patch.set_map({
        "contrato":  "E4",
        "proyecto":  "E5",
        "area_cod":  "D7",
        "plano":     "J7",
        "protocolo": "S7",
        "fecha":     "AB7",
        "tag":       "A10",
        "area_cod2": "D10",
        "area_desc": "F10",
        "sis_cod":   "I10",
        "sis_desc":  "L10",
        "sub_cod":   "P10",
        "sub_desc":  "S10",
        "pie":       "V10",
        "aconex":    "AB10",
        "lugar":     "F12",
        "espesor":   "H15",
        "min_des":   "C25",
        "min_cor":   "E25",
        "tens_eq":   "D27",
        "tens_cert": "J27",
    }, proyecto)

    _marcar_tipo_soldadura(patch, proyecto.get("tipo_prueba"), "F22", "F23")
    _marcar_escala(patch, proyecto.get("escala"), "D26", "E26")

    fila = 35
    for i, r in enumerate(registros, 1):
        patch.set(f"A{fila}",  r.get("fecha"))
        patch.set(f"B{fila}",  i)
        patch.set(f"C{fila}",  r.get("union"))
        patch.set(f"D{fila}",  r.get("ubicacion"))
        patch.set(f"E{fila}",  r.get("hora"))
        patch.set(f"F{fila}",  r.get("tamb"))
        patch.set(f"G{fila}",  r.get("operador"))
        patch.set(f"H{fila}",  r.get("maq"))
        patch.set(f"I{fila}",  r.get("tmaq"))
        patch.set(f"J{fila}",  r.get("tleister") or "-")
        patch.set(f"K{fila}",  r.get("veloc"))
        patch.set(f"L{fila}",  r.get("inspector"))
        for letra, key in zip(["N","O","P","Q","R","S","T","U","V","W"],
                              ["d1a","d1b","d2a","d2b","d3a","d3b","d4a","d4b","d5a","d5b"]):
            patch.set(f"{letra}{fila}", r.get(key))
        for letra, key in zip(["X","Y","Z","AA","AB"], ["c1","c2","c3","c4","c5"]):
            patch.set(f"{letra}{fila}", r.get(key))
        patch.set(f"AC{fila}", r.get("resultado") or "Aprobado")
        patch.set(f"AD{fila}", r.get("falla_prob") or "-")
        patch.set(f"AE{fila}", r.get("falla_tipo") or "-")
        fila += 1


def generar_ru(patch, proyecto, registros, manometro=None):
    """Uniones y Prueba de Aire. Tabla en filas 26+."""
    patch.set_map({
        "contrato":  "E3",
        "proyecto":  "E4",
        "area_cod":  "D6",
        "plano":     "H6",
        "protocolo": "P6",
        "fecha":     "V6",
        "tag":       "B8",
        "area_cod2": "D9",
        "area_desc": "E9",
        "sis_cod":   "H9",
        "sis_desc":  "K9",
        "sub_cod":   "O9",
        "sub_desc":  "Q9",
        "pie":       "S9",
        "aconex":    "U9",
        "lugar":     "E12",
        "espesor":   "C15",
    }, proyecto)

    fila = 26
    for r in registros:
        patch.set(f"B{fila}", r.get("union"))
        patch.set(f"C{fila}", r.get("distancia"))
        patch.set(f"D{fila}", r.get("fecha"))
        patch.set(f"E{fila}", r.get("operador"))
        patch.set(f"F{fila}", r.get("maq"))
        patch.set(f"G{fila}", r.get("tmaq"))
        patch.set(f"H{fila}", r.get("veloc"))
        patch.set(f"I{fila}", r.get("hora_union"))
        linea = (r.get("linea") or "SI").upper()
        if linea == "SI":
            patch.set(f"J{fila}", "X"); patch.set(f"K{fila}", "-")
        else:
            patch.set(f"J{fila}", "-"); patch.set(f"K{fila}", "X")
        patch.set(f"L{fila}", r.get("ds"))
        patch.set(f"N{fila}", r.get("hora_ini"))
        patch.set(f"O{fila}", r.get("psi_ini"))
        patch.set(f"P{fila}", r.get("hora_fin"))
        patch.set(f"Q{fila}", r.get("psi_fin"))
        patch.set(f"R{fila}", r.get("psi_dif"))
        patch.set(f"S{fila}", r.get("resultado") or "Aprobado")
        patch.set(f"T{fila}", r.get("tec"))
        patch.set(f"U{fila}", r.get("fecha_prueba"))
        patch.set(f"V{fila}", r.get("tipo_falla") or "-")
        patch.set(f"W{fila}", r.get("mano_etiq"))
        fila += 1

    if manometro:
        patch.set("N39", manometro.get("serie"))
        patch.set("Q39", manometro.get("fecha_cal"))
        patch.set("T39", manometro.get("certificado"))
        patch.set("V39", manometro.get("etiqueta"))


def generar_rep(patch, proyecto, registros, equipo=None):
    """Reparaciones y Vacío/Chispa/Lanza. Tabla en filas 27+."""
    patch.set_map({
        "contrato":  "G3",
        "proyecto":  "G4",
        "area_cod":  "D6",
        "plano":     "I6",
        "protocolo": "Q6",
        "fecha":     "X6",
        "tag":       "B9",
        "area_cod2": "D9",
        "area_desc": "E9",
        "sis_cod":   "G9",
        "sis_desc":  "H9",
        "sub_cod":   "J9",
        "sub_desc":  "L9",
        "pie":       "N9",
        "aconex":    "U9",
        "lugar":     "E11",
        "espesor":   "C14",
    }, proyecto)

    tipo = (proyecto.get("tipo_prueba_rep") or "vacio").lower()
    if tipo == "vacio":
        patch.set("I12", "X"); patch.set("I14", "-")
    elif tipo == "spark":
        patch.set("I12", "-"); patch.set("I14", "X")

    if equipo:
        patch.set("G18", equipo.get("nombre"))
        patch.set("H18", equipo.get("serie") or "S/I")
        patch.set("J18", equipo.get("fecha_cal"))
        patch.set("L18", equipo.get("certificado"))
        patch.set("N18", equipo.get("etiqueta") or "-")

    fila = 27
    for i, r in enumerate(registros, 1):
        patch.set(f"B{fila}", r.get("fecha"))
        patch.set(f"C{fila}", r.get("hora"))
        patch.set(f"D{fila}", i)
        patch.set(f"E{fila}", r.get("tipo_rep"))
        patch.set(f"F{fila}", r.get("union"))
        patch.set(f"H{fila}", r.get("tecnico"))
        patch.set(f"J{fila}", r.get("dim_largo"))
        patch.set(f"K{fila}", r.get("dim_ancho"))
        patch.set(f"L{fila}", r.get("maq"))
        patch.set(f"M{fila}", r.get("tmaq"))
        patch.set(f"N{fila}", r.get("tleister"))
        patch.set(f"O{fila}", r.get("ds") or "-")
        patch.set(f"Q{fila}", r.get("prueba") or "VA")
        patch.set(f"S{fila}", "OK" if r.get("resultado") == "Aprobado" else "X")
        patch.set(f"T{fila}", r.get("tipo_falla1") or "-")
        patch.set(f"U{fila}", r.get("tipo_falla2") or "-")
        patch.set(f"V{fila}", r.get("equipo") or "VAC-26")
        patch.set(f"W{fila}", r.get("hora_fin"))
        patch.set(f"X{fila}", r.get("fecha_fin"))
        patch.set(f"Y{fila}", r.get("op_fin"))
        fila += 1


def generar_ri(patch, proyecto, registros):
    """Instalación de Geomembranas. Tabla en filas 22+."""
    patch.set_map({
        "contrato":  "G4",
        "proyecto":  "G5",
        "area_cod":  "E7",
        "plano":     "M7",
        "protocolo": "U7",
        "fecha":     "AB7",
        "tag":       "B9",
        "area_cod2": "E10",
        "area_desc": "I10",
        "sis_cod":   "L10",
        "sis_desc":  "O10",
        "sub_cod":   "R10",
        "sub_desc":  "T10",
        "pie":       "V10",
        "aconex":    "AB10",
        "lugar":     "B15",
    }, proyecto)

    fila = 22
    for r in registros:
        patch.set(f"B{fila}", r.get("fecha"))
        patch.set(f"D{fila}", r.get("hora"))
        # Clima: nublado/despejado c-viento/despejado s-viento
        clima = (r.get("clima") or "").lower()
        if "nub" in clima:
            patch.set(f"E{fila}", "X")
        elif "viento" in clima and "c/" in clima:
            patch.set(f"F{fila}", "X")
        else:
            patch.set(f"G{fila}", "X")
        patch.set(f"I{fila}", r.get("panel"))
        patch.set(f"K{fila}", r.get("rollo"))
        patch.set(f"M{fila}", r.get("m2_rollo"))
        patch.set(f"R{fila}", r.get("espesor"))
        patch.set(f"S{fila}", r.get("st_largo"))
        patch.set(f"V{fila}", r.get("st_ancho"))
        patch.set(f"W{fila}", r.get("st_m2"))
        patch.set(f"Y{fila}", r.get("ct_largo"))
        patch.set(f"AA{fila}", r.get("ct_ancho"))
        patch.set(f"AC{fila}", r.get("ct_m2"))
        patch.set(f"AE{fila}", r.get("obs") or "-")
        fila += 1


GENERADORES = {
    "pi":  generar_pi,
    "pd":  generar_pd,
    "ru":  generar_ru,
    "rep": generar_rep,
    "ri":  generar_ri,
}


# ===========================================================
#                       RUTAS FLASK
# ===========================================================

@app.route("/", methods=["GET"])
def index():
    return jsonify({
        "service": "Membrantec - Generador de Protocolos",
        "status": "ok",
        "version": "3.0",
        "plantillas_dir": PLANTILLAS_DIR,
        "endpoints": ["GET /health", "POST /generar-protocolo"],
        "tipos": list(PLANTILLAS.keys()),
    })


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/generar-protocolo", methods=["POST"])
def generar_protocolo():
    """
    Body JSON:
      {
        "tipo": "pi|pd|ru|rep|ri",
        "proyecto": { ... },
        "registros": [ ... ],
        "manometro": { ... },     # opcional, solo RU
        "equipo":    { ... }      # opcional, solo REP
      }
    """
    try:
        data = request.get_json(force=True) or {}
    except Exception:
        return jsonify({"error": "JSON invalido"}), 400

    tipo = (data.get("tipo") or "").lower()
    if tipo not in PLANTILLAS:
        return jsonify({"error": "tipo desconocido: " + tipo}), 400

    plantilla = os.path.join(PLANTILLAS_DIR, PLANTILLAS[tipo])
    if not os.path.isfile(plantilla):
        # Intentar variantes (con/sin sufijo _1, etc.)
        base = PLANTILLAS[tipo]
        candidatos = [
            base,
            base.replace("_1.xlsx", ".xlsx"),
            base.replace(".xlsx", "_1.xlsx"),
            base.replace("-RAI-001", "-RAI-001_1") if "_1" not in base else base.replace("_1", ""),
        ]
        encontrado = None
        for c in candidatos:
            ruta = os.path.join(PLANTILLAS_DIR, c)
            if os.path.isfile(ruta):
                encontrado = ruta
                break
        if encontrado:
            plantilla = encontrado
        else:
            try:
                disponibles = [f for f in os.listdir(PLANTILLAS_DIR) if f.endswith(".xlsx")]
            except Exception:
                disponibles = []
            return jsonify({
                "error": "plantilla no encontrada: " + base,
                "plantillas_dir": PLANTILLAS_DIR,
                "candidatos_probados": candidatos,
                "disponibles": disponibles
            }), 500

    proyecto = data.get("proyecto") or {}
    registros = data.get("registros") or []

    if not registros:
        return jsonify({"error": "no hay registros para generar"}), 400

    fd, dst = tempfile.mkstemp(suffix=".xlsx", prefix=tipo + "_")
    os.close(fd)

    try:
        patch = XlsxPatcher(plantilla)
        gen = GENERADORES[tipo]
        if tipo == "ru":
            gen(patch, proyecto, registros, manometro=data.get("manometro"))
        elif tipo == "rep":
            gen(patch, proyecto, registros, equipo=data.get("equipo"))
        else:
            gen(patch, proyecto, registros)
        patch.save(dst)

        nombre = (proyecto.get("protocolo") or tipo.upper()) + ".xlsx"
        return send_file(
            dst,
            as_attachment=True,
            download_name=nombre,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

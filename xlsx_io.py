"""Lectura y escritura de archivos .xlsx usando solo la libreria estandar de Python.

No requiere openpyxl ni pandas: un .xlsx es un ZIP con XML adentro, y aca se
maneja directamente. Asi el script corre en cualquier maquina con Python 3.8+
sin instalar nada.
"""

import datetime
import re
import zipfile
import xml.etree.ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def _NOW():
    return datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

# Caracteres de control que Excel no acepta dentro de un valor de celda.
_CTRL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _col_letters(ref):
    """'BC12' -> 'BC'"""
    return "".join(ch for ch in ref if ch.isalpha())


def _col_index(letters):
    """'A' -> 0, 'B' -> 1, 'AA' -> 26"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _index_col(idx):
    """0 -> 'A', 26 -> 'AA'"""
    letters = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def read_xlsx(path, sheet_index=0):
    """Devuelve (encabezados, filas) donde cada fila es una lista de strings.

    Las celdas vacias se rellenan como '' para que todas las filas tengan el
    mismo largo que los encabezados.
    """
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in root:
                shared.append("".join(t.text or "" for t in si.iter(NS + "t")))

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = wb.find(NS + "sheets")
        if sheet_index >= len(sheets):
            raise ValueError("El archivo no tiene la hoja %d" % sheet_index)

        # Los nombres de hoja no siempre siguen el orden del zip: se resuelve
        # el rId contra workbook.xml.rels.
        rid = sheets[sheet_index].get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels:
            if rel.get("Id") == rid:
                target = rel.get("Target")
                break
        if target is None:
            target = "worksheets/sheet1.xml"
        target = target.lstrip("/")
        name = target if target.startswith("xl/") else "xl/" + target

        sheet = ET.fromstring(z.read(name))

    def cell_value(c):
        if c.get("t") == "inlineStr":
            is_el = c.find(NS + "is")
            return "".join(t.text or "" for t in is_el.iter(NS + "t")) if is_el is not None else ""
        v = c.find(NS + "v")
        if v is None or v.text is None:
            return ""
        if c.get("t") == "s":
            return shared[int(v.text)]
        return v.text

    raw = []
    width = 0
    for r in sheet.iter(NS + "row"):
        cells = {}
        for c in r:
            ref = c.get("r")
            if not ref:
                continue
            idx = _col_index(_col_letters(ref))
            cells[idx] = cell_value(c)
            width = max(width, idx + 1)
        raw.append(cells)

    if not raw:
        return [], []

    rows = [[cells.get(i, "") for i in range(width)] for cells in raw]
    headers = rows[0]
    # Recorta columnas fantasma a la derecha del ultimo encabezado con texto.
    last = 0
    for i, h in enumerate(headers):
        if str(h).strip():
            last = i + 1
    headers = headers[:last]
    data = [row[:last] + [""] * (last - len(row[:last])) for row in rows[1:]]
    # Descarta filas totalmente vacias.
    data = [row for row in data if any(str(v).strip() for v in row)]
    return headers, data


def _esc(text):
    text = _CTRL.sub("", str(text))
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


_NUMERIC = re.compile(r"^-?\d+(\.\d+)?$")


def write_xlsx(path, headers, rows, sheet_name="Hoja1", numeric_cols=()):
    """Escribe un .xlsx valido con encabezado en negrita, filtro y panel fijo.

    numeric_cols: indices de columna que deben guardarse como numero (y no
    como texto) cuando el valor lo permita.
    """
    numeric_cols = set(numeric_cols)
    n_rows = len(rows) + 1
    n_cols = len(headers)
    dim = "A1:%s%d" % (_index_col(max(n_cols - 1, 0)), n_rows)

    out = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="%s"/>' % dim,
        '<sheetViews><sheetView workbookViewId="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        "</sheetView></sheetViews>",
        '<sheetFormatPr defaultRowHeight="15"/>',
        "<cols>",
    ]
    widths = {0: 44, 1: 22, 2: 20, 3: 34, 4: 20}
    for i in range(n_cols):
        out.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (i + 1, i + 1, widths.get(i, 18)))
    out.append("</cols><sheetData>")

    out.append('<row r="1">')
    for i, h in enumerate(headers):
        out.append(
            '<c r="%s1" s="1" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (_index_col(i), _esc(h))
        )
    out.append("</row>")

    for ri, row in enumerate(rows, start=2):
        out.append('<row r="%d">' % ri)
        for ci in range(n_cols):
            val = row[ci] if ci < len(row) else ""
            if val is None or val == "":
                continue
            ref = "%s%d" % (_index_col(ci), ri)
            if ci in numeric_cols and _NUMERIC.match(str(val).strip()):
                out.append('<c r="%s"><v>%s</v></c>' % (ref, str(val).strip()))
            else:
                out.append(
                    '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (ref, _esc(val))
                )
        out.append("</row>")

    out.append("</sheetData>")
    out.append('<autoFilter ref="%s"/>' % dim)
    out.append("</worksheet>")
    sheet_xml = "".join(out)

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<workbookPr/>'
        '<bookViews><workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="12000"/></bookViews>'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets>'
        '<calcPr calcId="0"/></workbook>' % _esc(sheet_name[:31])
    )
    wb_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        "</Relationships>"
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="3">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FF1F3864"/><bgColor indexed="64"/></patternFill></fill>'
        "</fills>"
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    core = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:title>%s</dc:title>"
        '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>'
        "</cp:coreProperties>" % (_esc(sheet_name), _NOW(), _NOW())
    )
    app = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>aio_check.py</Application></Properties>"
    )

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("docProps/core.xml", core)
        z.writestr("docProps/app.xml", app)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", workbook)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/worksheets/sheet1.xml", sheet_xml)

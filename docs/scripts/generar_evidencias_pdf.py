# -*- coding: utf-8 -*-
"""Genera el PDF de evidencias con las 16 capturas reales, una por página.

Uso:
    python docs/scripts/generar_evidencias_pdf.py
    # luego exportar el DOCX a PDF con Word
"""

from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs"
CAP = DOCS / "capturas"
OUT = DOCS / "Evidencias_Capturas_S11_Edu_Galindo.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GRAY = RGBColor(0x47, 0x55, 0x69)
DARK = RGBColor(0x0F, 0x17, 0x2A)
GREEN = RGBColor(0x04, 0x78, 0x57)

# (n, archivo, sección, título, descripción, figura del informe)
EVID = [
    (1, "01_carga_esqueleto.png", "APLICACIÓN", "Estado de carga (loading)",
     "La bandeja muestra el esqueleto de carga y la cabecera el contador «…». La petición simulada tarda 900 ms; "
     "es el estado loading real, no un spinner genérico.", 7),
    (2, "02_planes_cargados.png", "APLICACIÓN", "Bandeja cargada (ready)",
     "Los seis planes reales (P-01 a P-06) y el contador «Pendientes: 6». Es el estado ready con los datos que "
     "devuelve fetchPlans.", 8),
    (3, "03_error_red.png", "APLICACIÓN", "Estado de error y reintento",
     "Con ?fail=1 la petición falla: se muestra el estado error, el aviso «No se pudieron cargar los planes» y el "
     "botón «Reintentar».", 9),
    (4, "04_busqueda_antigua.png", "APLICACIÓN", "Búsqueda antes del debounce",
     "El campo ya contiene «values», pero la lista sigue completa: useDebounce espera 300 ms de calma antes de "
     "aplicar el valor.", 10),
    (5, "05_busqueda_filtrada.png", "APLICACIÓN", "Búsqueda después del debounce",
     "Pasados los 300 ms, la lista se reduce a P-03: el valor «tranquilo» de useDebounce ya se usó para filtrar.", 11),
    (6, "06_contador_pendientes.png", "APLICACIÓN", "Contador de pendientes actualizado",
     "Tras aprobar un plan, la cabecera muestra «Pendientes: 5» y el panel «Decididos: 1». El contador se "
     "recalcula en App, no dentro de Header.", 12),
    (7, "07_ejecucion_r2_espera_aprobacion.png", "APLICACIÓN", "R2 pendiente de confirmación",
     "El plan R2 P-01 se ensayó, pero no se aplicó: el panel exige «Confirmar y aplicar» o «Descartar». La barrera "
     "de seguridad funciona.", 13),
    (8, "08_ejecucion_aplicada.png", "APLICACIÓN", "R2 aplicado con hash",
     "Tras confirmar, el estado es done y se muestra «P-01 aplicado (exams_destroy ensayado sobre d68ebfa913ec)». "
     "El hash es el real de la simulación.", 14),
    (9, "09_persistencia_recarga.png", "APLICACIÓN", "Persistencia tras recargar",
     "Después de recargar, las decisiones se recuperan de localStorage: «Pendientes: 4» y «Decididos: 2».", 15),
    (10, "10_vista_movil.png", "APLICACIÓN", "Vista móvil",
     "A 390 × 844 la bandeja, el buscador y las tarjetas se adaptan a la pantalla sin desbordarse.", 16),
    (11, "11_terminal_test_antes.png", "TERMINAL", "Pruebas team: antes",
     "npm test antes de implementar: 17 pruebas pasan y 40 fallan; código de salida 1. Es el punto de partida real.", 2),
    (12, "12_terminal_test_ai_antes.png", "TERMINAL", "Pruebas AI: antes",
     "npm run test:ai antes de corregir: 37 pasan y 20 fallan. Son los seis fallos de los hooks de la IA.", 3),
    (13, "13_terminal_test_despues.png", "TERMINAL", "Pruebas team: después",
     "npm test al terminar: 5 archivos y 57/57 pruebas en verde; código de salida 0.", 4),
    (14, "14_terminal_test_ai_despues.png", "TERMINAL", "Pruebas AI: después",
     "npm run test:ai al terminar: las mismas 57 pruebas en verde sobre src/ai.", 5),
    (15, "15_terminal_build.png", "TERMINAL", "Build de producción",
     "npm run build: compilación de producción correcta, 44 módulos transformados.", 6),
    (16, "16_vscode_estructura.png", "VS CODE", "Estructura del proyecto",
     "Árbol real del repositorio en VS Code y src/hooks/usePlans.js abierto. La estructura coincide con la del "
     "repositorio publicado.", 1),
]

doc = Document()

section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(1.6)
section.bottom_margin = Cm(1.5)
section.left_margin = Cm(1.9)
section.right_margin = Cm(1.9)
section.header_distance = Cm(0.8)
section.footer_distance = Cm(0.7)
section.different_first_page_header_footer = True

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = DARK
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.1


def para(text="", size=10.5, bold=False, color=DARK, align=None, before=0, after=4,
         italic=False, font=None):
    p = doc.add_paragraph()
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = color
    if font:
        run.font.name = font
    return p


def rich(p, text, size=9, color=GRAY):
    """Texto con `code` en Consolas y **negrita**."""
    parts = text.split("`")
    for i, part in enumerate(parts):
        if not part:
            continue
        run = p.add_run(part)
        run.font.size = Pt(size)
        run.font.color.rgb = color
        if i % 2 == 1:
            run.font.name = "Consolas"
            run.font.size = Pt(size - 0.4)


def page_break():
    doc.add_page_break()


def shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def cell_text(cell, text, size=7.8, bold=False, color=DARK, font=None, align=None):
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.color.rgb = color
    if font:
        run.font.name = font


def set_table_borders(table, color="D5DCE8", size=4):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def footer_setup():
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Edu Edward Galindo Huamán · DAE Semana 11 (extra) · pág. ")
    run.font.size = Pt(7.5)
    run.font.color.rgb = GRAY
    run2 = p.add_run()
    run2.font.size = Pt(7.5)
    run2.font.color.rgb = GRAY
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run2._r.append(fld1)
    run2._r.append(instr)
    run2._r.append(fld2)

    header = section.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hr = hp.add_run("Evidencias de capturas reales · Los hooks del agente")
    hr.font.size = Pt(7.5)
    hr.font.color.rgb = GRAY


footer_setup()

# --------------------------------------------------------------------------- #
# PORTADA
# --------------------------------------------------------------------------- #
for _ in range(2):
    para("", after=0)

para("TECSUP", size=12, bold=True, color=NAVY, align="center", after=2)
para("Desarrollo de Aplicaciones Empresariales", size=11.5, color=GRAY, align="center", after=26)

para("Evidencias de capturas reales", size=31, bold=True, color=NAVY, align="center", after=4)
para("Laboratorio de la semana 11 (extra) · Los hooks del agente", size=14.5, color=BLUE,
     align="center", after=8)
para("16 capturas de la aplicación, la terminal y VS Code", size=12, color=DARK,
     align="center", after=26)

rule = para("", after=24)
pPr = rule._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "12")
bottom.set(qn("w:color"), "1F3864")
pBdr.append(bottom)
pPr.append(pBdr)

rows = [
    ("Estudiante", "Edu Edward Galindo Huamán"),
    ("Curso", "Desarrollo de Aplicaciones Empresariales (DAE)"),
    ("Fecha", "9 de noviembre de 2026"),
    ("Repositorio", "https://github.com/edugalindo-jpg/dae-s11-hooks-del-agente"),
    ("Contenido", "10 capturas de la app · 5 de terminal · 1 de VS Code"),
]
for label, value in rows:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(5)
    r1 = p.add_run(label + ": ")
    r1.bold = True
    r1.font.size = Pt(10)
    r1.font.color.rgb = NAVY
    r2 = p.add_run(value)
    r2.font.size = Pt(10)
    r2.font.color.rgb = DARK

para("", after=22)
para("Capturas verificadas · sin montajes", size=11, bold=True, color=GREEN,
     align="center", after=4)
para("Cada imagen corresponde a un estado provocado y comprobado en la aplicación real "
     "o a una corrida real de la terminal.", size=9.5, color=GRAY, align="center", after=0)

page_break()

# --------------------------------------------------------------------------- #
# ÍNDICE
# --------------------------------------------------------------------------- #
para("Índice de capturas", size=20, bold=True, color=NAVY, after=6)
p = doc.add_paragraph()
p.paragraph_format.space_after = Pt(8)
rich(p, "Las 16 evidencias se presentan en el mismo orden en que se generaron. "
        "Cada una ocupa una página e indica qué demuestra y su figura en el informe.", size=9.5)

table = doc.add_table(rows=1, cols=4)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = False
set_table_borders(table)
widths = [1.1, 5.9, 9.1, 1.0]
headers = ["N.º", "Archivo", "Contenido", "Pág."]
for i, text in enumerate(headers):
    cell = table.rows[0].cells[i]
    cell.width = Cm(widths[i])
    shade(cell, "1F3864")
    cell_text(cell, text, size=7.8, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF),
              align="center" if i in (0, 3) else None)

for idx, (n, fname, sec, title, desc, fig) in enumerate(EVID):
    row = table.add_row()
    for i in range(4):
        row.cells[i].width = Cm(widths[i])
        if idx % 2 == 1:
            shade(row.cells[i], "F4F7FC")
    cell_text(row.cells[0], f"{n:02d}", size=7.8, bold=True, color=NAVY, align="center")
    cell_text(row.cells[1], fname, size=7.1, font="Consolas", color=DARK)
    cell_text(row.cells[2], f"{sec.title()} · {title}", size=7.8, color=DARK)
    cell_text(row.cells[3], str(n + 2), size=7.8, color=GRAY, align="center")

for row in table.rows[1:]:
    for cell in row.cells:
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(1.8)
            p.paragraph_format.space_after = Pt(1.8)

para("", after=8)

callout = doc.add_table(rows=1, cols=1)
callout.alignment = WD_TABLE_ALIGNMENT.CENTER
callout.autofit = False
callout.rows[0].cells[0].width = Cm(17.2)
shade(callout.rows[0].cells[0], "EEF4FF")
set_table_borders(callout, color="93B4F5", size=6)
cp = callout.rows[0].cells[0].paragraphs[0]
cp.paragraph_format.space_before = Pt(5)
cp.paragraph_format.space_after = Pt(5)
cp.paragraph_format.line_spacing = 1.08
r = cp.add_run("Cómo se obtuvieron. ")
r.bold = True
r.font.size = Pt(8.6)
r.font.color.rgb = NAVY
rich(cp,
     "Las capturas de la aplicación se tomaron con Playwright contra el servidor real de Vite en "
     "localhost:5173, con viewport de 900 × 900 y deviceScaleFactor 2, y se recortaron a la columna de 880 px. "
     "Antes de cada disparo, el script comprobó que el texto esperado estuviera visible. Las capturas de terminal "
     "se tomaron con PrintWindow sobre la ventana de PowerShell; la de VS Code, de la ventana del editor. "
     "Los originales y los 8 registros completos están en el repositorio (docs/capturas y docs/logs).",
     size=8.6, color=GRAY)

para("Cobertura de la entrega", size=11, bold=True, color=NAVY, before=14, after=3)
for label, text in (
    ("Tres hooks propios:",
     "useDebounce, usePlans y useAgentRun, con sus estados de carga, error y confirmación."),
    ("Seis fallos de la IA:",
     "corregidos y demostrados con el antes y el después de las pruebas."),
    ("Propiedad del estado:",
     "el contador de pendientes se calcula en App; Header solo recibe el número."),
    ("Pruebas y build:",
     "57/57 en modo team, 57/57 en modo AI y compilación de producción correcta."),
    ("Reproducibilidad:",
     "scripts de captura, 8 registros completos y hashes en el repositorio."),
):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.0
    r = p.add_run(label + " ")
    r.bold = True
    r.font.size = Pt(8.8)
    r.font.color.rgb = NAVY
    rich(p, text, size=8.8, color=GRAY)

page_break()

# --------------------------------------------------------------------------- #
# CAPTURAS
# --------------------------------------------------------------------------- #
MAX_W = Cm(17.2)
MAX_H = Cm(19.4)

for idx, (n, fname, sec, title, desc, fig) in enumerate(EVID):
    path = CAP / fname
    if not path.exists():
        raise FileNotFoundError(path)
    w, h = Image.open(path).size

    para(f"EVIDENCIA {n:02d} DE 16 · {sec}", size=8, bold=True, color=BLUE, after=1)
    para(title, size=15.5, bold=True, color=NAVY, after=2)

    p_img = doc.add_paragraph()
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(4)
    p_img.paragraph_format.space_after = Pt(5)
    p_img.paragraph_format.keep_with_next = True
    run = p_img.add_run()
    if w / h >= 17.2 / 19.4:
        run.add_picture(str(path), width=MAX_W)
    else:
        run.add_picture(str(path), height=MAX_H)

    p_cap = doc.add_paragraph()
    p_cap.paragraph_format.space_after = Pt(3)
    p_cap.paragraph_format.line_spacing = 1.05
    p_cap.paragraph_format.keep_together = True
    lead = p_cap.add_run("Qué demuestra. ")
    lead.bold = True
    lead.font.size = Pt(9)
    lead.font.color.rgb = NAVY
    rich(p_cap, desc, size=9, color=GRAY)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_before = Pt(2)
    p_meta.paragraph_format.space_after = Pt(0)
    mr = p_meta.add_run(
        f"Archivo: capturas/{fname}  ·  {w} × {h} px  ·  Figura {fig} del informe"
    )
    mr.font.name = "Consolas"
    mr.font.size = Pt(7.2)
    mr.font.color.rgb = RGBColor(0x6B, 0x77, 0x88)

    if idx < len(EVID) - 1:
        page_break()

props = doc.core_properties
props.title = "Evidencias de capturas reales — DAE S11 (extra)"
props.author = "Edu Edward Galindo Huamán"
props.subject = "Laboratorio S11: Los hooks del agente"
props.keywords = "DAE, React, hooks, evidencias, capturas, Edu Galindo"
props.comments = "16 capturas reales: 10 de la aplicación, 5 de terminal y 1 de VS Code."

doc.save(OUT)
print("DOCX generado:", OUT)
print("Capturas:", len(EVID))

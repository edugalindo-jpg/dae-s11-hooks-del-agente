# -*- coding: utf-8 -*-
"""Genera el entregable DOCX del laboratorio S11 (extra): Los hooks del agente.

Uso:
    python docs/scripts/generar_entregable.py
"""

import os
import re
from pathlib import Path

from PIL import Image
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[2]
DOCS = ROOT / "docs"
CAP = DOCS / "capturas"
LOGS = DOCS / "logs"
OUT = DOCS / "Entregable_S11_Hooks_del_Agente_Edu_Galindo.docx"

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x54, 0x96)
GRAY = RGBColor(0x47, 0x55, 0x69)
DARK = RGBColor(0x0F, 0x17, 0x2A)
GREEN = RGBColor(0x04, 0x78, 0x57)
RED = RGBColor(0xB9, 0x1C, 0x1C)

doc = Document()

# --------------------------------------------------------------------------- #
# PÃ¡gina y estilos
# --------------------------------------------------------------------------- #
section = doc.sections[0]
section.page_width = Cm(21)
section.page_height = Cm(29.7)
section.top_margin = Cm(1.9)
section.bottom_margin = Cm(1.8)
section.left_margin = Cm(2.2)
section.right_margin = Cm(2.2)
section.header_distance = Cm(1.0)
section.footer_distance = Cm(0.9)
section.different_first_page_header_footer = True

styles = doc.styles
normal = styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(10.5)
normal.font.color.rgb = DARK
normal.paragraph_format.space_after = Pt(4)
normal.paragraph_format.line_spacing = 1.12

for name, size, color, before, after in (
    ("Heading 1", 17, NAVY, 14, 6),
    ("Heading 2", 13.5, BLUE, 12, 4),
    ("Heading 3", 11.5, NAVY, 10, 3),
):
    st = styles[name]
    st.font.name = "Calibri"
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = color
    st.paragraph_format.space_before = Pt(before)
    st.paragraph_format.space_after = Pt(after)
    st.paragraph_format.keep_with_next = True

code_style = styles.add_style("CodeBlock", WD_STYLE_TYPE.PARAGRAPH)
code_style.base_style = styles["Normal"]
code_style.font.name = "Consolas"
code_style.font.size = Pt(8)
code_style.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)
code_style.paragraph_format.space_before = Pt(2)
code_style.paragraph_format.space_after = Pt(2)
code_style.paragraph_format.line_spacing = 1.0
code_style.paragraph_format.left_indent = Cm(0.25)
code_style.paragraph_format.right_indent = Cm(0.1)

caption_style = styles.add_style("FigCaption", WD_STYLE_TYPE.PARAGRAPH)
caption_style.base_style = styles["Normal"]
caption_style.font.name = "Calibri"
caption_style.font.size = Pt(8.5)
caption_style.font.italic = True
caption_style.font.color.rgb = GRAY
caption_style.paragraph_format.space_before = Pt(3)
caption_style.paragraph_format.space_after = Pt(10)
caption_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
TOKEN = re.compile(r"(\*\*[^*]+\*\*|`[^`]+`)")


def add_runs(paragraph, text, size=None, color=None, force_bold=False):
    """Escribe texto con **negrita** y `cÃ³digo` en lÃ­nea (tambiÃ©n dentro de la negrita)."""
    for part in TOKEN.split(text):
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            # Recurse so `code` inside **bold** still becomes an inline-code run.
            add_runs(paragraph, part[2:-2], size=size, color=color, force_bold=True)
            continue
        if part.startswith("`") and part.endswith("`"):
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt((size or 10.5) - 1)
            run.font.color.rgb = RGBColor(0x9A, 0x34, 0x12)
        else:
            run = paragraph.add_run(part)
        if force_bold:
            run.bold = True
        if size:
            run.font.size = Pt(size)
        if color is not None:
            run.font.color.rgb = color


def para(text="", size=None, bold=False, italic=False, color=None, align=None,
         space_before=None, space_after=None, style=None):
    p = doc.add_paragraph(style=style)
    add_runs(p, text, size=size, color=color)
    if bold:
        for r in p.runs:
            r.bold = True
    if italic:
        for r in p.runs:
            r.italic = True
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if space_before is not None:
        p.paragraph_format.space_before = Pt(space_before)
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def bullet(text, size=None):
    p = doc.add_paragraph(style="List Bullet")
    add_runs(p, text, size=size)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.08
    return p


def h1(text):
    return doc.add_heading(text, level=1)


def h2(text):
    return doc.add_heading(text, level=2)


def h3(text):
    return doc.add_heading(text, level=3)


def page_break():
    doc.add_page_break()


def shade(element_pr, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    element_pr.append(shd)


def code(text, size=None):
    p = doc.add_paragraph(style="CodeBlock")
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "4")
        el.set(qn("w:color"), "CBD5E1")
        pBdr.append(el)
    pPr.append(pBdr)
    shade(pPr, "F5F7FA")
    lines = text.strip("\n").split("\n")
    for index, line in enumerate(lines):
        run = p.add_run(line if line else " ")
        run.font.name = "Consolas"
        run.font.size = Pt(size or 8)
        if index < len(lines) - 1:
            run.add_break()
    # A block that is taller than a page would be split by Word anyway; forcing it
    # together only leaves a half-empty page behind. Keep short blocks intact.
    p.paragraph_format.keep_together = len(lines) <= 45
    return p


def callout(label, text, fill="EEF4FF", border="93B4F5"):
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.cell(0, 0)
    shade(cell._tc.get_or_add_tcPr(), fill)
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(label + " ")
    run.bold = True
    run.font.size = Pt(9.5)
    run.font.color.rgb = NAVY
    add_runs(p, text, size=9.5, color=DARK)
    p.paragraph_format.space_after = Pt(0)
    # borde
    tblPr = t._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + edge)
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:color"), border)
        borders.append(el)
    tblPr.append(borders)
    para("", space_after=4)


def add_table(headers, rows, widths_cm, font_size=8.5, header_fill="1F3864",
              align_center_cols=(), zebra=True):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    header = table.rows[0]
    trPr = header._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
    trPr.append(OxmlElement("w:cantSplit"))
    for i, text in enumerate(headers):
        cell = header.cells[i]
        cell.width = Cm(widths_cm[i])
        shade(cell._tc.get_or_add_tcPr(), header_fill)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(font_size)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_index, row in enumerate(rows):
        cells = table.add_row().cells
        trPr = table.rows[-1]._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        for i, value in enumerate(row):
            cells[i].width = Cm(widths_cm[i])
            if zebra and r_index % 2 == 1:
                shade(cells[i]._tc.get_or_add_tcPr(), "F2F6FC")
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.0
            if i in align_center_cols:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_runs(p, str(value), size=font_size)
    para("", space_after=4)
    return table


FIG = {"n": 0}


def figure(path, caption, max_w=6.35, max_h=8.2):
    assert path.exists(), f"Falta la imagen: {path}"
    w, h = Image.open(path).size
    width = max_w
    if width * (h / w) > max_h:
        width = max_h * (w / h)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(str(path), width=Inches(width))
    FIG["n"] += 1
    cap = doc.add_paragraph(style="FigCaption")
    run = cap.add_run(f"Figura {FIG['n']}. ")
    run.bold = True
    run.font.size = Pt(8.5)
    run.font.color.rgb = NAVY
    # Same inline markup as the body text (**bold**, `code`), so no backticks leak into the PDF.
    for part in TOKEN.split(caption):
        if not part:
            continue
        if part.startswith("`") and part.endswith("`"):
            cr = cap.add_run(part[1:-1])
            cr.font.name = "Consolas"
            cr.font.size = Pt(7.8)
            cr.font.color.rgb = RGBColor(0x9A, 0x34, 0x12)
        elif part.startswith("**") and part.endswith("**"):
            cr = cap.add_run(part[2:-2])
            cr.bold = True
        else:
            cap.add_run(part)
    cap.paragraph_format.keep_together = True


def source_block(path, label):
    """Incluye un archivo de cÃ³digo con un rÃ³tulo."""
    p = doc.add_paragraph()
    run = p.add_run(label)
    run.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = BLUE
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.keep_with_next = True
    code(path.read_text(encoding="utf-8"))


def footer_setup():
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Edu Edward Galindo HuamÃ¡n Â· DAE Semana 11 Â· Los hooks del agente Â· pÃ¡g. ")
    run.font.size = Pt(8)
    run.font.color.rgb = GRAY
    run2 = p.add_run()
    run2.font.size = Pt(8)
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
    hr = hp.add_run("Laboratorio S11 (extra) Â· Los hooks del agente")
    hr.font.size = Pt(8)
    hr.font.color.rgb = GRAY


footer_setup()

# --------------------------------------------------------------------------- #
# PORTADA
# --------------------------------------------------------------------------- #
for _ in range(2):
    para("", space_after=0)

p = para("TECSUP", size=12, bold=True, color=NAVY, align="center", space_after=2)
p = para("Desarrollo de Aplicaciones Empresariales", size=11.5, color=GRAY, align="center", space_after=24)

p = para("Los hooks del agente", size=34, bold=True, color=NAVY, align="center", space_after=4)
p = para("Laboratorio de la semana 11 (extra)", size=15, color=BLUE, align="center", space_after=10)
p = para("ImplementaciÃ³n de useDebounce, usePlans y useAgentRun", size=12.5, color=DARK, align="center", space_after=2)
p = para("y revisiÃ³n de seis fallos en hooks escritos por una IA", size=12.5, color=DARK, align="center", space_after=30)

# LÃ­nea decorativa
p = para("", space_after=26)
pPr = p._p.get_or_add_pPr()
pBdr = OxmlElement("w:pBdr")
bottom = OxmlElement("w:bottom")
bottom.set(qn("w:val"), "single")
bottom.set(qn("w:sz"), "12")
bottom.set(qn("w:color"), "1F3864")
pBdr.append(bottom)
pPr.append(pBdr)

cover_rows = [
    ("Estudiante", "Edu Edward Galindo HuamÃ¡n"),
    ("Curso", "Desarrollo de Aplicaciones Empresariales (DAE)"),
    ("Fecha", "9 de noviembre de 2026"),
    ("Repositorio", "https://github.com/Hellscythe25/dae-s11-hooks-del-agente"),
    ("Entorno", "Node.js v24.21.0 Â· npm 11.19.0 Â· Vite 7.3.7 Â· Vitest 3.2.7"),
]
for label, value in cover_rows:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = p.add_run(label + ": ")
    r1.bold = True
    r1.font.size = Pt(10.5)
    r1.font.color.rgb = NAVY
    r2 = p.add_run(value)
    r2.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(4)

para("", space_after=18)
p = para("Resultado obtenido", size=11, bold=True, color=GREEN, align="center", space_after=6)
p = para("57 / 57 pruebas en modo team  Â·  57 / 57 pruebas en modo ai  Â·  build correcto",
         size=12, bold=True, color=DARK, align="center", space_after=4)
p = para("16 capturas reales: 10 de la aplicaciÃ³n, 5 de terminal y 1 de VS Code",
         size=10, color=GRAY, align="center", space_after=0)

page_break()

# --------------------------------------------------------------------------- #
# ÃNDICE (campo TOC; Word lo actualiza al exportar)
# --------------------------------------------------------------------------- #
h1("Contenido")
p = doc.add_paragraph()
run = p.add_run()
fld_begin = OxmlElement("w:fldChar")
fld_begin.set(qn("w:fldCharType"), "begin")
instr = OxmlElement("w:instrText")
instr.set(qn("xml:space"), "preserve")
instr.text = r' TOC \o "1-2" \h \z \u '
fld_sep = OxmlElement("w:fldChar")
fld_sep.set(qn("w:fldCharType"), "separate")
run._r.append(fld_begin)
run._r.append(instr)
run._r.append(fld_sep)

placeholder_run = p.add_run("Ãndice en preparaciÃ³n.")
placeholder_run.font.size = Pt(9)
placeholder_run.font.color.rgb = GRAY

end_run = p.add_run()
fld_end = OxmlElement("w:fldChar")
fld_end.set(qn("w:fldCharType"), "end")
end_run._r.append(fld_end)

# --------------------------------------------------------------------------- #
# 1. RESUMEN
# --------------------------------------------------------------------------- #
page_break()
h1("1. Resumen ejecutivo")

para(
    "Este trabajo completa el laboratorio extra de la semana 11 de Desarrollo de Aplicaciones Empresariales. "
    "La aplicaciÃ³n es una bandeja de aprobaciones: un agente de IA propone planes con un `plan_hash` y una persona "
    "decide si los aprueba o los rechaza. La bandeja carga los planes con una peticiÃ³n simulada, busca con retardo y "
    "permite ejecutar lo aprobado; un plan destructivo (R2) no se aplica sin una confirmaciÃ³n explÃ­cita.",
    align="justify",
)
para(
    "El repositorio de partida traÃ­a los componentes y las pruebas ya escritos, pero los tres hooks centrales estaban "
    "sin implementar: por eso la aplicaciÃ³n se quedaba cargando y 40 de las 57 pruebas fallaban. El trabajo tuvo dos "
    "frentes: (1) implementar `useDebounce`, `usePlans` y `useAgentRun` contra las pruebas, y (2) revisar tres hooks "
    "Â«escritos por una IAÂ» que compilaban y parecÃ­an correctos, pero contenÃ­an seis fallos tÃ­picos de React.",
    align="justify",
)
para(
    "No se modificaron las pruebas, `package.json`, `vite.config.js` ni los archivos `.env.*`. Todo el cambio de "
    "cÃ³digo estÃ¡ en `src/hooks/` y `src/ai/`. Las capturas se tomaron de la aplicaciÃ³n real corriendo en el servidor "
    "de desarrollo y de la terminal real; no hay maquetas ni imÃ¡genes simuladas.",
    align="justify",
)

h2("1.1 Resultados")
add_table(
    ["Indicador", "Estado inicial", "Estado final"],
    [
        ["`npm test` (modo team, `src/hooks`)", "40 fallan Â· 17 pasan (57)", "**57 pasan (57)**"],
        ["`npm run test:ai` (modo ai, `src/ai`)", "20 fallan Â· 37 pasan (57)", "**57 pasan (57)**"],
        ["`npm run build`", "no se habÃ­a ejecutado", "**correcto (44 mÃ³dulos)**"],
        ["Capturas reales", "â€”", "**16** (10 app + 5 terminal + 1 VS Code)"],
    ],
    widths_cm=[6.4, 5.2, 5.0],
    font_size=9,
)

h2("1.2 QuÃ© contiene este entregable")
for item in (
    "La explicaciÃ³n de la arquitectura y de **por quÃ© el contador de pendientes vive en `App` y no en `Header`** "
    "(secciÃ³n 2.4), con el cÃ³digo que lo demuestra y las pruebas que lo respaldan.",
    "El cÃ³digo final comentado de los tres hooks propios y las decisiones de diseÃ±o de cada uno (secciÃ³n 3).",
    "La revisiÃ³n de los hooks de la IA: los **seis fallos** clasificados, el Â«antes y despuÃ©sÂ» de cada correcciÃ³n y "
    "la evidencia de las pruebas (secciÃ³n 4).",
    "Las tablas de pruebas, los comandos y las **capturas reales** de terminal, VS Code y aplicaciÃ³n (secciones 5 y 6).",
    "La **tabla de uso de IA**, lo que se aceptÃ³, lo que se rechazÃ³ y las cuatro preguntas de revisiÃ³n del diff "
    "(secciÃ³n 7).",
    "Los anexos con el cÃ³digo corregido de `src/ai/` y el catÃ¡logo completo de las 57 pruebas.",
):
    bullet(item)

# --------------------------------------------------------------------------- #
# 2. APLICACIÃ“N Y PROPIEDAD DEL ESTADO
# --------------------------------------------------------------------------- #
h1("2. La aplicaciÃ³n y la propiedad del estado")

h2("2.1 QuÃ© hace la aplicaciÃ³n")
para(
    "`App` pide los planes al agente con `usePlans` (red simulada con 900 ms de retardo y un modo `?fail=1` que "
    "rechaza). Mientras llegan, muestra un esqueleto de carga; si la peticiÃ³n falla, muestra un aviso con "
    "Â«ReintentarÂ»; cuando llegan, dibuja la bandeja. Cada plan se puede aprobar o rechazar. Aprobar un plan R2 exige "
    "escribir su hash exacto. Las decisiones se guardan en `localStorage` y sobreviven a una recarga. Los planes "
    "aprobados aparecen en el panel de ejecuciÃ³n, que usa `useAgentRun` para ensayar primero y aplicar despuÃ©s.",
    align="justify",
)

h2("2.2 Ãrbol de componentes")
code(
    """App                                  â† usePlans + estado decisions + persistencia
â”œâ”€â”€ Header({ pendingCount })           â† solo dibuja el nÃºmero
â”œâ”€â”€ toolbar                            â† simular fallo / recargar
â””â”€â”€ Inbox({ plans, decisions, onDecide, runAgent })
    â”œâ”€â”€ searchbox                      â† useState(query) + useDebounce(query, 300)
    â”œâ”€â”€ PlanCard({ plan, onDecide })   â† aprobar / rechazar
    â”‚   â””â”€â”€ HashConfirm                â† input del hash para R2
    â””â”€â”€ RunPanel({ approvedPlans, runAgent })
        â””â”€â”€ useAgentRun                â† idle Â· running Â· awaiting_approval Â· done Â· error""",
    size=7.8,
)

figure(
    CAP / "16_vscode_estructura.png",
    "Estructura real del proyecto en VS Code: el Ã¡rbol de archivos del repositorio y `src/hooks/usePlans.js` abierto.",
)

h2("2.3 DÃ³nde vive cada dato")
add_table(
    ["Dato", "DÃ³nde vive", "QuiÃ©n lo escribe", "QuiÃ©n lo lee"],
    [
        ["`status`, `plans`, `error`", "`usePlans`, invocado en `App`", "La promesa de `fetchPlans`", "`App` (decide quÃ© dibujar) e `Inbox`"],
        ["`decisions`", "`App`: `useState` + `localStorage`", "`decide()` cuando `Inbox` avisa", "`Header` (conteo) e `Inbox` (listas)"],
        ["`pendingCount`", "**Derivado en `App`**, no es estado", "Se recalcula en cada render", "`Header`, como nÃºmero"],
        ["`query`", "`Inbox` (`useState`)", "El campo de bÃºsqueda", "`Inbox`, a travÃ©s de `useDebounce`"],
        ["`run`", "`useAgentRun`, invocado en `RunPanel`", "`start`, `approve`, `reject`", "`RunPanel` (estado de la ejecuciÃ³n)"],
    ],
    widths_cm=[3.2, 4.4, 4.4, 4.6],
    font_size=8.3,
)

h2("2.4 Por quÃ© el contador de pendientes vive en App y no en Header")
para(
    "El contador de pendientes **no es un dato independiente**: es el resultado de cruzar la lista de planes con el "
    "mapa de decisiones. Por eso no se guarda: se deriva. `App` ya tiene las dos entradas (los planes que devuelve "
    "`usePlans` y las decisiones que administra con `useState`), asÃ­ que allÃ­ se calcula una sola vez y se pasa a "
    "`Header` Ãºnicamente el nÃºmero.",
    align="justify",
)
code(
    """// App.jsx â€” las decisiones viven en el ancestro comÃºn mÃ¡s bajo
const { status, plans, error, reload } = usePlans(fetchPlans);
const [decisions, setDecisions] = useState(readStoredDecisions);

// Derivado en cada render, nunca almacenado. La cabecera recibe solo este nÃºmero.
const pendingCount =
  status === "ready" ? plans.filter((plan) => !(plan.id in decisions)).length : null;

return (
  <main className="app">
    <Header pendingCount={pendingCount} />
    ...
    {status === "ready" && (
      <Inbox plans={plans} decisions={decisions} onDecide={decide} runAgent={runAgent} />
    )}
  </main>
);""",
    size=8.2,
)

para("Las razones concretas, en orden de importancia:", align="justify")
bullet(
    "**El ancestro comÃºn mÃ¡s bajo es `App`.** `Header` e `Inbox` son hermanos. Las decisiones se producen al decidir "
    "dentro de `Inbox` (en `PlanCard`) y las necesitan tanto `Inbox` (para separar pendientes de decididos) como "
    "`Header` (para contar). El Ãºnico componente que es ancestro de ambos es `App`; si el contador viviera en "
    "`Header`, las decisiones tendrÃ­an que subir de todas formas y se duplicarÃ­a el camino."
)
bullet(
    "**Una sola fuente de verdad.** `decisions` se guarda una vez y se persiste en `localStorage`. El contador es una "
    "vista de ese estado. Si `Header` tuviera su propio contador, habrÃ­a dos copias que podrÃ­an desincronizarse al "
    "recargar, al rechazar o al cambiar la lista de planes."
)
bullet(
    "**`Header` es presentacional.** Recibe solo `pendingCount` y no la lista. No se vuelve a dibujar por cambios en "
    "los planes que no le interesan y se puede probar aislado. La prueba Â«la cabecera solo necesita un nÃºmero: se "
    "dibuja sin la lista de planesÂ» lo verifica explÃ­citamente."
)
bullet(
    "**La persistencia encaja sin esfuerzo.** Al recargar, `App` lee las decisiones de `localStorage` y el contador se "
    "recalcula solo. Un contador dentro de `Header` no sabrÃ­a cuÃ¡ndo cambiÃ³ la lista ni sobrevivirÃ­a a la recarga."
)
bullet(
    "**Menos complejidad, no mÃ¡s.** La alternativa serÃ­a un contexto global o un estado en `Header` actualizado por "
    "eventos. Con solo dos consumidores que ya cuelgan de `App`, eso agrega indirecciÃ³n sin beneficio. La regla del "
    "curso es Â«el dato vive en el ancestro mÃ¡s bajo que lo necesitaÂ», y ese ancestro es `App`, no `main.jsx`."
)

callout(
    "Respuesta corta:",
    "el contador vive en `App` porque allÃ­ estÃ¡n los planes y las decisiones, que son sus dos entradas; `Header` solo "
    "lo muestra. Guardarlo en la cabecera duplicarÃ­a el estado y romperÃ­a la Ãºnica fuente de verdad.",
)

# --------------------------------------------------------------------------- #
# 3. HOOKS PROPIOS
# --------------------------------------------------------------------------- #
page_break()
h1("3. ImplementaciÃ³n de los tres hooks propios")

para(
    "Los tres hooks se implementaron contra las pruebas, sin modificarlas. Cada archivo conserva las convenciones del "
    "curso: cÃ³digo y nombres en inglÃ©s, comentarios en inglÃ©s, y explicaciÃ³n en espaÃ±ol en este informe. Los hooks se "
    "llaman siempre en el nivel superior y en el mismo orden.",
    align="justify",
)

# 3.1 useDebounce
h2("3.1 useDebounce: el valor tranquilo")
para(
    "**Contrato.** `useDebounce(value, delay)` devuelve el Ãºltimo `value` que lleva `delay` milisegundos sin cambiar. "
    "En el primer dibujado devuelve `value` de inmediato. Cada cambio cancela el temporizador anterior; si el retardo "
    "cambia, se vuelve a contar. Al desmontar no queda ningÃºn temporizador vivo.",
    align="justify",
)
source_block(ROOT / "src/hooks/useDebounce.js", "src/hooks/useDebounce.js")

para("**Decisiones de diseÃ±o.**", space_after=2)
bullet(
    "`useState(value)` arranca con el valor actual: la interfaz no espera `delay` ms en el primer dibujado."
)
bullet(
    "El temporizador se crea dentro del `useEffect` con dependencias `[value, delay]`. React ejecuta la limpieza "
    "**antes** de repetir el efecto, asÃ­ que `clearTimeout(timer)` cancela el temporizador anterior en cada cambio. "
    "Solo sobrevive el Ãºltimo; los valores intermedios nunca llegan a `setDebounced`."
)
bullet(
    "La misma limpieza corre al desmontar. No hace falta un `useRef`: el temporizador pertenece a una corrida concreta "
    "del efecto y la clausura lo captura. Guardarlo en un `ref` fue precisamente el error del hook de la IA."
)
bullet(
    "Si `delay` cambia, el efecto se repite y la cuenta empieza de cero con el nuevo retardo, tal como exige la prueba."
)
para(
    "**Pruebas (10/10).** El archivo `useDebounce.test.jsx` agrupa: el valor tranquilo (3), la cancelaciÃ³n del "
    "temporizador anterior (4) y la limpieza al desmontar (3). Usa `vi.useFakeTimers()` y `act()` para mover el reloj; "
    "no espera en tiempo real.",
    align="justify",
)

# 3.2 usePlans
h2("3.2 usePlans: los tres estados y la respuesta vieja")
para(
    "**Contrato.** `usePlans(fetchPlans)` devuelve `{ status, plans, error, reload }`. `status` es `loading`, `ready` "
    "o `error`. Mientras carga y si falla, `plans` es `[]`; en `error`, `error` es el motivo del rechazo. `reload()` "
    "vuelve a pedir y regresa a `loading`. La respuesta de una peticiÃ³n vieja nunca debe pisar a la nueva.",
    align="justify",
)
source_block(ROOT / "src/hooks/usePlans.js", "src/hooks/usePlans.js")

para("**Decisiones de diseÃ±o.**", space_after=2)
bullet(
    "El estado se agrupa en un solo objeto (`INITIAL` con `status`, `plans` y `error`) para que los tres estados se "
    "actualicen juntos y no existan combinaciones imposibles, como `status: \"ready\"` con un error viejo."
)
bullet(
    "`reload` incrementa un contador `attempt` en lugar de modificar el estado directamente. AsÃ­ el efecto se repite "
    "sin duplicar la lÃ³gica de carga, y `reload` conserva la misma identidad entre renders gracias a `useCallback`."
)
bullet(
    "Cada corrida del efecto crea una bandera `cancelled`. La limpieza la pone en `true`; antes de escribir el estado "
    "se comprueba. Esto cubre cuatro casos que las pruebas ejercitan: una respuesta vieja que llega despuÃ©s de la "
    "nueva, un error viejo que llega tarde, un cambio de `fetchPlans` y el doble efecto de `StrictMode`."
)
bullet(
    "La bandera tambiÃ©n cubre el desmontaje: si la respuesta llega despuÃ©s, no se llama `setState`. No se usa "
    "`AbortController` aquÃ­ porque la funciÃ³n `fetchPlans` de la aplicaciÃ³n no recibe una seÃ±al; la bandera es "
    "suficiente y no cambia el contrato."
)
para(
    "**Pruebas (12/12).** `usePlans.test.jsx` cubre los tres estados (4), `reload` (3), la carrera de peticiones (4) y "
    "el desmontaje (1). Las promesas son Â«deferredÂ»: la prueba decide cuÃ¡ndo y en quÃ© orden responde cada peticiÃ³n.",
    align="justify",
)

# 3.3 useAgentRun
h2("3.3 useAgentRun: la mÃ¡quina de estados de una ejecuciÃ³n")
para(
    "**Contrato.** `useAgentRun(runAgent)` devuelve `{ status, plan, result, error, start, approve, reject }`. La "
    "ejecuciÃ³n empieza en `idle`; `start(plan)` pasa a `running`; un plan R1 termina en `done`; un plan R2 termina en "
    "`awaiting_approval` y solo `approve()` lo lleva a `done`; un fallo lleva a `error`; `reject()` cancela y vuelve a "
    "`idle`. Si ya hay una ejecuciÃ³n en curso, `start()` no hace nada.",
    align="justify",
)
code(
    """idle â”€â”€start(plan)â”€â”€â–¶ running â”€â”€ok, R1â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â–¶ done
                          â”‚
                          â”œâ”€â”€ok, R2â”€â”€â–¶ awaiting_approval â”€â”€approve()â”€â”€â–¶ done
                          â”‚
                          â””â”€â”€fallaâ”€â”€â”€â”€â–¶ error

reject() en running o en awaiting_approval: aborta y vuelve a idle.
start() se ignora en running y en awaiting_approval.""",
    size=8.2,
)
source_block(ROOT / "src/hooks/useAgentRun.js", "src/hooks/useAgentRun.js")

para("**Decisiones de diseÃ±o.**", space_after=2)
bullet(
    "`latest` es un `ref` que refleja el Ãºltimo estado. Los manejadores `start`, `approve` y `reject` lo leen para "
    "decidir sin depender de una clausura vieja, incluso si se llaman varias veces antes de que React vuelva a dibujar."
)
bullet(
    "`controller` guarda el `AbortController` de la ejecuciÃ³n actual. `start` crea uno nuevo y pasa su seÃ±al a "
    "`runAgent`; `reject` y el desmontaje llaman a `abort()`. La promesa comprueba `signal.aborted` antes de tocar el "
    "estado, asÃ­ que una respuesta tardÃ­a no puede revivir una ejecuciÃ³n cancelada."
)
bullet(
    "**Regla de seguridad:** el Ã©xito de un plan R2 no va a `done`; va a `awaiting_approval` conservando el resultado. "
    "Solo `approve()` cambia el estado a `done`, y solo funciona si el estado actual es `awaiting_approval`."
)
bullet(
    "Todos los hooks se ejecutan siempre, antes de cualquier `return`. La comprobaciÃ³n de `runAgent` ausente estÃ¡ "
    "dentro de `start`, no como un retorno temprano del hook. Esto mantiene el orden de hooks entre dibujados."
)
para(
    "**Pruebas (18/18).** `useAgentRun.test.jsx` cubre el punto de partida (2), el plan R1 (2), la barrera del R2 (5), "
    "cancelar/fallar/reintentar (5), la limpieza al desmontar (2) y las reglas de los hooks (2).",
    align="justify",
)

# --------------------------------------------------------------------------- #
# 4. HOOKS DE LA IA
# --------------------------------------------------------------------------- #
page_break()
h1("4. RevisiÃ³n de los hooks Â«escritos por una IAÂ»")

para(
    "La carpeta `src/ai/` contiene los mismos tres hooks, escritos por una IA. Compilan, se leen bien y la aplicaciÃ³n "
    "funciona en parte, pero esconden seis fallos tÃ­picos. El modo `npm run test:ai` corre exactamente las mismas 57 "
    "pruebas contra esa carpeta: al empezar, **20 fallan y 37 pasan**.",
    align="justify",
)

h2("4.1 MÃ©todo de revisiÃ³n")
bullet("Se ejecutÃ³ `npm run test:ai` y se guardÃ³ la salida completa en `docs/logs/02_test_ai_inicial.txt`.")
bullet("Se leyÃ³ cada prueba fallida y se buscÃ³ la causa en el hook, no en la prueba. Las pruebas no se tocaron.")
bullet("Se corrigiÃ³ un archivo por turno y se volviÃ³ a correr el modo `ai` despuÃ©s de cada cambio.")
bullet("Se clasificaron los fallos en seis causas raÃ­z, que se documentan en la tabla siguiente.")
bullet("Al final, `npm test` y `npm run test:ai` quedaron los dos en **57/57**.")

h2("4.2 Los seis fallos")
add_table(
    ["#", "Archivo", "Fallo y sÃ­ntoma", "CorrecciÃ³n aplicada", "Pruebas que lo demuestran"],
    [
        ["1", "`src/ai/useDebounce.js`",
         "El temporizador solo se limpia al desmontar. Un segundo `useEffect` con `[]` guarda el timer en un `ref`; "
         "al cambiar el valor, el timer anterior sigue vivo y emite valores intermedios.",
         "Devolver `clearTimeout(timer)` desde el efecto `[value, delay]`, sin `ref`.",
         "5 pruebas de Â«cada cambio cancela el temporizador anteriorÂ» y Â«limpieza al desmontarÂ»."],
        ["2", "`src/ai/usePlans.js`",
         "El rechazo no cambia el estado: el `.catch` solo hace `console.error`. El status se queda en `loading` para "
         "siempre y `error` sigue en `null`.",
         "En el rechazo: `setPlans([])`, `setError(err)` y `setStatus(\"error\")` (con guard).",
         "Â«si la peticiÃ³n falla pasa a errorâ€¦Â», Â«despuÃ©s de un error, reload()â€¦Â» y la prueba de App con red caÃ­da."],
        ["3", "`src/ai/usePlans.js`",
         "No hay bandera de cancelaciÃ³n: la respuesta de un `reload` anterior (o posterior al desmontaje) puede pisar "
         "la mÃ¡s nueva. Es la carrera que las pruebas provocan a propÃ³sito.",
         "`let cancelled = false` por corrida del efecto; la limpieza lo pone en `true`; se comprueba antes de cada "
         "`setState`.",
         "Las 3 pruebas de Â«ignora la respuesta viejaÂ» que fallaban (respuesta tardÃ­a, cambio de `fetchPlans` y "
         "`StrictMode`)."],
        ["4", "`src/ai/useAgentRun.js`",
         "Hay un `return` temprano antes de los hooks: `if (!runAgent) return { ... }`. Eso cambia el nÃºmero de hooks "
         "entre dibujados; cuando `runAgent` aparece, React lanza Â«Rendered more hooks than during the previous renderÂ».",
         "Quitar el retorno temprano; mantener todos los hooks y comprobar `if (!runAgent) return;` **dentro** de "
         "`start`.",
         "Â«sin runAgent se queda en idleâ€¦Â» y Â«si runAgent aparece en un dibujado posterior, el hook no se rompeÂ»."],
        ["5", "`src/ai/useAgentRun.js`",
         "Un plan R2 se aplica sin aprobaciÃ³n: el Ã©xito siempre hace `update({ status: \"done\" })`. No existe la "
         "barrera de seguridad de la semana 10.",
         "Si `plan.risk === \"R2\"`, pasar a `awaiting_approval` conservando el resultado; solo `approve()` lleva a "
         "`done`.",
         "4 pruebas de la barrera R2, Â«esperando aprobaciÃ³n tampoco se puede lanzar otra ejecuciÃ³nÂ» y las 2 de Inbox "
         "con R2."],
        ["6", "`src/ai/useAgentRun.js`",
         "No hay limpieza al desmontar: la seÃ±al no se aborta y una respuesta tardÃ­a intenta actualizar el estado de "
         "un componente que ya no existe.",
         "`useEffect(() => () => controller.current?.abort(), [])`.",
         "Â«al desmontar con una ejecuciÃ³n en curso aborta su seÃ±alÂ» y Â«si la respuesta llega despuÃ©s de desmontar, no "
         "pasa nadaÂ»."],
    ],
    widths_cm=[0.7, 3.0, 5.0, 4.4, 3.5],
    font_size=7.9,
    align_center_cols=(0,),
)

h2("4.3 Antes y despuÃ©s de cada correcciÃ³n")
para(
    "Los fragmentos siguientes muestran el cambio mÃ­nimo que resuelve cada fallo. El cÃ³digo completo corregido estÃ¡ "
    "en el anexo A.",
    align="justify",
)

h3("Fallo 1 Â· limpieza del temporizador")
code(
    """// ANTES (src/ai/useDebounce.js)
useEffect(() => {
  timer.current = setTimeout(() => setDebounced(value), delay);
}, [value, delay]);
useEffect(() => () => clearTimeout(timer.current), []);   // solo al desmontar

// DESPUÃ‰S
useEffect(() => {
  const timer = setTimeout(() => setDebounced(value), delay);
  return () => clearTimeout(timer);                        // al cambiar y al desmontar
}, [value, delay]);""",
)

h3("Fallos 2 y 3 Â· estado de error y respuesta vieja")
code(
    """// ANTES (src/ai/usePlans.js)
fetchPlans()
  .then((data) => { setPlans(data); setStatus("ready"); })
  .catch((err) => { console.error("Error loading plans", err); });   // nunca hay estado "error"

// DESPUÃ‰S
let cancelled = false;                        // una bandera por corrida del efecto
fetchPlans().then(
  (data) => { if (cancelled) return; setPlans(data); setStatus("ready"); setError(null); },
  (err)  => { if (cancelled) return; setPlans([]); setError(err); setStatus("error"); }
);
return () => { cancelled = true; };           // reload, cambio de fetchPlans y desmontaje""",
)

h3("Fallo 4 Â· orden de los hooks")
code(
    """// ANTES (src/ai/useAgentRun.js)
if (!runAgent) {
  return { ...IDLE, start: () => {}, approve: () => {}, reject: () => {} };  // hooks condicionales
}
const latest = useRef(IDLE);   // este hook se salta en el primer dibujado

// DESPUÃ‰S
const latest = useRef(IDLE);   // todos los hooks siempre, en el mismo orden
const start = useCallback((plan) => {
  if (!runAgent) return;       // la comprobaciÃ³n va dentro del manejador
  ...
}, [runAgent, update]);""",
)

h3("Fallo 5 Â· la barrera del R2")
code(
    """// ANTES
runAgent(plan, { signal }).then((result) => {
  update({ status: "done", plan, result, error: null });     // R2 aplicado sin aprobaciÃ³n
});

// DESPUÃ‰S
runAgent(plan, { signal }).then((result) => {
  if (plan.risk === "R2") {
    update({ status: "awaiting_approval", plan, result, error: null });
  } else {
    update({ status: "done", plan, result, error: null });
  }
});""",
)

h3("Fallo 6 Â· limpieza al desmontar")
code(
    """// ANTES: no habÃ­a ningÃºn efecto de limpieza.

// DESPUÃ‰S
useEffect(() => {
  return () => {
    controller.current?.abort();   // la promesa comprueba signal.aborted antes de setState
    controller.current = null;
  };
}, []);""",
)

para(
    "**Resultado:** `npm run test:ai` pasÃ³ de **20 fallan / 37 pasan** a **57 pasan**. La salida completa estÃ¡ en "
    "`docs/logs/05_test_ai_final.txt` y la captura de terminal en la figura 5.",
    align="justify",
)

# --------------------------------------------------------------------------- #
# 5. PRUEBAS Y EVIDENCIAS
# --------------------------------------------------------------------------- #
page_break()
h1("5. Pruebas y evidencias")

h2("5.1 Cobertura por archivo")
add_table(
    ["Suite", "QuÃ© verifica", "N.Âº", "Inicial", "Final"],
    [
        ["`useDebounce.test.jsx`", "Valor tranquilo, cancelaciÃ³n del timer anterior y limpieza al desmontar", "10", "5 fallan", "10 pasan"],
        ["`usePlans.test.jsx`", "Tres estados, `reload`, carrera de peticiones y desmontaje", "12", "11 fallan", "12 pasan"],
        ["`useAgentRun.test.jsx`", "MÃ¡quina de estados, barrera R2, cancelaciÃ³n, error y limpieza", "18", "14 fallan", "18 pasan"],
        ["`App.test.jsx`", "Carga, error, contador en `App` y persistencia", "7", "5 fallan", "7 pasan"],
        ["`Inbox.test.jsx`", "BÃºsqueda con debounce, decisiones y ejecuciÃ³n", "10", "5 fallan", "10 pasan"],
        ["**Total**", "**Las 57 pruebas del curso**", "**57**", "**40 fallan Â· 17 pasan**", "**57 pasan**"],
    ],
    widths_cm=[3.7, 7.9, 1.2, 2.1, 1.8],
    font_size=8.2,
    align_center_cols=(2, 3, 4),
)

h2("5.2 Comandos ejecutados")
code(
    """npm install                 # instala las dependencias del repositorio
npm test                    # modo team: pruebas contra src/hooks
npm run test:ai             # modo ai: las mismas pruebas contra src/ai
npm run build               # compilaciÃ³n de producciÃ³n (vite build)
npm run dev                 # aplicaciÃ³n real para las capturas (http://localhost:5173)""",
)

h2("5.3 Evidencia de terminal")
para(
    "Las capturas siguientes se tomaron de ventanas de PowerShell reales, con la salida real de cada comando y el "
    "cÃ³digo de salida visible. Para que el texto fuera legible en el documento se usÃ³ una fuente de consola de 28 px; "
    "no se editÃ³ el contenido.",
    align="justify",
)
figure(CAP / "11_terminal_test_antes.png",
       "Estado inicial de `npm test`: 40 pruebas fallan y 17 pasan, con cÃ³digo de salida 1.")
figure(CAP / "12_terminal_test_ai_antes.png",
       "Estado inicial de `npm run test:ai`: 20 fallan y 37 pasan. Son los seis fallos de los hooks de la IA.")
figure(CAP / "13_terminal_test_despues.png",
       "Estado final de `npm test`: los 5 archivos y las 57 pruebas en verde, con cÃ³digo de salida 0.")
figure(CAP / "14_terminal_test_ai_despues.png",
       "Estado final de `npm run test:ai`: las mismas 57 pruebas en verde sobre `src/ai`.")
figure(CAP / "15_terminal_build.png",
       "`npm run build`: compilaciÃ³n de producciÃ³n correcta, 44 mÃ³dulos transformados.")

h2("5.4 Registros completos")
para(
    "AdemÃ¡s de las capturas, el repositorio guarda la salida completa de cada corrida en `docs/logs/`. Estos archivos "
    "permiten auditar cada prueba, no solo el resumen.",
    align="justify",
)
add_table(
    ["Archivo", "Contenido"],
    [
        ["`01_test_inicial.txt`", "`npm test` antes de implementar: 40 fallan / 17 pasan"],
        ["`02_test_ai_inicial.txt`", "`npm run test:ai` antes de corregir: 20 fallan / 37 pasan"],
        ["`03_hooks_dirigidos.txt`", "Las 40 pruebas de los tres hooks, en verde"],
        ["`04_test_final.txt`", "`npm test` final: 57/57"],
        ["`05_test_ai_final.txt`", "`npm run test:ai` final: 57/57"],
        ["`06_build.txt`", "`npm run build` final"],
        ["`11_antes_test_consola.txt`", "Salida de consola usada para la figura 2"],
        ["`12_antes_test_ai_consola.txt`", "Salida de consola usada para la figura 3"],
    ],
    widths_cm=[5.0, 11.3],
    font_size=8.5,
)

h2("5.5 QuÃ© no se modificÃ³")
para(
    "El contrato del repositorio prohÃ­be tocar las pruebas, la configuraciÃ³n y los archivos de entorno. El diff de "
    "cÃ³digo afecta Ãºnicamente a los seis hooks. Las pruebas, `package.json`, `vite.config.js` y `.env.*` siguen "
    "exactamente como estaban en el commit inicial.",
    align="justify",
)
code(
    """$ git diff --stat
 src/ai/useAgentRun.js    |  35 ++++++++-----
 src/ai/useDebounce.js    |  14 +++---
 src/ai/usePlans.js       |  30 ++++++++---
 src/hooks/useAgentRun.js | 126 ++++++++++++++++++++++++++++-------------------
 src/hooks/useDebounce.js |  42 ++++++++--------
 src/hooks/usePlans.js    |  70 +++++++++++++++-----------
 6 files changed, 191 insertions(+), 126 deletions(-)""",
    size=7.6,
)

# --------------------------------------------------------------------------- #
# 6. CAPTURAS DE LA APLICACIÃ“N
# --------------------------------------------------------------------------- #
page_break()
h1("6. Capturas de la aplicaciÃ³n")

para(
    "Las diez capturas siguientes se generaron con Playwright controlando Chrome real contra el servidor de desarrollo "
    "de Vite (`http://localhost:5173`). No son maquetas: cada una corresponde a un estado provocado en la aplicaciÃ³n "
    "y verificado en el DOM antes de disparar la captura. La imagen se recorta a la columna de la aplicaciÃ³n para que "
    "el texto siga siendo legible en el documento.",
    align="justify",
)
para(
    "El guion completo estÃ¡ en `docs/scripts/capturar_app.py` y puede reproducirse con `python docs/scripts/capturar_app.py` "
    "con el servidor levantado.",
    align="justify",
)

figure(CAP / "01_carga_esqueleto.png",
       "Estado `loading`: el esqueleto se muestra mientras la red simulada tarda 900 ms. El contador de la cabecera "
       "aparece como Â«â€¦Â».")
figure(CAP / "02_planes_cargados.png",
       "Estado `ready` con los datos reales: el contador Â«Pendientes: 6Â» y las primeras tarjetas de la bandeja.")
figure(CAP / "03_error_red.png",
       "Estado `error` abierto con `?fail=1`: aviso Â«No se pudieron cargar los planesÂ» y botÃ³n Â«ReintentarÂ».")
figure(CAP / "04_busqueda_antigua.png",
       "La bÃºsqueda acaba de escribir Â«valuesÂ». El campo muestra el texto al instante, pero la lista todavÃ­a no "
       "filtra: `useDebounce` espera 300 ms de calma.")
figure(CAP / "05_busqueda_filtrada.png",
       "DespuÃ©s de 300 ms sin teclear, la lista se reduce a un solo plan (P-03). Es el efecto del valor Â«tranquiloÂ».")
figure(CAP / "06_contador_pendientes.png",
       "Tras aprobar un plan R1, el contador de la cabecera baja a Â«Pendientes: 5Â» y el panel a Â«Decididos: 1Â». "
       "El nÃºmero se actualiza porque vive en `App`, no en la cabecera.")
figure(CAP / "07_ejecucion_r2_espera_aprobacion.png",
       "Un plan R2 (P-01) se ensayÃ³ correctamente, pero no se aplicÃ³: el panel pide una Ãºltima confirmaciÃ³n explÃ­cita "
       "con Â«Confirmar y aplicarÂ» o Â«DescartarÂ».")
figure(CAP / "08_ejecucion_aplicada.png",
       "DespuÃ©s de Â«Confirmar y aplicarÂ», el estado pasa a `done` y muestra Â«P-01 aplicado (exams_destroy ensayado "
       "sobre d68ebfa913ec)Â».")
figure(CAP / "09_persistencia_recarga.png",
       "Tras recargar la pÃ¡gina, las decisiones siguen en `localStorage`: Â«Pendientes: 4Â» y Â«Decididos: 2Â».")
figure(CAP / "10_vista_movil.png",
       "Vista mÃ³vil (390 Ã— 844): la bandeja, el buscador y las tarjetas se adaptan a una pantalla estrecha.",
       max_w=3.6, max_h=7.8)

# --------------------------------------------------------------------------- #
# 7. USO DE IA
# --------------------------------------------------------------------------- #
page_break()
h1("7. Uso de inteligencia artificial")

para(
    "Se usÃ³ un agente de codificaciÃ³n (OpenCode, modelo Â«Big PickleÂ») como asistente para leer el repositorio, "
    "proponer implementaciones, diagnosticar los seis fallos y preparar este documento. El agente trabajÃ³ bajo el "
    "contrato de `AGENTS.md`: solo editÃ³ `src/hooks/`, `src/ai/` y `docs/`; no tocÃ³ pruebas, configuraciÃ³n ni archivos "
    "de entorno. Cada cambio se verificÃ³ con las pruebas y con revisiÃ³n humana del diff.",
    align="justify",
)

h2("7.1 Tabla de uso de IA")
add_table(
    ["#", "Actividad", "Herramienta", "CÃ³mo se usÃ³", "Resultado", "VerificaciÃ³n"],
    [
        ["1", "Lectura del repositorio y plan", "OpenCode (Big Pickle)",
         "Leer `README.md`, `AGENTS.md`, las pruebas y los hooks; proponer un plan por partes.",
         "Plan de trabajo y restricciones", "RevisiÃ³n humana del plan"],
        ["2", "Implementar `useDebounce`", "OpenCode (Big Pickle)",
         "Escribir el hook segÃºn el `TODO`, explicando la limpieza del temporizador.",
         "Hook de 23 lÃ­neas", "10/10 pruebas; diff revisado"],
        ["3", "Implementar `usePlans`", "OpenCode (Big Pickle)",
         "Resolver los tres estados, `reload` y la bandera `cancelled` contra la respuesta vieja.",
         "Hook de 42 lÃ­neas", "12/12 pruebas; diff revisado"],
        ["4", "Implementar `useAgentRun`", "OpenCode (Big Pickle)",
         "Construir la mÃ¡quina de estados, la barrera R2 y el aborto al desmontar.",
         "Hook de 77 lÃ­neas", "18/18 pruebas; diff revisado"],
        ["5", "Diagnosticar los 6 fallos de `src/ai`", "OpenCode (Big Pickle)",
         "Correr `npm run test:ai`, leer los fallos y corregir sin tocar las pruebas.",
         "6 correcciones puntuales", "20 fallos â†’ 0; 57/57 en modo ai"],
        ["6", "Capturas reales", "Playwright + Chrome; script PrintWindow",
         "Automatizar capturas de la app, de la terminal y de VS Code.",
         "16 PNG", "OCR y revisiÃ³n visual; estados verificados en el DOM"],
        ["7", "RedacciÃ³n y maquetaciÃ³n", "OpenCode + `python-docx`",
         "Estructurar el informe, redactar las explicaciones y generar el DOCX/PDF.",
         "Este documento", "Cifras contrastadas con los logs"],
    ],
    widths_cm=[0.7, 3.1, 2.5, 4.5, 2.9, 2.9],
    font_size=7.9,
    align_center_cols=(0,),
)

h2("7.2 Lo que se aceptÃ³ y lo que se rechazÃ³")
para("**Se aceptÃ³:**", space_after=2)
bullet("La implementaciÃ³n de los tres hooks, despuÃ©s de leerla y comprobar que cada efecto limpia lo que abre.")
bullet("La clasificaciÃ³n de los seis fallos de la IA, contrastada una por una con la prueba que fallaba.")
bullet("Los guiones de captura y el texto del informe, revisados contra los resultados reales.")
para("**Se rechazÃ³:**", space_after=2)
bullet("Mover el contador de pendientes a `Header` Â«para simplificarÂ». Rompe la Ãºnica fuente de verdad y obliga a "
       "duplicar las decisiones; el ancestro comÃºn es `App`.")
bullet("Cambiar una prueba o `package.json` para que el resultado diera verde. Las pruebas son el contrato.")
bullet("Dejar el retorno temprano de `useAgentRun` cuando falta `runAgent`: cambia el orden de los hooks.")
bullet("Dejar el `.catch` con `console.error` en `usePlans`: la pantalla se quedaba cargando para siempre.")
bullet("Aplicar un R2 directo a `done`: la barrera de seguridad es el objetivo del laboratorio.")

h2("7.3 Las cuatro preguntas de la revisiÃ³n del diff")
add_table(
    ["Pregunta del contrato", "Respuesta", "Evidencia"],
    [
        ["1. Â¿Los hooks se llaman siempre en el mismo orden y en el nivel superior, sin `if`, bucles ni `return` antes?",
         "**SÃ­.** En `useAgentRun` todos los hooks se ejecutan siempre; el guard `if (!runAgent)` estÃ¡ dentro de "
         "`start`. El retorno temprano de la IA se eliminÃ³.",
         "Fallo 4 y prueba Â«si runAgent aparece en un dibujado posteriorâ€¦Â»"],
        ["2. Â¿Todo `useEffect` que abre algo lo limpia al repetirse y al desmontar?",
         "**SÃ­.** El temporizador con `clearTimeout`; la peticiÃ³n con la bandera `cancelled`; la ejecuciÃ³n con "
         "`AbortController.abort()`.",
         "Pruebas de desmontaje de los tres hooks"],
        ["3. Â¿Se resuelven los tres estados y se ven en pantalla?",
         "**SÃ­.** `loading` muestra el esqueleto, `ready` las tarjetas y `error` el aviso con Â«ReintentarÂ».",
         "Figuras 7, 8 y 9; pruebas de App"],
        ["4. Â¿El dato vive en el ancestro mÃ¡s bajo que lo necesita?",
         "**SÃ­.** `decisions` y `pendingCount` viven en `App`; `Header` solo recibe el nÃºmero y `Inbox` recibe props.",
         "SecciÃ³n 2.4 y prueba Â«la cabecera solo necesita un nÃºmeroâ€¦Â»"],
    ],
    widths_cm=[5.0, 6.5, 4.8],
    font_size=8.2,
)

# --------------------------------------------------------------------------- #
# 8. CONCLUSIONES
# --------------------------------------------------------------------------- #
h1("8. Conclusiones")
para(
    "El laboratorio confirma que los hooks no son Â«cÃ³digo de apoyoÂ»: son donde viven las reglas difÃ­ciles de una "
    "interfaz. Los tres hooks propios resuelven tres problemas distintos y complementarios: el tiempo (un valor que "
    "debe esperar), la red (tres estados y una carrera de peticiones) y el proceso (una mÃ¡quina de estados con una "
    "barrera de seguridad).",
    align="justify",
)
bullet(
    "**La limpieza es parte del contrato, no un detalle.** Un `useEffect` que abre un temporizador o una peticiÃ³n "
    "debe cerrarlo en su funciÃ³n de limpieza. Esa funciÃ³n corre al repetirse y al desmontar; si se coloca en un "
    "segundo efecto con `[]`, los recursos viejos sobreviven y aparecen valores intermedios o respuestas tardÃ­as."
)
bullet(
    "**Las carreras de peticiones se resuelven con una bandera por corrida.** No basta con cancelar el temporizador: "
    "una respuesta vieja puede llegar despuÃ©s de la nueva. Marcar cada corrida como cancelada y comprobarlo antes de "
    "escribir el estado es simple y cubre `reload`, `StrictMode` y el desmontaje."
)
bullet(
    "**Una mÃ¡quina de estados explÃ­cita evita los atajos peligrosos.** El estado `awaiting_approval` existe para que "
    "un plan destructivo no pueda pasar de `running` a `done` sin una decisiÃ³n humana. La prueba lo verifica; el "
    "cÃ³digo lo hace evidente."
)
bullet(
    "**La IA acelera, pero no decide.** Los seis fallos de `src/ai` son plausibles y fÃ¡ciles de pasar por alto: "
    "cÃ³digo que compila, se lee bien y falla en los bordes. La revisiÃ³n humana y las pruebas fueron las que "
    "determinaron quÃ© se aceptaba. Las cuatro preguntas del contrato son una lista de control eficaz."
)
bullet(
    "**El estado se eleva solo hasta donde hace falta.** `App` es el ancestro comÃºn de la cabecera y la bandeja; "
    "llevar las decisiones mÃ¡s arriba (a `main.jsx`) o mÃ¡s abajo (a `Header`) solo agrega duplicaciÃ³n o "
    "desincronizaciÃ³n."
)
para(
    "**Limitaciones y siguiente paso.** La red es simulada y las decisiones viven en el navegador; el servidor real y "
    "la autenticaciÃ³n llegan en la semana 12. Cuando eso ocurra, `usePlans` y `useAgentRun` podrÃ¡n recibir funciones "
    "que usen `fetch` con `AbortSignal` real: la estructura de estados y la limpieza ya estÃ¡n preparadas para ello.",
    align="justify",
)

# --------------------------------------------------------------------------- #
# ANEXO A
# --------------------------------------------------------------------------- #
page_break()
h1("Anexo A. CÃ³digo corregido de los hooks de la IA")
para("Los tres archivos de `src/ai/` quedaron asÃ­ despuÃ©s de la revisiÃ³n. Los comentarios `FIX:` seÃ±alan el cambio.",
     align="justify")

for rel, label in (
    ("src/ai/useDebounce.js", "src/ai/useDebounce.js"),
    ("src/ai/usePlans.js", "src/ai/usePlans.js"),
    ("src/ai/useAgentRun.js", "src/ai/useAgentRun.js"),
):
    h3(label)
    code((ROOT / rel).read_text(encoding="utf-8"), size=7.6)

# --------------------------------------------------------------------------- #
# ANEXO B
# --------------------------------------------------------------------------- #
page_break()
h1("Anexo B. CatÃ¡logo de las 57 pruebas")
para("Se listan los nombres reales de las pruebas, agrupados por archivo. Todas quedaron en verde en los dos modos.",
     align="justify")

test_rows = []
order = [
    ("useDebounce.test.jsx", "useDebounce"),
    ("usePlans.test.jsx", "usePlans"),
    ("useAgentRun.test.jsx", "useAgentRun"),
    ("App.test.jsx", "App"),
    ("Inbox.test.jsx", "Inbox"),
]
n = 0
for filename, _ in order:
    text = (ROOT / "src" / "test" / filename).read_text(encoding="utf-8")
    for match in re.finditer(r'\bit\("([^"]+)"', text):
        n += 1
        test_rows.append([str(n), filename, match.group(1)])

add_table(["#", "Archivo", "Prueba"], test_rows, widths_cm=[0.9, 3.5, 11.9], font_size=7.6,
          align_center_cols=(0,))
para(f"Total: {n} pruebas.", size=9, italic=True, align="right")

# --------------------------------------------------------------------------- #
# ANEXO C
# --------------------------------------------------------------------------- #
h1("Anexo C. Archivos, comandos y reproducibilidad")
add_table(
    ["Ruta", "QuÃ© contiene"],
    [
        ["`src/hooks/useDebounce.js`", "Hook propio: valor tranquilo con limpieza del temporizador"],
        ["`src/hooks/usePlans.js`", "Hook propio: tres estados, reload y respuesta vieja"],
        ["`src/hooks/useAgentRun.js`", "Hook propio: mÃ¡quina de estados y barrera R2"],
        ["`src/ai/*.js`", "Hooks de la IA corregidos (6 fallos)"],
        ["`docs/capturas/`", "16 capturas reales (app, terminal y VS Code)"],
        ["`docs/logs/`", "Salidas completas de pruebas y build"],
        ["`docs/scripts/capturar_app.py`", "Guion de capturas de la aplicaciÃ³n (Playwright + Chrome)"],
        ["`docs/scripts/capturar_ventana.ps1`", "Captura de ventanas de terminal"],
        ["`docs/scripts/capturar_vscode.ps1`", "Captura de la ventana de VS Code"],
        ["`docs/scripts/generar_entregable.py`", "Generador de este DOCX"],
    ],
    widths_cm=[5.6, 10.7],
    font_size=8.5,
)

para("**Reproducir las pruebas:**", space_after=2)
code(
    """npm install
npm test          # 57/57 en modo team
npm run test:ai   # 57/57 en modo ai
npm run build     # compilaciÃ³n de producciÃ³n""",
)
para("**Reproducir las capturas de la aplicaciÃ³n:**", space_after=2)
code(
    """npm run dev       # deja el servidor en http://localhost:5173
python docs/scripts/capturar_app.py""",
)

para(
    "Fin del documento. Estudiante: Edu Edward Galindo HuamÃ¡n Â· DAE Semana 11 (extra) Â· 2026.",
    size=9, italic=True, color=GRAY, align="center", space_before=12,
)

# --------------------------------------------------------------------------- #
# Guardar
# --------------------------------------------------------------------------- #
doc.save(OUT)
print("DOCX generado:", OUT)
print("Figuras:", FIG["n"])
print("Pruebas en el catÃ¡logo:", n)

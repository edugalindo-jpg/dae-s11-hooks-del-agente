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
# Página y estilos
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
    """Escribe texto con **negrita** y `código` en línea (también dentro de la negrita)."""
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
    """Incluye un archivo de código con un rótulo."""
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
    run = p.add_run("Edu Edward Galindo Huamán · DAE Semana 11 · Los hooks del agente · pág. ")
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
    hr = hp.add_run("Laboratorio S11 (extra) · Los hooks del agente")
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
p = para("Implementación de useDebounce, usePlans y useAgentRun", size=12.5, color=DARK, align="center", space_after=2)
p = para("y revisión de seis fallos en hooks escritos por una IA", size=12.5, color=DARK, align="center", space_after=30)

# Línea decorativa
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
    ("Estudiante", "Edu Edward Galindo Huamán"),
    ("Curso", "Desarrollo de Aplicaciones Empresariales (DAE)"),
    ("Fecha", "9 de noviembre de 2026"),
    ("Repositorio", "https://github.com/edugalindo-jpg/dae-s11-hooks-del-agente"),
    ("Entorno", "Node.js v24.21.0 · npm 11.19.0 · Vite 7.3.7 · Vitest 3.2.7"),
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
p = para("57 / 57 pruebas en modo team  ·  57 / 57 pruebas en modo ai  ·  build correcto",
         size=12, bold=True, color=DARK, align="center", space_after=4)
p = para("16 capturas reales: 10 de la aplicación, 5 de terminal y 1 de VS Code",
         size=10, color=GRAY, align="center", space_after=0)

page_break()

# --------------------------------------------------------------------------- #
# ÍNDICE (campo TOC; Word lo actualiza al exportar)
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

placeholder_run = p.add_run("Índice en preparación.")
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
    "La aplicación es una bandeja de aprobaciones: un agente de IA propone planes con un `plan_hash` y una persona "
    "decide si los aprueba o los rechaza. La bandeja carga los planes con una petición simulada, busca con retardo y "
    "permite ejecutar lo aprobado; un plan destructivo (R2) no se aplica sin una confirmación explícita.",
    align="justify",
)
para(
    "El repositorio de partida traía los componentes y las pruebas ya escritos, pero los tres hooks centrales estaban "
    "sin implementar: por eso la aplicación se quedaba cargando y 40 de las 57 pruebas fallaban. El trabajo tuvo dos "
    "frentes: (1) implementar `useDebounce`, `usePlans` y `useAgentRun` contra las pruebas, y (2) revisar tres hooks "
    "«escritos por una IA» que compilaban y parecían correctos, pero contenían seis fallos típicos de React.",
    align="justify",
)
para(
    "No se modificaron las pruebas, `package.json`, `vite.config.js` ni los archivos `.env.*`. Todo el cambio de "
    "código está en `src/hooks/` y `src/ai/`. Las capturas se tomaron de la aplicación real corriendo en el servidor "
    "de desarrollo y de la terminal real; no hay maquetas ni imágenes simuladas.",
    align="justify",
)

h2("1.1 Resultados")
add_table(
    ["Indicador", "Estado inicial", "Estado final"],
    [
        ["`npm test` (modo team, `src/hooks`)", "40 fallan · 17 pasan (57)", "**57 pasan (57)**"],
        ["`npm run test:ai` (modo ai, `src/ai`)", "20 fallan · 37 pasan (57)", "**57 pasan (57)**"],
        ["`npm run build`", "no se había ejecutado", "**correcto (44 módulos)**"],
        ["Capturas reales", "—", "**16** (10 app + 5 terminal + 1 VS Code)"],
    ],
    widths_cm=[6.4, 5.2, 5.0],
    font_size=9,
)

h2("1.2 Qué contiene este entregable")
for item in (
    "La explicación de la arquitectura y de **por qué el contador de pendientes vive en `App` y no en `Header`** "
    "(sección 2.4), con el código que lo demuestra y las pruebas que lo respaldan.",
    "El código final comentado de los tres hooks propios y las decisiones de diseño de cada uno (sección 3).",
    "La revisión de los hooks de la IA: los **seis fallos** clasificados, el «antes y después» de cada corrección y "
    "la evidencia de las pruebas (sección 4).",
    "Las tablas de pruebas, los comandos y las **capturas reales** de terminal, VS Code y aplicación (secciones 5 y 6).",
    "La **tabla de uso de IA**, lo que se aceptó, lo que se rechazó y las cuatro preguntas de revisión del diff "
    "(sección 7).",
    "Los anexos con el código corregido de `src/ai/` y el catálogo completo de las 57 pruebas.",
):
    bullet(item)

# --------------------------------------------------------------------------- #
# 2. APLICACIÓN Y PROPIEDAD DEL ESTADO
# --------------------------------------------------------------------------- #
h1("2. La aplicación y la propiedad del estado")

h2("2.1 Qué hace la aplicación")
para(
    "`App` pide los planes al agente con `usePlans` (red simulada con 900 ms de retardo y un modo `?fail=1` que "
    "rechaza). Mientras llegan, muestra un esqueleto de carga; si la petición falla, muestra un aviso con "
    "«Reintentar»; cuando llegan, dibuja la bandeja. Cada plan se puede aprobar o rechazar. Aprobar un plan R2 exige "
    "escribir su hash exacto. Las decisiones se guardan en `localStorage` y sobreviven a una recarga. Los planes "
    "aprobados aparecen en el panel de ejecución, que usa `useAgentRun` para ensayar primero y aplicar después.",
    align="justify",
)

h2("2.2 Árbol de componentes")
code(
    """App                                  ← usePlans + estado decisions + persistencia
├── Header({ pendingCount })           ← solo dibuja el número
├── toolbar                            ← simular fallo / recargar
└── Inbox({ plans, decisions, onDecide, runAgent })
    ├── searchbox                      ← useState(query) + useDebounce(query, 300)
    ├── PlanCard({ plan, onDecide })   ← aprobar / rechazar
    │   └── HashConfirm                ← input del hash para R2
    └── RunPanel({ approvedPlans, runAgent })
        └── useAgentRun                ← idle · running · awaiting_approval · done · error""",
    size=7.8,
)

figure(
    CAP / "16_vscode_estructura.png",
    "Estructura real del proyecto en VS Code: el árbol de archivos del repositorio y `src/hooks/usePlans.js` abierto.",
)

h2("2.3 Dónde vive cada dato")
add_table(
    ["Dato", "Dónde vive", "Quién lo escribe", "Quién lo lee"],
    [
        ["`status`, `plans`, `error`", "`usePlans`, invocado en `App`", "La promesa de `fetchPlans`", "`App` (decide qué dibujar) e `Inbox`"],
        ["`decisions`", "`App`: `useState` + `localStorage`", "`decide()` cuando `Inbox` avisa", "`Header` (conteo) e `Inbox` (listas)"],
        ["`pendingCount`", "**Derivado en `App`**, no es estado", "Se recalcula en cada render", "`Header`, como número"],
        ["`query`", "`Inbox` (`useState`)", "El campo de búsqueda", "`Inbox`, a través de `useDebounce`"],
        ["`run`", "`useAgentRun`, invocado en `RunPanel`", "`start`, `approve`, `reject`", "`RunPanel` (estado de la ejecución)"],
    ],
    widths_cm=[3.2, 4.4, 4.4, 4.6],
    font_size=8.3,
)

h2("2.4 Por qué el contador de pendientes vive en App y no en Header")
para(
    "El contador de pendientes **no es un dato independiente**: es el resultado de cruzar la lista de planes con el "
    "mapa de decisiones. Por eso no se guarda: se deriva. `App` ya tiene las dos entradas (los planes que devuelve "
    "`usePlans` y las decisiones que administra con `useState`), así que allí se calcula una sola vez y se pasa a "
    "`Header` únicamente el número.",
    align="justify",
)
code(
    """// App.jsx — las decisiones viven en el ancestro común más bajo
const { status, plans, error, reload } = usePlans(fetchPlans);
const [decisions, setDecisions] = useState(readStoredDecisions);

// Derivado en cada render, nunca almacenado. La cabecera recibe solo este número.
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
    "**El ancestro común más bajo es `App`.** `Header` e `Inbox` son hermanos. Las decisiones se producen al decidir "
    "dentro de `Inbox` (en `PlanCard`) y las necesitan tanto `Inbox` (para separar pendientes de decididos) como "
    "`Header` (para contar). El único componente que es ancestro de ambos es `App`; si el contador viviera en "
    "`Header`, las decisiones tendrían que subir de todas formas y se duplicaría el camino."
)
bullet(
    "**Una sola fuente de verdad.** `decisions` se guarda una vez y se persiste en `localStorage`. El contador es una "
    "vista de ese estado. Si `Header` tuviera su propio contador, habría dos copias que podrían desincronizarse al "
    "recargar, al rechazar o al cambiar la lista de planes."
)
bullet(
    "**`Header` es presentacional.** Recibe solo `pendingCount` y no la lista. No se vuelve a dibujar por cambios en "
    "los planes que no le interesan y se puede probar aislado. La prueba «la cabecera solo necesita un número: se "
    "dibuja sin la lista de planes» lo verifica explícitamente."
)
bullet(
    "**La persistencia encaja sin esfuerzo.** Al recargar, `App` lee las decisiones de `localStorage` y el contador se "
    "recalcula solo. Un contador dentro de `Header` no sabría cuándo cambió la lista ni sobreviviría a la recarga."
)
bullet(
    "**Menos complejidad, no más.** La alternativa sería un contexto global o un estado en `Header` actualizado por "
    "eventos. Con solo dos consumidores que ya cuelgan de `App`, eso agrega indirección sin beneficio. La regla del "
    "curso es «el dato vive en el ancestro más bajo que lo necesita», y ese ancestro es `App`, no `main.jsx`."
)

callout(
    "Respuesta corta:",
    "el contador vive en `App` porque allí están los planes y las decisiones, que son sus dos entradas; `Header` solo "
    "lo muestra. Guardarlo en la cabecera duplicaría el estado y rompería la única fuente de verdad.",
)

# --------------------------------------------------------------------------- #
# 3. HOOKS PROPIOS
# --------------------------------------------------------------------------- #
page_break()
h1("3. Implementación de los tres hooks propios")

para(
    "Los tres hooks se implementaron contra las pruebas, sin modificarlas. Cada archivo conserva las convenciones del "
    "curso: código y nombres en inglés, comentarios en inglés, y explicación en español en este informe. Los hooks se "
    "llaman siempre en el nivel superior y en el mismo orden.",
    align="justify",
)

# 3.1 useDebounce
h2("3.1 useDebounce: el valor tranquilo")
para(
    "**Contrato.** `useDebounce(value, delay)` devuelve el último `value` que lleva `delay` milisegundos sin cambiar. "
    "En el primer dibujado devuelve `value` de inmediato. Cada cambio cancela el temporizador anterior; si el retardo "
    "cambia, se vuelve a contar. Al desmontar no queda ningún temporizador vivo.",
    align="justify",
)
source_block(ROOT / "src/hooks/useDebounce.js", "src/hooks/useDebounce.js")

para("**Decisiones de diseño.**", space_after=2)
bullet(
    "`useState(value)` arranca con el valor actual: la interfaz no espera `delay` ms en el primer dibujado."
)
bullet(
    "El temporizador se crea dentro del `useEffect` con dependencias `[value, delay]`. React ejecuta la limpieza "
    "**antes** de repetir el efecto, así que `clearTimeout(timer)` cancela el temporizador anterior en cada cambio. "
    "Solo sobrevive el último; los valores intermedios nunca llegan a `setDebounced`."
)
bullet(
    "La misma limpieza corre al desmontar. No hace falta un `useRef`: el temporizador pertenece a una corrida concreta "
    "del efecto y la clausura lo captura. Guardarlo en un `ref` fue precisamente el error del hook de la IA."
)
bullet(
    "Si `delay` cambia, el efecto se repite y la cuenta empieza de cero con el nuevo retardo, tal como exige la prueba."
)
para(
    "**Pruebas (10/10).** El archivo `useDebounce.test.jsx` agrupa: el valor tranquilo (3), la cancelación del "
    "temporizador anterior (4) y la limpieza al desmontar (3). Usa `vi.useFakeTimers()` y `act()` para mover el reloj; "
    "no espera en tiempo real.",
    align="justify",
)

# 3.2 usePlans
h2("3.2 usePlans: los tres estados y la respuesta vieja")
para(
    "**Contrato.** `usePlans(fetchPlans)` devuelve `{ status, plans, error, reload }`. `status` es `loading`, `ready` "
    "o `error`. Mientras carga y si falla, `plans` es `[]`; en `error`, `error` es el motivo del rechazo. `reload()` "
    "vuelve a pedir y regresa a `loading`. La respuesta de una petición vieja nunca debe pisar a la nueva.",
    align="justify",
)
source_block(ROOT / "src/hooks/usePlans.js", "src/hooks/usePlans.js")

para("**Decisiones de diseño.**", space_after=2)
bullet(
    "El estado se agrupa en un solo objeto (`INITIAL` con `status`, `plans` y `error`) para que los tres estados se "
    "actualicen juntos y no existan combinaciones imposibles, como `status: \"ready\"` con un error viejo."
)
bullet(
    "`reload` incrementa un contador `attempt` en lugar de modificar el estado directamente. Así el efecto se repite "
    "sin duplicar la lógica de carga, y `reload` conserva la misma identidad entre renders gracias a `useCallback`."
)
bullet(
    "Cada corrida del efecto crea una bandera `cancelled`. La limpieza la pone en `true`; antes de escribir el estado "
    "se comprueba. Esto cubre cuatro casos que las pruebas ejercitan: una respuesta vieja que llega después de la "
    "nueva, un error viejo que llega tarde, un cambio de `fetchPlans` y el doble efecto de `StrictMode`."
)
bullet(
    "La bandera también cubre el desmontaje: si la respuesta llega después, no se llama `setState`. No se usa "
    "`AbortController` aquí porque la función `fetchPlans` de la aplicación no recibe una señal; la bandera es "
    "suficiente y no cambia el contrato."
)
para(
    "**Pruebas (12/12).** `usePlans.test.jsx` cubre los tres estados (4), `reload` (3), la carrera de peticiones (4) y "
    "el desmontaje (1). Las promesas son «deferred»: la prueba decide cuándo y en qué orden responde cada petición.",
    align="justify",
)

# 3.3 useAgentRun
h2("3.3 useAgentRun: la máquina de estados de una ejecución")
para(
    "**Contrato.** `useAgentRun(runAgent)` devuelve `{ status, plan, result, error, start, approve, reject }`. La "
    "ejecución empieza en `idle`; `start(plan)` pasa a `running`; un plan R1 termina en `done`; un plan R2 termina en "
    "`awaiting_approval` y solo `approve()` lo lleva a `done`; un fallo lleva a `error`; `reject()` cancela y vuelve a "
    "`idle`. Si ya hay una ejecución en curso, `start()` no hace nada.",
    align="justify",
)
code(
    """idle ──start(plan)──▶ running ──ok, R1──────────────▶ done
                          │
                          ├──ok, R2──▶ awaiting_approval ──approve()──▶ done
                          │
                          └──falla────▶ error

reject() en running o en awaiting_approval: aborta y vuelve a idle.
start() se ignora en running y en awaiting_approval.""",
    size=8.2,
)
source_block(ROOT / "src/hooks/useAgentRun.js", "src/hooks/useAgentRun.js")

para("**Decisiones de diseño.**", space_after=2)
bullet(
    "`latest` es un `ref` que refleja el último estado. Los manejadores `start`, `approve` y `reject` lo leen para "
    "decidir sin depender de una clausura vieja, incluso si se llaman varias veces antes de que React vuelva a dibujar."
)
bullet(
    "`controller` guarda el `AbortController` de la ejecución actual. `start` crea uno nuevo y pasa su señal a "
    "`runAgent`; `reject` y el desmontaje llaman a `abort()`. La promesa comprueba `signal.aborted` antes de tocar el "
    "estado, así que una respuesta tardía no puede revivir una ejecución cancelada."
)
bullet(
    "**Regla de seguridad:** el éxito de un plan R2 no va a `done`; va a `awaiting_approval` conservando el resultado. "
    "Solo `approve()` cambia el estado a `done`, y solo funciona si el estado actual es `awaiting_approval`."
)
bullet(
    "Todos los hooks se ejecutan siempre, antes de cualquier `return`. La comprobación de `runAgent` ausente está "
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
h1("4. Revisión de los hooks «escritos por una IA»")

para(
    "La carpeta `src/ai/` contiene los mismos tres hooks, escritos por una IA. Compilan, se leen bien y la aplicación "
    "funciona en parte, pero esconden seis fallos típicos. El modo `npm run test:ai` corre exactamente las mismas 57 "
    "pruebas contra esa carpeta: al empezar, **20 fallan y 37 pasan**.",
    align="justify",
)

h2("4.1 Método de revisión")
bullet("Se ejecutó `npm run test:ai` y se guardó la salida completa en `docs/logs/02_test_ai_inicial.txt`.")
bullet("Se leyó cada prueba fallida y se buscó la causa en el hook, no en la prueba. Las pruebas no se tocaron.")
bullet("Se corrigió un archivo por turno y se volvió a correr el modo `ai` después de cada cambio.")
bullet("Se clasificaron los fallos en seis causas raíz, que se documentan en la tabla siguiente.")
bullet("Al final, `npm test` y `npm run test:ai` quedaron los dos en **57/57**.")

h2("4.2 Los seis fallos")
add_table(
    ["#", "Archivo", "Fallo y síntoma", "Corrección aplicada", "Pruebas que lo demuestran"],
    [
        ["1", "`src/ai/useDebounce.js`",
         "El temporizador solo se limpia al desmontar. Un segundo `useEffect` con `[]` guarda el timer en un `ref`; "
         "al cambiar el valor, el timer anterior sigue vivo y emite valores intermedios.",
         "Devolver `clearTimeout(timer)` desde el efecto `[value, delay]`, sin `ref`.",
         "5 pruebas de «cada cambio cancela el temporizador anterior» y «limpieza al desmontar»."],
        ["2", "`src/ai/usePlans.js`",
         "El rechazo no cambia el estado: el `.catch` solo hace `console.error`. El status se queda en `loading` para "
         "siempre y `error` sigue en `null`.",
         "En el rechazo: `setPlans([])`, `setError(err)` y `setStatus(\"error\")` (con guard).",
         "«si la petición falla pasa a error…», «después de un error, reload()…» y la prueba de App con red caída."],
        ["3", "`src/ai/usePlans.js`",
         "No hay bandera de cancelación: la respuesta de un `reload` anterior (o posterior al desmontaje) puede pisar "
         "la más nueva. Es la carrera que las pruebas provocan a propósito.",
         "`let cancelled = false` por corrida del efecto; la limpieza lo pone en `true`; se comprueba antes de cada "
         "`setState`.",
         "Las 3 pruebas de «ignora la respuesta vieja» que fallaban (respuesta tardía, cambio de `fetchPlans` y "
         "`StrictMode`)."],
        ["4", "`src/ai/useAgentRun.js`",
         "Hay un `return` temprano antes de los hooks: `if (!runAgent) return { ... }`. Eso cambia el número de hooks "
         "entre dibujados; cuando `runAgent` aparece, React lanza «Rendered more hooks than during the previous render».",
         "Quitar el retorno temprano; mantener todos los hooks y comprobar `if (!runAgent) return;` **dentro** de "
         "`start`.",
         "«sin runAgent se queda en idle…» y «si runAgent aparece en un dibujado posterior, el hook no se rompe»."],
        ["5", "`src/ai/useAgentRun.js`",
         "Un plan R2 se aplica sin aprobación: el éxito siempre hace `update({ status: \"done\" })`. No existe la "
         "barrera de seguridad de la semana 10.",
         "Si `plan.risk === \"R2\"`, pasar a `awaiting_approval` conservando el resultado; solo `approve()` lleva a "
         "`done`.",
         "4 pruebas de la barrera R2, «esperando aprobación tampoco se puede lanzar otra ejecución» y las 2 de Inbox "
         "con R2."],
        ["6", "`src/ai/useAgentRun.js`",
         "No hay limpieza al desmontar: la señal no se aborta y una respuesta tardía intenta actualizar el estado de "
         "un componente que ya no existe.",
         "`useEffect(() => () => controller.current?.abort(), [])`.",
         "«al desmontar con una ejecución en curso aborta su señal» y «si la respuesta llega después de desmontar, no "
         "pasa nada»."],
    ],
    widths_cm=[0.7, 3.0, 5.0, 4.4, 3.5],
    font_size=7.9,
    align_center_cols=(0,),
)

h2("4.3 Antes y después de cada corrección")
para(
    "Los fragmentos siguientes muestran el cambio mínimo que resuelve cada fallo. El código completo corregido está "
    "en el anexo A.",
    align="justify",
)

h3("Fallo 1 · limpieza del temporizador")
code(
    """// ANTES (src/ai/useDebounce.js)
useEffect(() => {
  timer.current = setTimeout(() => setDebounced(value), delay);
}, [value, delay]);
useEffect(() => () => clearTimeout(timer.current), []);   // solo al desmontar

// DESPUÉS
useEffect(() => {
  const timer = setTimeout(() => setDebounced(value), delay);
  return () => clearTimeout(timer);                        // al cambiar y al desmontar
}, [value, delay]);""",
)

h3("Fallos 2 y 3 · estado de error y respuesta vieja")
code(
    """// ANTES (src/ai/usePlans.js)
fetchPlans()
  .then((data) => { setPlans(data); setStatus("ready"); })
  .catch((err) => { console.error("Error loading plans", err); });   // nunca hay estado "error"

// DESPUÉS
let cancelled = false;                        // una bandera por corrida del efecto
fetchPlans().then(
  (data) => { if (cancelled) return; setPlans(data); setStatus("ready"); setError(null); },
  (err)  => { if (cancelled) return; setPlans([]); setError(err); setStatus("error"); }
);
return () => { cancelled = true; };           // reload, cambio de fetchPlans y desmontaje""",
)

h3("Fallo 4 · orden de los hooks")
code(
    """// ANTES (src/ai/useAgentRun.js)
if (!runAgent) {
  return { ...IDLE, start: () => {}, approve: () => {}, reject: () => {} };  // hooks condicionales
}
const latest = useRef(IDLE);   // este hook se salta en el primer dibujado

// DESPUÉS
const latest = useRef(IDLE);   // todos los hooks siempre, en el mismo orden
const start = useCallback((plan) => {
  if (!runAgent) return;       // la comprobación va dentro del manejador
  ...
}, [runAgent, update]);""",
)

h3("Fallo 5 · la barrera del R2")
code(
    """// ANTES
runAgent(plan, { signal }).then((result) => {
  update({ status: "done", plan, result, error: null });     // R2 aplicado sin aprobación
});

// DESPUÉS
runAgent(plan, { signal }).then((result) => {
  if (plan.risk === "R2") {
    update({ status: "awaiting_approval", plan, result, error: null });
  } else {
    update({ status: "done", plan, result, error: null });
  }
});""",
)

h3("Fallo 6 · limpieza al desmontar")
code(
    """// ANTES: no había ningún efecto de limpieza.

// DESPUÉS
useEffect(() => {
  return () => {
    controller.current?.abort();   // la promesa comprueba signal.aborted antes de setState
    controller.current = null;
  };
}, []);""",
)

para(
    "**Resultado:** `npm run test:ai` pasó de **20 fallan / 37 pasan** a **57 pasan**. La salida completa está en "
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
    ["Suite", "Qué verifica", "N.º", "Inicial", "Final"],
    [
        ["`useDebounce.test.jsx`", "Valor tranquilo, cancelación del timer anterior y limpieza al desmontar", "10", "5 fallan", "10 pasan"],
        ["`usePlans.test.jsx`", "Tres estados, `reload`, carrera de peticiones y desmontaje", "12", "11 fallan", "12 pasan"],
        ["`useAgentRun.test.jsx`", "Máquina de estados, barrera R2, cancelación, error y limpieza", "18", "14 fallan", "18 pasan"],
        ["`App.test.jsx`", "Carga, error, contador en `App` y persistencia", "7", "5 fallan", "7 pasan"],
        ["`Inbox.test.jsx`", "Búsqueda con debounce, decisiones y ejecución", "10", "5 fallan", "10 pasan"],
        ["**Total**", "**Las 57 pruebas del curso**", "**57**", "**40 fallan · 17 pasan**", "**57 pasan**"],
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
npm run build               # compilación de producción (vite build)
npm run dev                 # aplicación real para las capturas (http://localhost:5173)""",
)

h2("5.3 Evidencia de terminal")
para(
    "Las capturas siguientes se tomaron de ventanas de PowerShell reales, con la salida real de cada comando y el "
    "código de salida visible. Para que el texto fuera legible en el documento se usó una fuente de consola de 28 px; "
    "no se editó el contenido.",
    align="justify",
)
figure(CAP / "11_terminal_test_antes.png",
       "Estado inicial de `npm test`: 40 pruebas fallan y 17 pasan, con código de salida 1.")
figure(CAP / "12_terminal_test_ai_antes.png",
       "Estado inicial de `npm run test:ai`: 20 fallan y 37 pasan. Son los seis fallos de los hooks de la IA.")
figure(CAP / "13_terminal_test_despues.png",
       "Estado final de `npm test`: los 5 archivos y las 57 pruebas en verde, con código de salida 0.")
figure(CAP / "14_terminal_test_ai_despues.png",
       "Estado final de `npm run test:ai`: las mismas 57 pruebas en verde sobre `src/ai`.")
figure(CAP / "15_terminal_build.png",
       "`npm run build`: compilación de producción correcta, 44 módulos transformados.")

h2("5.4 Registros completos")
para(
    "Además de las capturas, el repositorio guarda la salida completa de cada corrida en `docs/logs/`. Estos archivos "
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

h2("5.5 Qué no se modificó")
para(
    "El contrato del repositorio prohíbe tocar las pruebas, la configuración y los archivos de entorno. El diff de "
    "código afecta únicamente a los seis hooks. Las pruebas, `package.json`, `vite.config.js` y `.env.*` siguen "
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
# 6. CAPTURAS DE LA APLICACIÓN
# --------------------------------------------------------------------------- #
page_break()
h1("6. Capturas de la aplicación")

para(
    "Las diez capturas siguientes se generaron con Playwright controlando Chrome real contra el servidor de desarrollo "
    "de Vite (`http://localhost:5173`). No son maquetas: cada una corresponde a un estado provocado en la aplicación "
    "y verificado en el DOM antes de disparar la captura. La imagen se recorta a la columna de la aplicación para que "
    "el texto siga siendo legible en el documento.",
    align="justify",
)
para(
    "El guion completo está en `docs/scripts/capturar_app.py` y puede reproducirse con `python docs/scripts/capturar_app.py` "
    "con el servidor levantado.",
    align="justify",
)

figure(CAP / "01_carga_esqueleto.png",
       "Estado `loading`: el esqueleto se muestra mientras la red simulada tarda 900 ms. El contador de la cabecera "
       "aparece como «…».")
figure(CAP / "02_planes_cargados.png",
       "Estado `ready` con los datos reales: el contador «Pendientes: 6» y las primeras tarjetas de la bandeja.")
figure(CAP / "03_error_red.png",
       "Estado `error` abierto con `?fail=1`: aviso «No se pudieron cargar los planes» y botón «Reintentar».")
figure(CAP / "04_busqueda_antigua.png",
       "La búsqueda acaba de escribir «values». El campo muestra el texto al instante, pero la lista todavía no "
       "filtra: `useDebounce` espera 300 ms de calma.")
figure(CAP / "05_busqueda_filtrada.png",
       "Después de 300 ms sin teclear, la lista se reduce a un solo plan (P-03). Es el efecto del valor «tranquilo».")
figure(CAP / "06_contador_pendientes.png",
       "Tras aprobar un plan R1, el contador de la cabecera baja a «Pendientes: 5» y el panel a «Decididos: 1». "
       "El número se actualiza porque vive en `App`, no en la cabecera.")
figure(CAP / "07_ejecucion_r2_espera_aprobacion.png",
       "Un plan R2 (P-01) se ensayó correctamente, pero no se aplicó: el panel pide una última confirmación explícita "
       "con «Confirmar y aplicar» o «Descartar».")
figure(CAP / "08_ejecucion_aplicada.png",
       "Después de «Confirmar y aplicar», el estado pasa a `done` y muestra «P-01 aplicado (exams_destroy ensayado "
       "sobre d68ebfa913ec)».")
figure(CAP / "09_persistencia_recarga.png",
       "Tras recargar la página, las decisiones siguen en `localStorage`: «Pendientes: 4» y «Decididos: 2».")
figure(CAP / "10_vista_movil.png",
       "Vista móvil (390 × 844): la bandeja, el buscador y las tarjetas se adaptan a una pantalla estrecha.",
       max_w=3.6, max_h=7.8)

# --------------------------------------------------------------------------- #
# 7. USO DE IA
# --------------------------------------------------------------------------- #
page_break()
h1("7. Uso de inteligencia artificial")

para(
    "Se usó un agente de codificación (OpenCode, modelo «Big Pickle») como asistente para leer el repositorio, "
    "proponer implementaciones, diagnosticar los seis fallos y preparar este documento. El agente trabajó bajo el "
    "contrato de `AGENTS.md`: solo editó `src/hooks/`, `src/ai/` y `docs/`; no tocó pruebas, configuración ni archivos "
    "de entorno. Cada cambio se verificó con las pruebas y con revisión humana del diff.",
    align="justify",
)

h2("7.1 Tabla de uso de IA")
add_table(
    ["#", "Actividad", "Herramienta", "Cómo se usó", "Resultado", "Verificación"],
    [
        ["1", "Lectura del repositorio y plan", "OpenCode (Big Pickle)",
         "Leer `README.md`, `AGENTS.md`, las pruebas y los hooks; proponer un plan por partes.",
         "Plan de trabajo y restricciones", "Revisión humana del plan"],
        ["2", "Implementar `useDebounce`", "OpenCode (Big Pickle)",
         "Escribir el hook según el `TODO`, explicando la limpieza del temporizador.",
         "Hook de 23 líneas", "10/10 pruebas; diff revisado"],
        ["3", "Implementar `usePlans`", "OpenCode (Big Pickle)",
         "Resolver los tres estados, `reload` y la bandera `cancelled` contra la respuesta vieja.",
         "Hook de 42 líneas", "12/12 pruebas; diff revisado"],
        ["4", "Implementar `useAgentRun`", "OpenCode (Big Pickle)",
         "Construir la máquina de estados, la barrera R2 y el aborto al desmontar.",
         "Hook de 77 líneas", "18/18 pruebas; diff revisado"],
        ["5", "Diagnosticar los 6 fallos de `src/ai`", "OpenCode (Big Pickle)",
         "Correr `npm run test:ai`, leer los fallos y corregir sin tocar las pruebas.",
         "6 correcciones puntuales", "20 fallos → 0; 57/57 en modo ai"],
        ["6", "Capturas reales", "Playwright + Chrome; script PrintWindow",
         "Automatizar capturas de la app, de la terminal y de VS Code.",
         "16 PNG", "OCR y revisión visual; estados verificados en el DOM"],
        ["7", "Redacción y maquetación", "OpenCode + `python-docx`",
         "Estructurar el informe, redactar las explicaciones y generar el DOCX/PDF.",
         "Este documento", "Cifras contrastadas con los logs"],
    ],
    widths_cm=[0.7, 3.1, 2.5, 4.5, 2.9, 2.9],
    font_size=7.9,
    align_center_cols=(0,),
)

h2("7.2 Lo que se aceptó y lo que se rechazó")
para("**Se aceptó:**", space_after=2)
bullet("La implementación de los tres hooks, después de leerla y comprobar que cada efecto limpia lo que abre.")
bullet("La clasificación de los seis fallos de la IA, contrastada una por una con la prueba que fallaba.")
bullet("Los guiones de captura y el texto del informe, revisados contra los resultados reales.")
para("**Se rechazó:**", space_after=2)
bullet("Mover el contador de pendientes a `Header` «para simplificar». Rompe la única fuente de verdad y obliga a "
       "duplicar las decisiones; el ancestro común es `App`.")
bullet("Cambiar una prueba o `package.json` para que el resultado diera verde. Las pruebas son el contrato.")
bullet("Dejar el retorno temprano de `useAgentRun` cuando falta `runAgent`: cambia el orden de los hooks.")
bullet("Dejar el `.catch` con `console.error` en `usePlans`: la pantalla se quedaba cargando para siempre.")
bullet("Aplicar un R2 directo a `done`: la barrera de seguridad es el objetivo del laboratorio.")

h2("7.3 Las cuatro preguntas de la revisión del diff")
add_table(
    ["Pregunta del contrato", "Respuesta", "Evidencia"],
    [
        ["1. ¿Los hooks se llaman siempre en el mismo orden y en el nivel superior, sin `if`, bucles ni `return` antes?",
         "**Sí.** En `useAgentRun` todos los hooks se ejecutan siempre; el guard `if (!runAgent)` está dentro de "
         "`start`. El retorno temprano de la IA se eliminó.",
         "Fallo 4 y prueba «si runAgent aparece en un dibujado posterior…»"],
        ["2. ¿Todo `useEffect` que abre algo lo limpia al repetirse y al desmontar?",
         "**Sí.** El temporizador con `clearTimeout`; la petición con la bandera `cancelled`; la ejecución con "
         "`AbortController.abort()`.",
         "Pruebas de desmontaje de los tres hooks"],
        ["3. ¿Se resuelven los tres estados y se ven en pantalla?",
         "**Sí.** `loading` muestra el esqueleto, `ready` las tarjetas y `error` el aviso con «Reintentar».",
         "Figuras 7, 8 y 9; pruebas de App"],
        ["4. ¿El dato vive en el ancestro más bajo que lo necesita?",
         "**Sí.** `decisions` y `pendingCount` viven en `App`; `Header` solo recibe el número y `Inbox` recibe props.",
         "Sección 2.4 y prueba «la cabecera solo necesita un número…»"],
    ],
    widths_cm=[5.0, 6.5, 4.8],
    font_size=8.2,
)

# --------------------------------------------------------------------------- #
# 8. CONCLUSIONES
# --------------------------------------------------------------------------- #
h1("8. Conclusiones")
para(
    "El laboratorio confirma que los hooks no son «código de apoyo»: son donde viven las reglas difíciles de una "
    "interfaz. Los tres hooks propios resuelven tres problemas distintos y complementarios: el tiempo (un valor que "
    "debe esperar), la red (tres estados y una carrera de peticiones) y el proceso (una máquina de estados con una "
    "barrera de seguridad).",
    align="justify",
)
bullet(
    "**La limpieza es parte del contrato, no un detalle.** Un `useEffect` que abre un temporizador o una petición "
    "debe cerrarlo en su función de limpieza. Esa función corre al repetirse y al desmontar; si se coloca en un "
    "segundo efecto con `[]`, los recursos viejos sobreviven y aparecen valores intermedios o respuestas tardías."
)
bullet(
    "**Las carreras de peticiones se resuelven con una bandera por corrida.** No basta con cancelar el temporizador: "
    "una respuesta vieja puede llegar después de la nueva. Marcar cada corrida como cancelada y comprobarlo antes de "
    "escribir el estado es simple y cubre `reload`, `StrictMode` y el desmontaje."
)
bullet(
    "**Una máquina de estados explícita evita los atajos peligrosos.** El estado `awaiting_approval` existe para que "
    "un plan destructivo no pueda pasar de `running` a `done` sin una decisión humana. La prueba lo verifica; el "
    "código lo hace evidente."
)
bullet(
    "**La IA acelera, pero no decide.** Los seis fallos de `src/ai` son plausibles y fáciles de pasar por alto: "
    "código que compila, se lee bien y falla en los bordes. La revisión humana y las pruebas fueron las que "
    "determinaron qué se aceptaba. Las cuatro preguntas del contrato son una lista de control eficaz."
)
bullet(
    "**El estado se eleva solo hasta donde hace falta.** `App` es el ancestro común de la cabecera y la bandeja; "
    "llevar las decisiones más arriba (a `main.jsx`) o más abajo (a `Header`) solo agrega duplicación o "
    "desincronización."
)
para(
    "**Limitaciones y siguiente paso.** La red es simulada y las decisiones viven en el navegador; el servidor real y "
    "la autenticación llegan en la semana 12. Cuando eso ocurra, `usePlans` y `useAgentRun` podrán recibir funciones "
    "que usen `fetch` con `AbortSignal` real: la estructura de estados y la limpieza ya están preparadas para ello.",
    align="justify",
)

# --------------------------------------------------------------------------- #
# ANEXO A
# --------------------------------------------------------------------------- #
page_break()
h1("Anexo A. Código corregido de los hooks de la IA")
para("Los tres archivos de `src/ai/` quedaron así después de la revisión. Los comentarios `FIX:` señalan el cambio.",
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
h1("Anexo B. Catálogo de las 57 pruebas")
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
    ["Ruta", "Qué contiene"],
    [
        ["`src/hooks/useDebounce.js`", "Hook propio: valor tranquilo con limpieza del temporizador"],
        ["`src/hooks/usePlans.js`", "Hook propio: tres estados, reload y respuesta vieja"],
        ["`src/hooks/useAgentRun.js`", "Hook propio: máquina de estados y barrera R2"],
        ["`src/ai/*.js`", "Hooks de la IA corregidos (6 fallos)"],
        ["`docs/capturas/`", "16 capturas reales (app, terminal y VS Code)"],
        ["`docs/logs/`", "Salidas completas de pruebas y build"],
        ["`docs/scripts/capturar_app.py`", "Guion de capturas de la aplicación (Playwright + Chrome)"],
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
npm run build     # compilación de producción""",
)
para("**Reproducir las capturas de la aplicación:**", space_after=2)
code(
    """npm run dev       # deja el servidor en http://localhost:5173
python docs/scripts/capturar_app.py""",
)

para(
    "Fin del documento. Estudiante: Edu Edward Galindo Huamán · DAE Semana 11 (extra) · 2026.",
    size=9, italic=True, color=GRAY, align="center", space_before=12,
)

# --------------------------------------------------------------------------- #
# Guardar
# --------------------------------------------------------------------------- #
doc.save(OUT)
print("DOCX generado:", OUT)
print("Figuras:", FIG["n"])
print("Pruebas en el catálogo:", n)

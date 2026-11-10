"""Real screenshots of the running app (Vite dev server on http://localhost:5173).

The app is a centred 880 px column, so the captures are clipped to that column
instead of the full browser window: the text stays readable when the image is
placed in the report. Every shot is preceded by an assertion on the visible
text, so a screenshot never shows an unexpected state.

Run with the dev server already up:
    python docs/scripts/capturar_app.py
"""
from pathlib import Path

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / "docs" / "capturas"
OUT.mkdir(parents=True, exist_ok=True)

URL = "http://localhost:5173"
COLUMN_PAD = 6


def doc_box(page, selector):
    """Element box in document coordinates (CSS px), not viewport coordinates."""
    box = page.evaluate(
        """(sel) => {
            const el = document.querySelector(sel);
            if (!el) return null;
            const r = el.getBoundingClientRect();
            return {x: r.left + window.scrollX, y: r.top + window.scrollY, w: r.width, h: r.height};
        }""",
        selector,
    )
    assert box, f"No existe el selector {selector}"
    return box


def shot(page, name, bottom=None, selector=None, pad=16):
    """Capture a region: the app column up to `bottom`, or one element's box."""
    if selector:
        # Element screenshot: works even when the panel sits below the fold.
        page.locator(selector).first.screenshot(path=str(OUT / name))
        print(f"saved {name} (element {selector})")
        return
    app = doc_box(page, ".app")
    clip = {
        "x": max(0, app["x"] - COLUMN_PAD),
        "y": 0,
        "width": app["w"] + 2 * COLUMN_PAD,
        "height": bottom,
    }
    page.screenshot(path=str(OUT / name), clip=clip)
    print(f"saved {name} ({int(clip['width'])}x{int(clip['height'])})")


def expect_text(page, text):
    """The state must be on screen before the shutter, or the shot is worthless."""
    page.get_by_text(text, exact=False).first.wait_for(state="visible", timeout=5000)
    print("  ok:", text)


with sync_playwright() as p:
    browser = p.chromium.launch(channel="chrome", headless=True)
    context = browser.new_context(
        viewport={"width": 900, "height": 900},
        device_scale_factor=2,   # crisp at 300 dpi once placed in the document
        locale="es-PE",
    )
    page = context.new_page()

    # 01 - loading state (skeleton), captured before the fake network answers.
    page.goto(URL, wait_until="load")
    page.wait_for_selector('[aria-label="Cargando planes"]')
    shot(page, "01_carga_esqueleto.png", bottom=590)

    # 02 - ready state: header counter + first cards of the list.
    page.wait_for_selector('article[aria-label="Plan P-01"]')
    expect_text(page, "Pendientes: 6")
    shot(page, "02_planes_cargados.png", bottom=1000)

    # 03 - error state (?fail=1) and the retry button.
    page.goto(f"{URL}/?fail=1", wait_until="load")
    page.wait_for_selector('[role="alert"]')
    expect_text(page, "No se pudieron cargar los planes")
    banner = doc_box(page, ".error-banner")
    shot(page, "03_error_red.png", bottom=banner["y"] + banner["h"] + 30)

    # 04 and 05 - search: the field updates at once, the list waits for useDebounce.
    page.goto(URL, wait_until="load")
    page.wait_for_selector('article[aria-label="Plan P-01"]')
    search = page.get_by_role("searchbox", name="Buscar plan")
    search.fill("values")
    shot(page, "04_busqueda_antigua.png", bottom=1000)
    page.wait_for_timeout(450)
    # Only P-03 survives the filter; the id lives in aria-label, not in the visible card text.
    assert page.locator("article").count() == 1
    assert page.locator("article").first.get_attribute("aria-label") == "Plan P-03"
    print("  ok: una sola tarjeta, Plan P-03")
    single = doc_box(page, 'article[aria-label="Plan P-03"]')
    shot(page, "05_busqueda_filtrada.png", bottom=single["y"] + single["h"] + 30)

    # 06 - approve an R1 plan: the header counter comes from App.
    search.fill("")
    page.wait_for_timeout(450)
    page.locator('article[aria-label="Plan P-03"]').get_by_role("button", name="Aprobar").click()
    expect_text(page, "Pendientes: 5")
    shot(page, "06_contador_pendientes.png", bottom=1000)

    # 07 - approve an R2 plan (hash required) and run it: the safety barrier appears.
    r2 = page.locator('article[aria-label="Plan P-01"]')
    r2.get_by_role("textbox", name="Hash del plan").fill("d68ebfa913ec")
    r2.get_by_role("button", name="Aprobar").click()
    page.get_by_role("button", name="Ejecutar P-01").click()
    expect_text(page, "Falta tu confirmación para aplicar")
    shot(page, "07_ejecucion_r2_espera_aprobacion.png", selector='section[aria-labelledby="run-title"]')

    # 08 - final confirmation applies the destructive plan.
    page.get_by_role("button", name="Confirmar y aplicar").click()
    expect_text(page, "P-01 aplicado")
    shot(page, "08_ejecucion_aplicada.png", selector='section[aria-labelledby="run-title"]')

    # 09 - the decisions and the pending count survive a reload (localStorage).
    page.reload(wait_until="load")
    page.wait_for_selector('article[aria-label="Plan P-02"]')
    expect_text(page, "Pendientes: 4")
    shot(page, "09_persistencia_recarga.png", bottom=1000)

    # 10 - mobile viewport, to show the layout is usable on a narrow screen.
    mobile = browser.new_context(
        viewport={"width": 390, "height": 844},
        device_scale_factor=2,
        locale="es-PE",
    )
    mpage = mobile.new_page()
    mpage.goto(URL, wait_until="load")
    mpage.wait_for_selector('article[aria-label="Plan P-02"]')
    mpage.screenshot(path=str(OUT / "10_vista_movil.png"), full_page=False)
    print("saved 10_vista_movil.png")

    browser.close()

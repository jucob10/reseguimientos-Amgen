# -*- coding: utf-8 -*-
"""
Convierte manual.html → Manual_de_usuario_Reseguimientos.pdf (raíz del proyecto).
Requiere: pip install playwright  (usa el Chrome instalado en el equipo)

    python docs/manual/generar_pdf.py
"""
from pathlib import Path
from playwright.sync_api import sync_playwright

AQUI = Path(__file__).parent
SALIDA = AQUI.parents[1] / "Manual_de_usuario_Reseguimientos.pdf"

PIE = """
<div style="width:100%; font-size:8px; color:#5F6B7A; padding:0 18mm; display:flex; justify-content:space-between;
            font-family:Segoe UI, Arial, sans-serif;">
  <span>Reseguimientos Amgen · Manual de usuario v1.0</span>
  <span>Página <span class="pageNumber"></span> de <span class="totalPages"></span></span>
</div>"""

with sync_playwright() as p:
    navegador = p.chromium.launch(channel="chrome")
    pagina = navegador.new_page()
    pagina.goto((AQUI / "manual.html").as_uri(), wait_until="networkidle")
    pagina.pdf(path=str(SALIDA), format="A4", print_background=True, prefer_css_page_size=True,
               display_header_footer=True, header_template="<span></span>", footer_template=PIE)
    navegador.close()
print(f"PDF generado: {SALIDA}")

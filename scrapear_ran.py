#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extrae la lista de sujetos agrarios del RAN (modulo informativo
"Lista de Ejidatarios, Comuneros, Posesionarios y Avecindados") a un Excel.

Sitio: https://consultasimcr.ran.gob.mx/consultassujetoagrario.aspx

Instalacion (una sola vez):
    pip install playwright pandas openpyxl
    python -m playwright install chromium

Uso:
    python scrapear_ran.py                      # busqueda por defecto (YUCATAN/ACANCEH)
    python scrapear_ran.py --headless           # sin ventana (uso normal)
    python scrapear_ran.py --estado "YUCATÁN" --municipio ACANCEH \
        --tipo EJIDO --nucleo ACANCEH --salida ejidatarios_acanceh.xlsx
"""
from __future__ import annotations

import argparse
import sys
import time

import pandas as pd
from playwright.sync_api import sync_playwright

# --------------------------------------------------------------------------
# Parametros por defecto (se pueden sobreescribir con argumentos de linea
# de comandos; ver --help).
# --------------------------------------------------------------------------
URL = "https://consultasimcr.ran.gob.mx/consultassujetoagrario.aspx"
ESTADO = "YUCATÁN"
MUNICIPIO = "ACANCEH"
TIPO = "EJIDO"
NUCLEO = "ACANCEH"
SALIDA = "ejidatarios_acanceh.xlsx"

PAUSA = 1.5          # segundos entre paginas (no saturar el servidor)
MAX_PAGINAS = 2000   # tope de seguridad
HEADLESS = True      # True = sin ventana; usa --visible para ver el navegador

# IDs reales del formulario GeneXus (verificados contra el sitio)
SEL_ESTADO = "#vWEDO_ID"
SEL_MUNICIPIO = "#vWMUN_ID"
SEL_TIPO = "#vWNA_TIPO"
SEL_NUCLEO = "#vWNA_ID"
BTN_BUSCAR = "#BUTTON1"
TABLA = "#Grid1ContainerTbl"
BTN_SIGUIENTE = "button.PagingButtonsNext"

# Lee encabezados (thead) y filas de datos (tbody) de la tabla de resultados.
JS_LEER = r"""
() => {
    const t = document.querySelector('#Grid1ContainerTbl');
    if (!t) return { ths: [], filas: [] };
    const ths = [...t.querySelectorAll('thead th')]
        .map(th => th.innerText.trim().replace(/\s+/g, ' '));
    const filas = [...t.querySelectorAll('tbody > tr')]
        .map(tr => [...tr.children]
            .map(td => td.innerText.trim().replace(/\s+/g, ' ')));
    return { ths, filas };
}
"""

# Estado interno de GeneXus. OJO: en este sitio GRID1_nEOF se queda en "1"
# durante toda la paginacion y GRID1_nFirstRecordOnPage no se refresca en el
# DOM, asi que NO sirven para detectar el final. La senal fiable es que el
# boton "Siguiente" queda disabled en la ultima pagina (ver main()).
JS_ESTADO = r"""
() => {
    try {
        const s = JSON.parse(document.getElementById('GXState').value);
        return { nEOF: s.GRID1_nEOF };
    } catch (e) { return {}; }
}
"""


def elegir(page, selector: str, texto: str, campo: str) -> None:
    """Selecciona `texto` en el combo `selector`, esperando a que exista."""
    try:
        page.wait_for_function(
            """([sel, t]) => {
                const s = document.querySelector(sel);
                return s && [...s.options].some(o => o.text.trim() === t);
            }""",
            arg=[selector, texto],
            timeout=30000,
        )
    except Exception:
        opciones = page.eval_on_selector(
            selector, "s => [...s.options].map(o => o.text.trim()).join(', ')"
        )
        raise SystemExit(
            f"No existe la opcion {texto!r} en {campo}.\nOpciones: {opciones}"
        )
    page.select_option(selector, label=texto)
    page.wait_for_timeout(1500)  # deja que el sitio recargue el combo siguiente


def firma(filas) -> tuple:
    return tuple(tuple(f) for f in filas)


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Extrae sujetos agrarios del RAN a Excel."
    )
    ap.add_argument("--estado", default=ESTADO)
    ap.add_argument("--municipio", default=MUNICIPIO)
    ap.add_argument("--tipo", default=TIPO, help="EJIDO o COMUNIDAD")
    ap.add_argument("--nucleo", default=NUCLEO)
    ap.add_argument("--salida", default=SALIDA, help="archivo .xlsx de salida")
    ap.add_argument("--pausa", type=float, default=PAUSA,
                    help="segundos entre paginas (default 1.5)")
    ap.add_argument("--max-paginas", type=int, default=MAX_PAGINAS)
    ap.add_argument("--visible", action="store_true",
                    help="muestra el navegador (equivale a headless=False)")
    ap.add_argument("--headless", action="store_true",
                    help="sin ventana (default)")
    args = ap.parse_args()
    headless = not args.visible or args.headless

    print(f"Busqueda: {args.estado} / {args.municipio} / {args.tipo} / {args.nucleo}")
    print(f"Navegador: {'sin ventana' if headless else 'VISIBLE'}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        page = browser.new_page()
        page.goto(URL, wait_until="domcontentloaded")
        page.wait_for_timeout(3000)

        elegir(page, SEL_ESTADO, args.estado, "Estado")
        elegir(page, SEL_MUNICIPIO, args.municipio, "Municipio")
        elegir(page, SEL_TIPO, args.tipo, "Tipo de nucleo agrario")
        elegir(page, SEL_NUCLEO, args.nucleo, "Nucleo agrario")

        page.click(BTN_BUSCAR)

        # Esperar la primera pagina de resultados
        page.wait_for_selector(f"{TABLA} tbody > tr", timeout=60000)
        page.wait_for_timeout(2000)

        encabezado = None
        todas: list[list[str]] = []
        vistas: set[tuple] = set()
        pagina = 1
        motivo_fin = ""

        while True:
            datos = page.evaluate(JS_LEER)
            filas = datos["filas"]
            if datos["ths"]:
                encabezado = datos["ths"]

            if not filas:
                motivo_fin = "la tabla quedo vacia"
                break

            sig = firma(filas)
            if sig in vistas:
                motivo_fin = "el contenido se repitio (no hay mas paginas)"
                break
            vistas.add(sig)

            todas.extend(filas)
            print(f"Pagina {pagina}: {len(filas)} filas (acumulado {len(todas)})")

            if pagina >= args.max_paginas:
                motivo_fin = f"se alcanzo el tope de {args.max_paginas} paginas"
                break

            # Ultima pagina: GeneXus deshabilita "Siguiente" y "Ultimo".
            siguiente = page.locator(BTN_SIGUIENTE).first
            if siguiente.count() == 0 or not siguiente.is_visible() \
                    or siguiente.is_disabled():
                motivo_fin = "el boton 'Siguiente' quedo deshabilitado (ultima pagina)"
                break

            siguiente.click()

            # Esperar a que la tabla cambie de verdad (contenido distinto).
            cambio = False
            for _ in range(60):  # hasta ~30 s
                page.wait_for_timeout(500)
                nuevo = page.evaluate(JS_LEER)["filas"]
                if nuevo and firma(nuevo) != sig:
                    cambio = True
                    break
            if not cambio:
                motivo_fin = "la tabla no cambio tras pulsar 'Siguiente'"
                break

            time.sleep(args.pausa)
            pagina += 1

        # Guardar HTML de la ultima pagina por si hay que revisar algo
        with open("debug_ultima_pagina.html", "w", encoding="utf-8") as f:
            f.write(page.content())
        browser.close()

    if not todas:
        raise SystemExit("No se extrajo ningun registro.")

    # Quitar encabezados repetidos que pudieran venir como fila de datos.
    # (Los encabezados reales estan en <thead>, esto es solo una red de seguridad.)
    if encabezado:
        datos = [f for f in todas if [c.strip().upper() for c in f]
                 != [c.strip().upper() for c in encabezado]]
    else:
        datos = todas
        encabezado = [f"COLUMNA {i + 1}" for i in range(len(todas[0]))]

    # Normalizar filas que no tengan el ancho del encabezado
    ancho = len(encabezado)
    normalizadas = []
    for f in datos:
        if len(f) < ancho:
            f = f + [""] * (ancho - len(f))
        elif len(f) > ancho:
            f = f[:ancho]
        normalizadas.append(f)

    # No se eliminan duplicados: puede haber homonimos legitimos.
    df = pd.DataFrame(normalizadas, columns=encabezado)
    df.to_excel(args.salida, index=False)

    print(f"\nFin: {motivo_fin}")
    print(f"Paginas leidas: {pagina}")
    print(f"Registros en el Excel: {len(df)}  ->  {args.salida}")
    print("(Los duplicados NO se eliminaron: puede haber homonimos legitimos.)")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()

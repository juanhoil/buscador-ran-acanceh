#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convierte el Excel del RAN en el archivo de datos que consume index.html.

No modifica el Excel de entrada ni la estructura de los datos: conserva las
9 columnas originales con sus nombres exactos. Solo cambia el *formato de
empaquetado* (columnar en vez de fila-a-fila) porque ocupa 2.4x menos y se
parsea igual de rapido.

Formato de salida (datos.js):
    window.DATOS_RAN = {
      version, generado, fuente, consulta,
      columnas: [...],          # nombres originales, en su orden original
      filas:    [[...], ...]    # mismas celdas, mismos valores
    };

Se usa .js y no .json a proposito: asi se puede cargar con <script src> y la
pagina funciona tambien abriendo el archivo con doble clic (file://), donde
fetch() de un .json local lo bloquea CORS.

Uso:
    python generar_datos.py
    python generar_datos.py --entrada otro.xlsx --salida datos.js
    python generar_datos.py --json          # escribe ademas datos.json
"""
from __future__ import annotations

import argparse
import hashlib
import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

ENTRADA = "ejidatarios_acanceh.xlsx"
SALIDA = "datos.js"
FUENTE = "https://consultasimcr.ran.gob.mx/consultassujetoagrario.aspx"


def leer_excel(ruta: Path) -> pd.DataFrame:
    if not ruta.exists():
        raise SystemExit(f"No existe el archivo de entrada: {ruta}")
    # dtype=str + keep_default_na=False -> los vacios quedan como "" y no como
    # NaN, para no convertir "1801" en 1801 ni perder ceros a la izquierda.
    df = pd.read_excel(ruta, dtype=str, keep_default_na=False)
    df.columns = [str(c).strip() for c in df.columns]
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    return df


def metadatos_consulta(df: pd.DataFrame) -> dict:
    """Deduce la consulta a partir de las columnas constantes del Excel."""
    def unico(col: str) -> str:
        if col not in df.columns:
            return ""
        vals = df[col].unique()
        return vals[0] if len(vals) == 1 else "(varios)"

    return {
        "edo_id": unico("EDO_ID"),
        "municipio": unico("MUNICIPIO"),
        "tipo": unico("NA_TIPO"),
        "nucleo": unico("NUCLEO AGRARIO"),
    }


def main() -> None:
    ap = argparse.ArgumentParser(
        description="Convierte el Excel del RAN en datos.js para el buscador web."
    )
    ap.add_argument("--entrada", default=ENTRADA, help="Excel de entrada")
    ap.add_argument("--salida", default=SALIDA, help="archivo de datos de salida")
    ap.add_argument("--json", action="store_true",
                    help="escribe tambien un .json con el mismo contenido")
    args = ap.parse_args()

    ruta = Path(args.entrada)
    df = leer_excel(ruta)

    columnas = list(df.columns)
    filas = df.values.tolist()

    # El hash del contenido es la "version": el navegador la compara con la
    # que tiene en cache y solo se actualiza si realmente cambio algo.
    crudo = json.dumps({"columnas": columnas, "filas": filas},
                       ensure_ascii=False, separators=(",", ":"))
    version = hashlib.sha256(crudo.encode("utf-8")).hexdigest()[:12]

    datos = {
        "version": version,
        "generado": datetime.now().isoformat(timespec="seconds"),
        "fuente": FUENTE,
        "consulta": metadatos_consulta(df),
        "columnas": columnas,
        "filas": filas,
    }

    cuerpo = json.dumps(datos, ensure_ascii=False, separators=(",", ":"))
    salida = Path(args.salida)
    salida.write_text(
        "/* Generado por generar_datos.py - no editar a mano. */\n"
        f"window.DATOS_RAN = {cuerpo};\n",
        encoding="utf-8",
    )

    if args.json:
        salida.with_suffix(".json").write_text(cuerpo, encoding="utf-8")

    kb = salida.stat().st_size / 1024
    print(f"Entrada : {ruta}  ({len(df)} filas x {len(columnas)} columnas)")
    print(f"Salida  : {salida}  ({kb:.1f} KB)")
    print(f"Version : {version}")
    print(f"Consulta: {datos['consulta']}")

    constantes = [c for c in columnas if df[c].nunique() == 1]
    if constantes:
        print(f"Aviso   : columnas constantes conservadas tal cual: {constantes}")


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    main()

# Buscador de sujetos agrarios del RAN

Página web ligera para consultar la lista de sujetos agrarios del RAN
(Ejidatarios, Comuneros, Posesionarios y Avecindados) extraída de
[consultasimcr.ran.gob.mx](https://consultasimcr.ran.gob.mx/consultassujetoagrario.aspx).

Trabaja sin conexión, busca por nombre o apellido, edita un resaltado
mientras tipeas y permite ordenar los resultados por nombre o apellido
paterno. Todo en un solo archivo HTML estático, sin servidor.

---

## Cómo se usan los datos

| Archivo | Para qué sirve |
| --- | --- |
| `index.html` | La página. Abrela directamente en el navegador. |
| `datos.js` | Los datos (formato columnar), generados por `generar_datos.py`. |
| `ejidatarios_acanceh.xlsx` | Excel original descargado por `scrapear_ran.py`. |
| `scrapear_ran.py` | El scraper: descarga la lista del sitio y la guarda en el Excel. |
| `generar_datos.py` | Convierte el Excel en `datos.js`. |

---

## Flujo de trabajo

```powershell
# 1) Descargar datos del sitio del RAN
python scrapear_ran.py --municipio ACANCEH --nucleo ACANCEH --salida ejidatarios_acanceh.xlsx

# 2) Convertir el Excel a datos.js (lo que consume la página web)
python generar_datos.py

# 3) Abrir index.html con doble clic, o servirlo localmente:
python -m http.server 8080
# luego http://localhost:8080
```

Si solo necesitas la página y los datos ya existen, abrela con doble clic y listo.

---

## Estrategia técnica

- **Almacenamiento** — El Excel se convierte una sola vez a `datos.js` en
  formato columnar (`{columnas: [...], filas: [[...], ...]}`). Ocupa 2.4x
  menos que un array de objetos (211 KB vs 507 KB) y se parsea igual de
  rápido. Se carga con `<script src>` y no con `fetch()` para que la página
  funcione también abriendo el archivo con doble clic (file://), donde
  fetch() local está bloqueado por CORS.

- **Caché** — IndexedDB guarda el dataset completo. Al abrir, primero se
  pinta lo cacheado (instantáneo y sin conexión) y después se revalida
  contra `datos.js` comparando el hash de contenido; solo se repinta si
  cambió de verdad.

- **Búsqueda** — Claves normalizadas (minúsculas, sin acentos) precalculadas.
  Consulta multi-término con AND: `"canul virginio"` encuentra
  "VIRGINIO CANUL EK" sin importar el orden. Sin índices invertidos: con
  2608 filas el escaneo lineal tarda <1 ms.

- **Orden** — 4 arreglos de índices precalculados (2 campos × 2 sentidos).
  Filtrar nunca ordena, así que el coste por tecla es O(n) y no O(n log n).

- **Render** — Lista virtualizada: solo se pintan las filas visibles en cada
  momento (~40 en pantalla, no las 3000).

---

## Parámetros

`scrapear_ran.py` (descargar datos):

| Argumento | Default |
| --- | --- |
| `--estado` | YUCATÁN |
| `--municipio` | ACANCEH |
| `--tipo` | EJIDO |
| `--nucleo` | ACANCEH |
| `--salida` | ejidatarios_acanceh.xlsx |
| `--pausa` | 1.5 (segundos entre páginas) |
| `--max-paginas` | 2000 |
| `--visible` / `--headless` | headless |

`generar_datos.py` (convertir Excel a JS):

| Argumento | Default |
| --- | --- |
| `--entrada` | ejidatarios_acanceh.xlsx |
| `--salida` | datos.js |
| `--json` | (opcional) escribe también un `datos.json` |

---

## Atajos en la página

- `/` enfoca el buscador
- `Esc` limpia la búsqueda o cierra el detalle
- Clic en una fila abre el detalle de esa persona

---

## Publicación (GitHub Pages)

La página está publicada en **GitHub Pages**: basta con abrir la URL pública para
usarla desde cualquier navegador, sin clonar ni ejecutar nada local.

Flujo de actualización:

```powershell
# 1) Regenerar el Excel si quieres datos nuevos
python scrapear_ran.py --municipio MOTUL --nucleo MOTUL --salida ejidatarios_motul.xlsx

# 2) Regenerar datos.js
python generar_datos.py --entrada ejidatarios_motul.xlsx --salida datos.js

# 3) Subir cambios a GitHub (Pages se actualiza solo en ~1 min)
git add .
git commit -m "Actualizar datos"
git push
```

GitHub Pages sirve los archivos estáticos desde la rama `main`, así que cada
`push` redespliega la página automáticamente. La URL exacta aparece en la
configuración del repositorio en GitHub (Settings → Pages).
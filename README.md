# Generador de reportes (Excel)

Este proyecto contiene scripts en Python para construir hojas de reporte dentro de un archivo Excel de salida (`salida.xlsx`) a partir de un archivo de entrada (`input.xlsx`). Se incluyen módulos para comentarios, consolidado, índice de clima, por pregunta y variables demográficas.

## Requisitos

- Python 3.10+ con paquetes:
  - pandas
  - openpyxl
- Archivos de trabajo en esta carpeta:
  - `input.xlsx` (origen de datos)
  - `salida.xlsx` (destino de reportes; se crea/actualiza)
  - `modelo output.xlsx` (plantilla opcional para “Por pregunta”)
  - `variables_demograficas_mapping.json` (mapeo de variables demográficas)
  - `indice_clima_mapping.json` (mapeo del índice de clima)

Notas:
- Si Excel/OneDrive mantiene un archivo abierto, los scripts escriben una copia temporal en `%LocalAppData%\Temp` y muestran su ubicación en consola.
- Todos los comandos están pensados para ejecutarse en Windows PowerShell.

## Scripts disponibles y opciones principales

### 1) Variables demográficas (nuevo)

Genera la hoja `V. Demogr (nuevo)` y por defecto deja solo esa hoja en `salida.xlsx`.

Archivo: `generar_variables_demograficas_nueva.py`

- Entradas por defecto: `input.xlsx`, hoja `Datos A2`.
- Salida: `salida.xlsx` → hoja `V. Demogr (nuevo)`.
- Usa `variables_demograficas_mapping.json` para saber qué columnas leer (por letras) y el orden/normalización de categorías.
- Columnas conectadas (ejemplo actual del mapping):
  - Generación: `GU`
  - Tipo NPS: `GV`
  - Antigüedad: `CI`
  - Tipo contrato: `CJ`
  - VARIABLE1: `CM`
  - VARIABLE2: `CL`
  - VARIABLE3: `GS`

Opciones:
- `--in-file` archivo de entrada (por defecto `input.xlsx`).
- `--out salida.xlsx` cambia el nombre del archivo de salida.
- `--sheet "Datos A2"` cambia la hoja de entrada.
- `--id 123 456` filtra por ID(s) (columna A).
- `--sheet-out "V. Demogr (nuevo)"` cambia el nombre de la hoja de salida.
- `--label-mode raw|no-prefix|prefixed` controla cómo mostrar la columna “Variable”:
  - `raw`: texto tal cual en la columna origen (recomendado para ver exactamente lo que viene del Excel).
  - `no-prefix`: remueve prefijos numéricos del orden del JSON.
  - `prefixed`: usa las etiquetas con prefijo según el orden del JSON.
- `--keep-others` conserva otras hojas en `salida.xlsx` (por defecto se eliminan todas menos la nueva).
- `--remove-sheets "V. Demogr"` elimina hojas específicas tras escribir la nueva (opcional).
- `--mapping-file` archivo JSON alterno de mapeo.

Ejemplo (recomendado):
```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\generar_variables_demograficas_nueva.py" \
  --out salida.xlsx \
  --sheet-out "V. Demogr (nuevo)" \
  --label-mode raw
```

### 2) Variables demográficas (original)

Archivo: `generar_variables_demograficas.py`

- Genera `V. Demogr` con conteos, %, y desglose NPS por categoría.
- Usa también `variables_demograficas_mapping.json`.

Ejemplo:
```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\generar_variables_demograficas.py" \
  --out salida.xlsx
```

### 3) Por pregunta

Archivo: `generar_por_pregunta.py`

- Genera la hoja `Por pregunta` con promedios en escala 10 (promedio en 1–4 → redondeo 1 dec → conversión a 10 → redondeo 1 dec).
- Acepta plantilla opcional `modelo output.xlsx` para respetar orden/columnas auxiliares (Clima/IFE/PESO).

Opciones útiles:
- `--sheet "Datos A2"` hoja origen.
- `--id ...` filtra por ID(s).
- `--gen-col GU` columna de Generación (por defecto GU).
- `--nps-col GV` columna de tipo NPS (por defecto GV).
- `--template-file "...\modelo output.xlsx"` usa la plantilla.
- `--debug` imprime mapeo y cálculos.
- `--debug-summary` imprime un resumen por pregunta.

Ejemplo:
```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\generar_por_pregunta.py" \
  --debug --debug-summary
```

### 4) Índice de clima

Archivo: `generar_indice_clima.py`

- Genera la hoja `Indice de clima` con columnas: Eje, Real, Escala 4, Escala 10.
- Usa `indice_clima_mapping.json` para definir la jerarquía y las columnas (por letras) que se promedian.

Opciones:
- `--sheet` hoja origen.
- `--id` filtrar por ID(s).
- `--out` archivo destino.

Ejemplo:
```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\generar_indice_clima.py" \
  --out salida.xlsx
```

### 5) Comentarios

Archivo: `crear_hoja_comentarios.py`

- Crea la hoja `comentarios` extrayendo columnas por letras fijas (ID, Variables 1–3, NPS, Generación, Antigüedad, Tipo contrato, Comentarios).
- Soporta `--out` para escribir en el archivo que indiques.

Ejemplo:
```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\crear_hoja_comentarios.py" \
  --out salida.xlsx
```

### 6) Consolidado

Archivo: `generar_consolidado.py`

- Genera la hoja `consolidado` con métricas base (universo, respuestas, % participación), cortes por generaciones, NPS y otros bloques de preguntas convertidas a escala 10.

## Convenciones clave

- Dirección por letra de columna (A, IB, AL, …) para ubicar variables en el `input.xlsx`.
- Generación (por defecto `GU`) y NPS (por defecto `GV`).
- Conversión de escala: promedio en 1–4 con redondeo a 1 decimal; luego conversión a 10 y redondeo a 1 decimal.
- Manejo de bloqueos: si el archivo está en uso, se escribe una copia temporal y se informa la ruta.

## Personalización del mapeo

- `variables_demograficas_mapping.json`:
  - Define grupos, columna (por letra), orden (con o sin prefijos), y reglas `normalize` para consolidar variantes.
  - Puedes añadir grupos como `VARIABLE2`/`VARIABLE3` y asignarles columnas (por ejemplo `CL`, `GS`).
  - El generador “nuevo” permite elegir cómo mostrar la etiqueta (“Variable”) mediante `--label-mode`.

- `indice_clima_mapping.json`:
  - Define estructura jerárquica de filas y las columnas que promedian.
  - Para filas “group” con `columns` se calcula el promedio del conjunto especificado.

## Consejos y resolución de problemas

- “Me duplica categorías (con y sin prefijo)”: usa el generador “nuevo” con `--label-mode raw` para ver exactamente el texto original de tus columnas; ya se consolidan internamente las variantes.
- “No encuentro la hoja de origen”: usa `--sheet` o deja que el script haga fuzzy match por nombre.
- “El archivo quedó con varias hojas y solo quiero una”: el generador “nuevo” ya elimina todas las demás por defecto; si quieres conservarlas, usa `--keep-others`.
- “Archivo de salida bloqueado”: cierra `salida.xlsx` y vuelve a ejecutar; si se generó una copia temporal, ábrela y revisa resultados.

## Ejecución típica (solo dejar la hoja demográfica nueva)

```powershell
python \
  "c:\Users\anali\OneDrive\Documentos\Proyectos\Clima WT\Generador documento de analisis\generar_variables_demograficas_nueva.py" \
  --out salida.xlsx \
  --sheet-out "V. Demogr (nuevo)" \
  --label-mode raw
```

Si necesitas que prepare otro reporte además del demográfico, ejecuta su script correspondiente después (o indícame y te creo un pequeño orquestador).

---

¿Quieres que dejemos `V. Demogr (nuevo)` renombrado como `V. Demogr` tras escribir y eliminar las demás hojas? Se puede ajustar en el script si lo prefieres.

from pathlib import Path
import argparse
import sys
import os
import shutil
import tempfile
import pandas as pd
from difflib import SequenceMatcher
import unicodedata
import time

INPUT_FILE = 'input.xlsx'
OUTPUT_FILE = 'salida.xlsx'
DEFAULT_SHEET = 'Datos A2'
SHEET_OUT = 'consolidado'

# Orden de columnas solicitado: Universo, Respuestas, % Participación
# Agregamos métricas por cohortes con variantes: % Centenialls, % Milenialls, % Generación X, % Baby Boomers
CONSOLIDADO_COLS = [
    'Total Personas', 'Respuestas', '% Participación',
    '% Centenialls', '% Milenialls', '% Generación X', '% Baby Boomers',
    'Satisfacción General',
    'Satisfacción Centenialls',
    'Satisfacción Milenialls',
    'Satisfacción Generación X',
    'Satisfacción Baby Boomers',
    'COVID',
    'Condiciones de trabajo',
    'Indice de clima',
    'NPS',
    'NPS\nCentenialls',
    'NPS\nMilenialls',
    'NPS\nGeneración X',
    'NPS\nBaby Boomers',
    'NPS Entusiastas',
    'NPS Pasivos',
    'NPS Detractores',
    '¿Cómo es trabajar en la organización?',
    'Me siento orgulloso(a) cuando le cuento a otros que estoy trabajando en esta Organización.',
    'Recomendaría a otros trabajar en esta Organización.',
    'Me parece inspiradora la misión de la Organización y estoy comprometido (a) con ella.',
    'Comprendo los objetivos de la Organzación y sé como puedo aportar a conseguirlos desde mi cargo.',
    'Se reciben beneficios no monetarios que hacen más agradable la experiencia en la Organización.',
    'Se realizan actividades de bienestar que brindan espacios de esparcimiento para el empleado y su familia.',
    'Las características de mi trabajo me permiten tener un adecuado balance entre mi trabajo y mi vida personal.',
    'La empresa demuestra sensibilidad con las particularidades de la vida personal de sus empleados.',
    '¿Cómo lo soporta para lograr los resultados de su cargo?',
    'Cuento con los recursos mínimos que requiero para realizar de forma adecuada mi trabajo.',
    'Puedo acceder a toda la información que necesito para hacer bien mi trabajo.',
    'Siento que mi cargo me brinda la oportunidad de progresar y desarrollarme.',
    'Todos en la organización tienen claro su cargo y comprenden bien el alcance sus responsabilidades.',
    'La formación que brinda la Organización realmente me ayuda a realizar mejor mi trabajo.',
    'El entrenamiento que recibo en mi cargo me ayuda progresar a nivel personal y profesional.',
    'Me evalúan por mi desempeño y tengo opciones de recibir reconocimiento si he realizado un esfuerzo por hacer mi trabajo de manera sobresaliente.',
    'En esta Organización celebramos los éxitos y nos felicitamos por los logros obtenidos.',
    '¿Cómo es la dinámica del trabajo conjunto para lograr resultados?',
    'La comunicación entre áreas fluye de manera adecuada.',
    'Mi jefe se comunica de forma clara y verifica que yo haya comprendido lo que me dice.',
    'Conozco personas de diferentes áreas a la mía y se puede trabajar en un entorno agradable, de camaradería y cooperación.',
    'En mi área se facilita y promueve el trabajo en equipo, la integración y la colaboración entre compañeros.',
    'En la Organización las decisiones se toman de forma oportuna y no se aplazan decisiones importantes.',
    'Las decisiones que se toman en la organización se hacen con base en análisis completos y sistemáticos.',
    'Los líderes en la organización inspiran, motivan a los equipos y dan ejemplo con su comportamiento.',
    'Los jefes forman a sus colaboradores y les dan una adecuada y oportuna retroalimentación sobre su desempeño.',
    '¿Cómo es la cultura?',
    'Los jefes y directivos son cordiales y respetuosos en el trato con personas de todos los niveles en la Organización.',
    'Puedo dar mi punto de vista y sugerencias a la organización sabiendo que de ser posible se van a tener en cuenta.',
    'Siento que las relaciones en esta Organización se establecen sobre la base de la confianza en las personas.',
    'Puedo determinar formas propias de hacer mis labores mientras cumpla con las políticas de la Organización.',
    'En esta Organización el trato es justo y equitativo, sin discriminiaciones por género, edad, raza u orientación sexual.',
    'En esta Organización se evita que se utilice la intimidación y el hostigamiento para inducir la renuncia de algún empleado.',
    'Yo quiero a la Organización y me siento comprometido con mi trabajo.',
    'Me gusta mi trabajo y siento que estoy haciendo algo de valor para la Organización.',
    '¿Cómo genera un ambiente sano, limpio y seguro?',
    'Puedo desempeñar mi trabajo de forma segura y cómoda en cuanto a temperatura, sonido, espacio e iluminación.',
    'Puedo manejar el estrés que se genera en mi trabajo para evitar que se afecte mi salud.',
    'Los procesos de la organización cuentan con controles efectivos de calidad.',
    'Se revisa de forma constante la calidad de los procesos y servicios en busca de un mejoramiento continuo.',
    'Se nota una sensibilidad en la Organización por cuidar el impacto que su operación pueda generar en el medio ambiente.',
    'La Organización implementa políticas y/o procedimientos para que los empleados cuidemos el medio ambiente.',
    'La Organización contrata empleados y proveedores cumpliendo con todas los requisitos legales y pagando oportunamente.',
    'La Organización cuida el impacto que puede tener en la comunidad en donde opera y busca aportar a la misma.'
]


def norm_text(s: str) -> str:
    if s is None:
        return ''
    s = str(s).strip()
    s = ' '.join(s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').split())
    return s


def col_letter_to_index(letter: str) -> int:
    if not letter:
        return -1
    s = letter.strip().upper()
    total = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'):
            continue
        total = total * 26 + (ord(ch) - ord('A') + 1)
    return total - 1


def find_sheet(xls: pd.ExcelFile, wanted: str | None) -> str:
    sheets = xls.sheet_names
    if not sheets:
        raise SystemExit('El archivo no contiene hojas.')
    if wanted:
        wn = wanted.strip().lower()
        for s in sheets:
            if s.strip().lower() == wn or s.strip().lower().startswith(wn):
                return s
        # fuzzy
        best = None
        best_score = 0.0
        for s in sheets:
            score = SequenceMatcher(None, wn, s.strip().lower()).ratio()
            if score > best_score:
                best = s
                best_score = score
        if best and best_score >= 0.6:
            return best
        if len(sheets) == 1:
            return sheets[0]
        raise SystemExit(f'No se encontró la hoja "{wanted}". Hojas disponibles: {sheets}')
    return sheets[0] if len(sheets) == 1 else find_sheet(xls, DEFAULT_SHEET)


def normalize_label(s: str) -> str:
    if s is None:
        return ''
    s = str(s)
    # quitar acentos
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    # normalizar espacios y case
    s = ' '.join(s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').split()).strip().lower()
    return s


# Variantes normalizadas para cohortes etarias/generacionales
COHORT_VARIANTS = {
    'centenialls': {
        'centenialls', 'centennials', 'centennial', 'centenials', 'centenial', 'centennials',
        'generacion z', 'generación z', 'gen z', 'z'
    },
    # Nota: Dejamos preparado para futuros pasos, no se usan aún en esta iteración.
    'milenialls': {
        'milenialls', 'millennials', 'millenials', 'millennial', 'millenial',
        'generacion y', 'generación y', 'gen y', 'y'
    },
    'generacion x': {'generacion x', 'generación x', 'gen x', 'x'},
    'baby boomers': {'baby boomers', 'baby boomer', 'babyboomers', 'babyboomer', 'boomers', 'boomer'}
}

 


def open_excel_safely(p: Path) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(p, engine='openpyxl')
    except PermissionError:
        # Intentar copiar con reintentos por bloqueo de OneDrive/Excel
        tmpdir = Path(tempfile.gettempdir())
        last_err = None
        for i in range(3):
            try:
                tmp_path = tmpdir / f"tmp_{p.stem}_{os.getpid()}_{i}.xlsx"
                shutil.copy2(p, tmp_path)
                return pd.ExcelFile(tmp_path, engine='openpyxl')
            except PermissionError as e:
                last_err = e
                time.sleep(1.0)
        # Si no fue posible copiar, relanzar el último error
        raise last_err


def write_excel_safely(out_path: Path, df_outs: dict[str, pd.DataFrame]):
    mode = 'a' if out_path.exists() else 'w'
    writer_kwargs = dict(engine='openpyxl', mode=mode)
    if mode == 'a':
        writer_kwargs['if_sheet_exists'] = 'replace'
    try:
        with pd.ExcelWriter(out_path, **writer_kwargs) as writer:
            for sheet_name, df in df_outs.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        final_path = out_path
    except PermissionError:
        tmpdir = Path(tempfile.gettempdir())
        temp_out = tmpdir / f"tmp_{out_path.stem}_{os.getpid()}.xlsx"
        with pd.ExcelWriter(temp_out, engine='openpyxl', mode='w') as writer:
            for sheet_name, df in df_outs.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        try:
            shutil.copy2(temp_out, out_path)
            final_path = out_path
        except PermissionError:
            final_path = temp_out
            print(f'El archivo de salida estaba en uso. Se creó una copia temporal en: {temp_out}', file=sys.stderr)
    return final_path

def build_row_mask_all_true(n: int) -> pd.Series:
    return pd.Series([True] * n)


def get_aligned_series_by_letter(df: pd.DataFrame, df_raw: pd.DataFrame, letter: str | None) -> pd.Series:
    if not letter:
        return pd.Series([], dtype=object)
    idx = col_letter_to_index(letter)
    if idx < 0 or idx >= df_raw.shape[1]:
        return pd.Series([], dtype=object)
    series = df_raw.iloc[:, idx]
    # si df_raw incluye encabezado y df no, alinear
    if len(series) == len(df) + 1:
        series = series.iloc[1:]
    series = series.reset_index(drop=True)
    return series


def build_id_mask(df: pd.DataFrame, df_raw: pd.DataFrame, ids: list[str] | None) -> pd.Series | None:
    if not ids:
        return None
    idx_id = col_letter_to_index('A')
    if 0 <= idx_id < df_raw.shape[1]:
        series_id = df_raw.iloc[:, idx_id]
        if len(series_id) == len(df) + 1:
            series_id = series_id.iloc[1:]
        series_id = series_id.reset_index(drop=True).astype(str).str.strip()
        ids_norm = {str(x).strip() for x in ids}
        mask = series_id.isin(ids_norm)
        return mask
    return None


def get_responded_mask(df: pd.DataFrame, df_raw: pd.DataFrame, respuesta_letter: str | None, base_mask: pd.Series | None) -> pd.Series:
    if respuesta_letter:
        series = get_aligned_series_by_letter(df, df_raw, respuesta_letter)
        s = series.astype(str)
        # normalizar: reemplazar nbsp y tabs por espacio, luego strip y lower
        s_norm = (
            s.str.replace('\xa0', ' ', regex=False)
             .str.replace('\t', ' ', regex=False)
             .str.replace('\r', ' ', regex=False)
             .str.replace('\n', ' ', regex=False)
             .str.strip()
             .str.lower()
        )
        empty_like = {'', 'nan', 'none', 'na', 'n/a', 'n\\a', '-', '–', '--', 's/d', 'sd'}
        base = ~s_norm.isin(empty_like)
        # adicional: considerar vacías las que no contienen caracteres alfanuméricos
        has_alnum = s_norm.str.contains(r"[0-9a-zA-Z]", regex=True, na=False)
        mask = base & has_alnum
    else:
        # sin columna específica: una fila responde si tiene algún valor no vacío
        if len(df) == 0:
            return pd.Series([], dtype=bool)
        mask = (~df.isna()).any(axis=1)
    if base_mask is not None and len(base_mask) == len(mask):
        mask = mask & base_mask
    return mask


def count_respuestas(df: pd.DataFrame, df_raw: pd.DataFrame, respuesta_letter: str | None, base_mask: pd.Series | None) -> int:
    """Cuenta respuestas reales.
    Preferir columna específica (respuesta_letter) para determinar si una fila respondió.
    Si no se indica columna, considerar respondida si la fila tiene algún valor no vacío.
    Siempre intersectar con base_mask cuando exista (filtro por ID).
    """
    # Construir máscara de respondidos según columna indicada o por contenido de la fila
    resp_mask = get_responded_mask(df, df_raw, respuesta_letter, base_mask)
    if resp_mask is not None and len(resp_mask) == len(df):
        return int(resp_mask.sum())
    # Fallback robusto
    if base_mask is not None and len(base_mask) == len(df):
        return int(base_mask.sum())
    return int((~df.isna()).any(axis=1).sum())


def pct_cohort(df: pd.DataFrame, df_raw: pd.DataFrame, gen_letter: str, cohort_key: str, base_mask: pd.Series | None) -> float:
    gen_series = get_aligned_series_by_letter(df, df_raw, gen_letter)
    if gen_series.empty:
        return 0.0
    # Selección base: si hay filtro por ID, usarlo; si no, usar todas las filas
    if base_mask is not None and len(base_mask) == len(df):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(df))
    # Alinear longitudes
    n = min(len(gen_series), len(sel_mask))
    if n == 0:
        return 0.0
    gen_series = gen_series.iloc[:n]
    sel_mask = sel_mask.iloc[:n]
    denom = int(sel_mask.sum())
    if denom == 0:
        return 0.0
    variants = {normalize_label(v) for v in COHORT_VARIANTS.get(cohort_key.lower(), set())}
    if not variants:
        return 0.0
    gen_norm = gen_series.apply(normalize_label)
    gen_mask = gen_norm.isin(variants)
    num = int((sel_mask & gen_mask).sum())
    return round((num / denom) * 100.0, 2)


def avg_by_cohort(df: pd.DataFrame, df_raw: pd.DataFrame, value_letter: str, gen_letter: str, cohort_key: str, base_mask: pd.Series | None) -> float:
    vals = get_aligned_series_by_letter(df, df_raw, value_letter)
    gens = get_aligned_series_by_letter(df, df_raw, gen_letter)
    if vals.empty or gens.empty:
        return 0.0
    # Selección base como en Respuestas
    if base_mask is not None and len(base_mask) == len(df):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(df))
    # Alinear longitudes
    n = min(len(vals), len(gens), len(sel_mask))
    if n == 0:
        return 0.0
    vals = vals.iloc[:n]
    gens = gens.iloc[:n]
    sel_mask = sel_mask.iloc[:n]
    variants = {normalize_label(v) for v in COHORT_VARIANTS.get(cohort_key.lower(), set())}
    if not variants:
        return 0.0
    gen_norm = gens.apply(normalize_label)
    cohort_mask = gen_norm.isin(variants)
    final_mask = sel_mask & cohort_mask
    if int(final_mask.sum()) == 0:
        return 0.0
    nums = pd.to_numeric(vals, errors='coerce')
    nums = nums[final_mask.values]
    if nums.dropna().empty:
        return 0.0
    return round(float(nums.dropna().mean()), 2)


def avg_from_letter(df: pd.DataFrame, df_raw: pd.DataFrame, value_letter: str, base_mask: pd.Series | None) -> float:
    series = get_aligned_series_by_letter(df, df_raw, value_letter)
    if series.empty:
        return 0.0
    # Selección base como en Respuestas
    if base_mask is not None and len(base_mask) == len(df):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(df))
    # Alinear longitudes
    n = min(len(series), len(sel_mask))
    if n == 0:
        return 0.0
    series = series.iloc[:n]
    sel_mask = sel_mask.iloc[:n]
    # Convertir a numérico, ignorando no numéricos
    nums = pd.to_numeric(series, errors='coerce')
    nums = nums[sel_mask.values]
    if nums.dropna().empty:
        return 0.0
    return round(float(nums.dropna().mean()), 2)


def avg_from_letter_scale4_to10(df: pd.DataFrame, df_raw: pd.DataFrame, value_letter: str, base_mask: pd.Series | None) -> float:
    series = get_aligned_series_by_letter(df, df_raw, value_letter)
    if series.empty:
        return 0.0
    # Selección base como en Respuestas
    if base_mask is not None and len(base_mask) == len(df):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(df))
    # Alinear longitudes
    n = min(len(series), len(sel_mask))
    if n == 0:
        return 0.0
    series = series.iloc[:n]
    sel_mask = sel_mask.iloc[:n]
    # Convertir a numérico en escala 1–4 y quedarnos con filas seleccionadas
    nums = pd.to_numeric(series, errors='coerce')
    nums = nums[sel_mask.values].dropna()
    if nums.empty:
        return 0.0
    # Clamp 1..4 por estabilidad
    nums = nums.clip(lower=1.0, upper=4.0)
    # 1) Promedio en escala 1–4
    mean_4 = float(nums.mean())
    # 2) Redondear ese promedio a 1 decimal en escala 1–4
    mean_4_rounded = round(mean_4, 1)
    # 3) Convertir ese promedio redondeado a escala 10: f(x) = (x - 1) * (10/3)
    converted = (mean_4_rounded - 1.0) * (10.0 / 3.0)
    # 4) Redondear a 1 decimal el resultado en escala 10
    return round(float(converted), 1)


def avg_group_from_letters_scale4_to10(df: pd.DataFrame, df_raw: pd.DataFrame, letters: list[str], base_mask: pd.Series | None) -> float:
    """Promedio no ponderado de varias letras, cada una convertida 1–4 -> 10 con la misma regla.
    Equivale a la lógica en Power BI cuando hace promedio de subcomponentes "Actual".
    Ignora letras sin datos (> 0.0 criterio) y redondea el promedio final a 1 decimal.
    """
    vals = []
    for lt in letters:
        v = avg_from_letter_scale4_to10(df, df_raw, lt, base_mask)
        if v and v > 0.0:
            vals.append(float(v))
    if not vals:
        return 0.0
    return round(sum(vals) / len(vals), 1)


def compute_nps_from_letter(df: pd.DataFrame, df_raw: pd.DataFrame, nps_letter: str, base_mask: pd.Series | None):
    """Calcula NPS = (Promotores - Detractores) / Respuestas * 100.
    Usa la columna 'Tipo NPS' indicada por letra, sin conversión de escala.
    - Promotores: entusiasta(s), promotor(es)
    - Detractores: detractor(es)
    - Ignora pasivos/neutrales
    Denominador: mismo criterio que 'Respuestas':
      * con filtro por ID -> filas seleccionadas por ese filtro
      * sin filtro por ID -> todas las filas de datos menos la fila de título (primera)
    """
    if not nps_letter:
        return 0.0, 0, 0, 0
    idx = col_letter_to_index(nps_letter)
    if idx < 0 or idx >= df_raw.shape[1]:
        return 0.0, 0, 0, 0
    # Tomar solo filas de datos (df_raw tiene header en la primera fila)
    series = df_raw.iloc[1:, idx].reset_index(drop=True)
    # Selección base: si hay filtro por ID, usarlo directamente; si no, usar todas las filas
    if base_mask is not None:
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(series))
    # Normalizar etiquetas y descartar vacíos del NPS (mismo criterio que Respuestas)
    s_norm = series.astype(str).apply(normalize_label)
    empty_like = {'', 'nan', 'none', 'na', 'n/a', 'n\\a', '-', '–', '--', 's/d', 'sd'}
    has_alnum = s_norm.str.contains(r"[0-9a-zA-Z]", regex=True, na=False)
    non_empty = ~s_norm.isin(empty_like) & has_alnum
    final_mask = sel_mask & non_empty
    denom = int(final_mask.sum())
    if denom == 0:
        return 0.0, 0, 0, 0
    # Contar promotores/detractores
    promotor_variants = {
        'entusiasta', 'entusiastas', 'promotor', 'promotores', 'promoter', 'promoters'
    }
    detractor_variants = {
        'detractor', 'detractores'
    }
    is_prom = s_norm.isin(promotor_variants)
    is_det = s_norm.isin(detractor_variants)
    prom = int((is_prom & final_mask).sum())
    det = int((is_det & final_mask).sum())
    nps = ((prom - det) / denom) * 100.0
    return round(float(nps), 2), prom, det, denom


def compute_nps_by_cohort(df: pd.DataFrame, df_raw: pd.DataFrame, nps_letter: str, gen_letter: str, cohort_key: str, base_mask: pd.Series | None) -> float:
    """NPS por cohorte = (Promotores - Detractores) / Respuestas_cohorte * 100.
    Cohorte definida por gen_letter y COHORT_VARIANTS[cohort_key].
    """
    if not nps_letter or not gen_letter:
        return 0.0
    idx = col_letter_to_index(nps_letter)
    if idx < 0 or idx >= df_raw.shape[1]:
        return 0.0
    # Series NPS y generación alineadas a df
    nps_series = df_raw.iloc[1:, idx].reset_index(drop=True)
    gen_series = get_aligned_series_by_letter(df, df_raw, gen_letter)
    if nps_series.empty or gen_series.empty:
        return 0.0
    # Selección base
    if base_mask is not None and len(base_mask) == len(nps_series):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(nps_series))
    # Cohorte
    variants = {normalize_label(v) for v in COHORT_VARIANTS.get(cohort_key.lower(), set())}
    if not variants:
        return 0.0
    gen_norm = gen_series.apply(normalize_label)
    cohort_mask = gen_norm.isin(variants)
    # Alinear longitudes y combinar
    n = min(len(nps_series), len(cohort_mask), len(sel_mask))
    if n == 0:
        return 0.0
    nps_series = nps_series.iloc[:n]
    cohort_mask = cohort_mask.iloc[:n]
    sel_mask = sel_mask.iloc[:n]
    # Descartar NPS vacíos
    s_norm = nps_series.astype(str).apply(normalize_label)
    empty_like = {'', 'nan', 'none', 'na', 'n/a', 'n\\a', '-', '–', '--', 's/d', 'sd'}
    has_alnum = s_norm.str.contains(r"[0-9a-zA-Z]", regex=True, na=False)
    non_empty = ~s_norm.isin(empty_like) & has_alnum
    final_mask = sel_mask & cohort_mask & non_empty
    denom = int(final_mask.sum())
    if denom == 0:
        return 0.0
    # Clasificación promotores/detractores
    promotor_variants = {'entusiasta', 'entusiastas', 'promotor', 'promotores', 'promoter', 'promoters'}
    detractor_variants = {'detractor', 'detractores'}
    is_prom = s_norm.isin(promotor_variants)
    is_det = s_norm.isin(detractor_variants)
    prom = int((is_prom & final_mask).sum())
    det = int((is_det & final_mask).sum())
    nps = ((prom - det) / denom) * 100.0
    return round(float(nps), 2)


def compute_nps_category_pct(df: pd.DataFrame, df_raw: pd.DataFrame, nps_letter: str, category: str, base_mask: pd.Series | None) -> float:
    """Porcentaje de una categoría NPS: entusiastas/promotores, pasivos, detractores.
    Usa mismo denominador que NPS general (con/sin filtro por ID).
    """
    if not nps_letter:
        return 0.0
    idx = col_letter_to_index(nps_letter)
    if idx < 0 or idx >= df_raw.shape[1]:
        return 0.0
    series = df_raw.iloc[1:, idx].reset_index(drop=True)
    # Selección base
    if base_mask is not None and len(base_mask) == len(series):
        sel_mask = base_mask.copy()
    else:
        sel_mask = build_row_mask_all_true(len(series))
    s_norm = series.astype(str).apply(normalize_label)
    empty_like = {'', 'nan', 'none', 'na', 'n/a', 'n\\a', '-', '–', '--', 's/d', 'sd'}
    has_alnum = s_norm.str.contains(r"[0-9a-zA-Z]", regex=True, na=False)
    non_empty = ~s_norm.isin(empty_like) & has_alnum
    final_mask = sel_mask & non_empty
    denom = int(final_mask.sum())
    if denom == 0:
        return 0.0
    promotor_variants = {'entusiasta', 'entusiastas', 'promotor', 'promotores', 'promoter', 'promoters'}
    detractor_variants = {'detractor', 'detractores'}
    pasivo_variants = {'pasivo', 'pasivos', 'neutral', 'neutrales', 'indiferente', 'indiferentes'}
    if category.lower() in {'entusiasta', 'promotor', 'promotores', 'promoter', 'promoters'}:
        mask_cat = s_norm.isin(promotor_variants)
    elif category.lower() in {'pasivo', 'pasivos', 'neutral', 'neutrales', 'indiferente', 'indiferentes'}:
        mask_cat = s_norm.isin(pasivo_variants)
    else:
        mask_cat = s_norm.isin(detractor_variants)
    num = int((mask_cat & final_mask).sum())
    return round((num / denom) * 100.0, 2)


def generar():
    parser = argparse.ArgumentParser(description='Generar hoja "consolidado" con métricas paso a paso.')
    parser.add_argument('--sheet', help='Hoja de entrada (por defecto Datos A2)')
    parser.add_argument('--id', dest='ids', nargs='*', help='Filtrar por uno o varios ID (columna A)')
    parser.add_argument('--respuesta-col', help='Letra de columna que indica respuesta (ej: GV para NPS, H para comentario).')
    parser.add_argument('--gen-col', default='GU', help='Letra de columna para "Generación" (por defecto GU).')
    parser.add_argument('--satisf-col', default='E', help='Letra de columna para satisfacción general (por defecto E).')
    parser.add_argument('--clima-col', default='IA', help='Letra de columna para "Indice de clima" (por defecto IA).')
    parser.add_argument('--nps-col', default='GV', help='Letra de columna para "Tipo NPS" (por defecto GV).')
    parser.add_argument('--universo', type=float, help='Tamaño del universo. Si no se indica, se pedirá por consola.')
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    in_path = base_dir / INPUT_FILE
    if not in_path.exists():
        alt = Path.cwd() / INPUT_FILE
        if alt.exists():
            in_path = alt
        else:
            raise SystemExit(f'No se encuentra {INPUT_FILE} en {base_dir} ni en {Path.cwd()}')

    xls = open_excel_safely(in_path)
    hoja = find_sheet(xls, args.sheet or DEFAULT_SHEET)
    df = xls.parse(hoja)
    df_raw = xls.parse(hoja, header=None)

    # Construir máscara por ID(s) en columna A si se suministra
    id_mask = build_id_mask(df, df_raw, args.ids)
    # Si se desea, también filtramos df para mostrar consistencia en conteos impresos
    if id_mask is not None and len(id_mask) == len(df):
        df = df[id_mask.values].reset_index(drop=True)

    # 0) Universo: tomar argumento o solicitar por consola
    universo = args.universo
    if universo is None:
        while True:
            raw = input('Ingrese el valor de Universo: ').strip()
            if raw == '':
                print('No ingresaste un valor. Se usará 0.')
                universo = 0.0
                break
            try:
                universo = float(raw.replace(',', '.'))
                break
            except Exception:
                print('Valor no válido. Ingresa un número (puedes usar coma o punto).')

    # 1) Respuestas: por defecto, si no se indica --respuesta-col, usar la columna de NPS
    respuesta_letter = args.respuesta_col if args.respuesta_col else args.nps_col
    respuestas = count_respuestas(df, df_raw, respuesta_letter, id_mask)

    # 2) % Participación = respuestas / universo (en porcentaje) con 2 decimales
    if universo and universo > 0:
        participacion = round((respuestas / universo) * 100.0, 2)
    else:
        participacion = 0.0

    # 3) % Cohortes (sobre "Respuestas" según la regla aplicada)
    pct_centenialls = pct_cohort(df, df_raw, args.gen_col, 'centenialls', id_mask)
    pct_milenialls = pct_cohort(df, df_raw, args.gen_col, 'milenialls', id_mask)
    pct_genx = pct_cohort(df, df_raw, args.gen_col, 'generacion x', id_mask)
    pct_boomers = pct_cohort(df, df_raw, args.gen_col, 'baby boomers', id_mask)
    # 4) Nivel general de satisfacción (promedio de columna E)
    nivel_satisfaccion = avg_from_letter(df, df_raw, args.satisf_col, id_mask)
    # 5) (se eliminan las columnas por generación para condiciones de trabajo según nuevo orden)
    # 6) Condiciones de trabajo (promedio total de columna G)
    nivel_satisfaccion_condiciones = avg_from_letter(df, df_raw, 'G', id_mask)

    # 7) Indice de clima (promedio no ponderado de los 5 ejes principales, conversión 4->10)
    # Ejes como promedio de subcomponentes según fórmulas de Power BI:
    # 1. Trabajar en la organización (IB) = promedio(IC, ID, IE, IF)
    # 2. Soporte para resultados del cargo (IG) = promedio(IH, II, IJ, IK)
    # 3. Dinámica de trabajo conjunto (IL) = promedio(IM, IN, IO, IP)
    # 4. Cultura (IQ) = promedio(IR, IS, IT, IU)
    # 5. Ambiente sano, limpio y seguro (IV) = promedio(IW, IX, IY, IZ)
    trabajar_org = avg_group_from_letters_scale4_to10(df, df_raw, ['IC', 'ID', 'IE', 'IF'], id_mask)
    q_ig = avg_group_from_letters_scale4_to10(df, df_raw, ['IH', 'II', 'IJ', 'IK'], id_mask)
    q_il = avg_group_from_letters_scale4_to10(df, df_raw, ['IM', 'IN', 'IO', 'IP'], id_mask)
    q_iq = avg_group_from_letters_scale4_to10(df, df_raw, ['IR', 'IS', 'IT', 'IU'], id_mask)
    q_iv = avg_group_from_letters_scale4_to10(df, df_raw, ['IW', 'IX', 'IY', 'IZ'], id_mask)
    indice_clima = round(sum([v for v in [trabajar_org, q_ig, q_il, q_iq, q_iv] if v and v > 0.0]) /
                         max(1, len([v for v in [trabajar_org, q_ig, q_il, q_iq, q_iv] if v and v > 0.0])), 1)

    # 8) NPS (con categorías en columna Tipo NPS)
    nps, nps_prom, nps_det, nps_denom = compute_nps_from_letter(df, df_raw, args.nps_col, id_mask)
    # 9) NPS por cohorte
    nps_cent = compute_nps_by_cohort(df, df_raw, args.nps_col, args.gen_col, 'centenialls', id_mask)
    nps_mil = compute_nps_by_cohort(df, df_raw, args.nps_col, args.gen_col, 'milenialls', id_mask)
    nps_x = compute_nps_by_cohort(df, df_raw, args.nps_col, args.gen_col, 'generacion x', id_mask)
    nps_boom = compute_nps_by_cohort(df, df_raw, args.nps_col, args.gen_col, 'baby boomers', id_mask)
    # 10) Porcentajes por categoría NPS (globales)
    nps_pct_ent = compute_nps_category_pct(df, df_raw, args.nps_col, 'entusiasta', id_mask)
    nps_pct_pas = compute_nps_category_pct(df, df_raw, args.nps_col, 'pasivo', id_mask)
    nps_pct_det = compute_nps_category_pct(df, df_raw, args.nps_col, 'detractor', id_mask)

    # 12-19) Ocho preguntas AL..AS con conversión 4->10
    q_al = avg_from_letter_scale4_to10(df, df_raw, 'AL', id_mask)
    q_am = avg_from_letter_scale4_to10(df, df_raw, 'AM', id_mask)
    q_an = avg_from_letter_scale4_to10(df, df_raw, 'AN', id_mask)
    q_ao = avg_from_letter_scale4_to10(df, df_raw, 'AO', id_mask)
    q_ap = avg_from_letter_scale4_to10(df, df_raw, 'AP', id_mask)
    q_aq = avg_from_letter_scale4_to10(df, df_raw, 'AQ', id_mask)
    q_ar = avg_from_letter_scale4_to10(df, df_raw, 'AR', id_mask)
    q_as = avg_from_letter_scale4_to10(df, df_raw, 'AS', id_mask)
    # (Se omiten asignaciones individuales de IG/IL porque ya fueron calculados como promedios de sus subcomponentes)
    # 21-28) Preguntas AU..BB con conversión 4->10
    q_au = avg_from_letter_scale4_to10(df, df_raw, 'AU', id_mask)
    q_av = avg_from_letter_scale4_to10(df, df_raw, 'AV', id_mask)
    q_aw = avg_from_letter_scale4_to10(df, df_raw, 'AW', id_mask)
    q_ax = avg_from_letter_scale4_to10(df, df_raw, 'AX', id_mask)
    q_ay = avg_from_letter_scale4_to10(df, df_raw, 'AY', id_mask)
    q_az = avg_from_letter_scale4_to10(df, df_raw, 'AZ', id_mask)
    q_ba = avg_from_letter_scale4_to10(df, df_raw, 'BA', id_mask)
    q_bb = avg_from_letter_scale4_to10(df, df_raw, 'BB', id_mask)
    # 29-36) Preguntas BD..BK con conversión 4->10
    q_bd = avg_from_letter_scale4_to10(df, df_raw, 'BD', id_mask)
    q_be = avg_from_letter_scale4_to10(df, df_raw, 'BE', id_mask)
    q_bf = avg_from_letter_scale4_to10(df, df_raw, 'BF', id_mask)
    q_bg = avg_from_letter_scale4_to10(df, df_raw, 'BG', id_mask)
    q_bh = avg_from_letter_scale4_to10(df, df_raw, 'BH', id_mask)
    q_bi = avg_from_letter_scale4_to10(df, df_raw, 'BI', id_mask)
    q_bj = avg_from_letter_scale4_to10(df, df_raw, 'BJ', id_mask)
    q_bk = avg_from_letter_scale4_to10(df, df_raw, 'BK', id_mask)
    # (Se omite la asignación individual de IQ; usamos el promedio de IR, IS, IT, IU)
    # 38-45) Preguntas BM..BT con conversión 4->10
    q_bm = avg_from_letter_scale4_to10(df, df_raw, 'BM', id_mask)
    q_bn = avg_from_letter_scale4_to10(df, df_raw, 'BN', id_mask)
    q_bo = avg_from_letter_scale4_to10(df, df_raw, 'BO', id_mask)
    q_bp = avg_from_letter_scale4_to10(df, df_raw, 'BP', id_mask)
    q_bq = avg_from_letter_scale4_to10(df, df_raw, 'BQ', id_mask)
    q_br = avg_from_letter_scale4_to10(df, df_raw, 'BR', id_mask)
    q_bs = avg_from_letter_scale4_to10(df, df_raw, 'BS', id_mask)
    q_bt = avg_from_letter_scale4_to10(df, df_raw, 'BT', id_mask)
    # (Se omite la asignación individual de IV; usamos el promedio de IW, IX, IY, IZ)
    # 47-54) Preguntas BV..CC con conversión 4->10
    q_bv = avg_from_letter_scale4_to10(df, df_raw, 'BV', id_mask)
    q_bw = avg_from_letter_scale4_to10(df, df_raw, 'BW', id_mask)
    q_bx = avg_from_letter_scale4_to10(df, df_raw, 'BX', id_mask)
    q_by = avg_from_letter_scale4_to10(df, df_raw, 'BY', id_mask)
    q_bz = avg_from_letter_scale4_to10(df, df_raw, 'BZ', id_mask)
    q_ca = avg_from_letter_scale4_to10(df, df_raw, 'CA', id_mask)
    q_cb = avg_from_letter_scale4_to10(df, df_raw, 'CB', id_mask)
    q_cc = avg_from_letter_scale4_to10(df, df_raw, 'CC', id_mask)

    # 6) Satisfacción por generación (promedio de satisf-col por cohorte)
    sat_centenialls = avg_by_cohort(df, df_raw, args.satisf_col, args.gen_col, 'centenialls', id_mask)
    sat_milenialls = avg_by_cohort(df, df_raw, args.satisf_col, args.gen_col, 'milenialls', id_mask)
    sat_genx = avg_by_cohort(df, df_raw, args.satisf_col, args.gen_col, 'generacion x', id_mask)
    sat_boomers = avg_by_cohort(df, df_raw, args.satisf_col, args.gen_col, 'baby boomers', id_mask)

    # Preparar DataFrame en el orden definido
    out_row = {
        'Total Personas': universo,
        'Respuestas': respuestas,
        '% Participación': participacion,
        '% Centenialls': pct_centenialls,
        '% Milenialls': pct_milenialls,
        '% Generación X': pct_genx,
        '% Baby Boomers': pct_boomers,
        'Satisfacción General': nivel_satisfaccion,
        'Satisfacción Centenialls': sat_centenialls,
        'Satisfacción Milenialls': sat_milenialls,
        'Satisfacción Generación X': sat_genx,
        'Satisfacción Baby Boomers': sat_boomers,
        'COVID': '',
        'Condiciones de trabajo': nivel_satisfaccion_condiciones,
        'Indice de clima': indice_clima,
        'NPS': nps,
        'NPS\nCentenialls': nps_cent,
        'NPS\nMilenialls': nps_mil,
        'NPS\nGeneración X': nps_x,
        'NPS\nBaby Boomers': nps_boom,
        'NPS Entusiastas': nps_pct_ent,
        'NPS Pasivos': nps_pct_pas,
        'NPS Detractores': nps_pct_det,
        '¿Cómo es trabajar en la organización?': trabajar_org,
        'Me siento orgulloso(a) cuando le cuento a otros que estoy trabajando en esta Organización.': q_al,
        'Recomendaría a otros trabajar en esta Organización.': q_am,
        'Me parece inspiradora la misión de la Organización y estoy comprometido (a) con ella.': q_an,
        'Comprendo los objetivos de la Organzación y sé como puedo aportar a conseguirlos desde mi cargo.': q_ao,
        'Se reciben beneficios no monetarios que hacen más agradable la experiencia en la Organización.': q_ap,
        'Se realizan actividades de bienestar que brindan espacios de esparcimiento para el empleado y su familia.': q_aq,
    'Las características de mi trabajo me permiten tener un adecuado balance entre mi trabajo y mi vida personal.': q_ar,
        'La empresa demuestra sensibilidad con las particularidades de la vida personal de sus empleados.': q_as,
        '¿Cómo lo soporta para lograr los resultados de su cargo?': q_ig,
        'Cuento con los recursos mínimos que requiero para realizar de forma adecuada mi trabajo.': q_au,
        'Puedo acceder a toda la información que necesito para hacer bien mi trabajo.': q_av,
        'Siento que mi cargo me brinda la oportunidad de progresar y desarrollarme.': q_aw,
        'Todos en la organización tienen claro su cargo y comprenden bien el alcance sus responsabilidades.': q_ax,
        'La formación que brinda la Organización realmente me ayuda a realizar mejor mi trabajo.': q_ay,
        'El entrenamiento que recibo en mi cargo me ayuda progresar a nivel personal y profesional.': q_az,
        'Me evalúan por mi desempeño y tengo opciones de recibir reconocimiento si he realizado un esfuerzo por hacer mi trabajo de manera sobresaliente.': q_ba,
    'En esta Organización celebramos los éxitos y nos felicitamos por los logros obtenidos.': q_bb,
    '¿Cómo es la dinámica del trabajo conjunto para lograr resultados?': q_il,
    'La comunicación entre áreas fluye de manera adecuada.': q_bd,
    'Mi jefe se comunica de forma clara y verifica que yo haya comprendido lo que me dice.': q_be,
    'Conozco personas de diferentes áreas a la mía y se puede trabajar en un entorno agradable, de camaradería y cooperación.': q_bf,
    'En mi área se facilita y promueve el trabajo en equipo, la integración y la colaboración entre compañeros.': q_bg,
    'En la Organización las decisiones se toman de forma oportuna y no se aplazan decisiones importantes.': q_bh,
    'Las decisiones que se toman en la organización se hacen con base en análisis completos y sistemáticos.': q_bi,
    'Los líderes en la organización inspiran, motivan a los equipos y dan ejemplo con su comportamiento.': q_bj,
    'Los jefes forman a sus colaboradores y les dan una adecuada y oportuna retroalimentación sobre su desempeño.': q_bk,
    '¿Cómo es la cultura?': q_iq,
    'Los jefes y directivos son cordiales y respetuosos en el trato con personas de todos los niveles en la Organización.': q_bm,
    'Puedo dar mi punto de vista y sugerencias a la organización sabiendo que de ser posible se van a tener en cuenta.': q_bn,
    'Siento que las relaciones en esta Organización se establecen sobre la base de la confianza en las personas.': q_bo,
    'Puedo determinar formas propias de hacer mis labores mientras cumpla con las políticas de la Organización.': q_bp,
    'En esta Organización el trato es justo y equitativo, sin discriminiaciones por género, edad, raza u orientación sexual.': q_bq,
    'En esta Organización se evita que se utilice la intimidación y el hostigamiento para inducir la renuncia de algún empleado.': q_br,
    'Yo quiero a la Organización y me siento comprometido con mi trabajo.': q_bs,
    'Me gusta mi trabajo y siento que estoy haciendo algo de valor para la Organización.': q_bt,
    '¿Cómo genera un ambiente sano, limpio y seguro?': q_iv,
    'Puedo desempeñar mi trabajo de forma segura y cómoda en cuanto a temperatura, sonido, espacio e iluminación.': q_bv,
    'Puedo manejar el estrés que se genera en mi trabajo para evitar que se afecte mi salud.': q_bw,
    'Los procesos de la organización cuentan con controles efectivos de calidad.': q_bx,
    'Se revisa de forma constante la calidad de los procesos y servicios en busca de un mejoramiento continuo.': q_by,
    'Se nota una sensibilidad en la Organización por cuidar el impacto que su operación pueda generar en el medio ambiente.': q_bz,
    'La Organización implementa políticas y/o procedimientos para que los empleados cuidemos el medio ambiente.': q_ca,
    'La Organización contrata empleados y proveedores cumpliendo con todas los requisitos legales y pagando oportunamente.': q_cb,
    'La Organización cuida el impacto que puede tener en la comunidad en donde opera y busca aportar a la misma.': q_cc,
    }
    df_out = pd.DataFrame([out_row], columns=CONSOLIDADO_COLS)

    final_path = write_excel_safely(base_dir / OUTPUT_FILE, {SHEET_OUT: df_out})
    print(f'Hoja "{SHEET_OUT}" escrita en {final_path}')
    print('Valores:')
    converted_cols_one_decimal = {
        # Conversión 4->10
        'Indice de clima',
        '¿Cómo es trabajar en la organización?',
        # AL–AS
        'Me siento orgulloso(a) cuando le cuento a otros que estoy trabajando en esta Organización.',
        'Recomendaría a otros trabajar en esta Organización.',
        'Me parece inspiradora la misión de la Organización y estoy comprometido (a) con ella.',
        'Comprendo los objetivos de la Organzación y sé como puedo aportar a conseguirlos desde mi cargo.',
        'Se reciben beneficios no monetarios que hacen más agradable la experiencia en la Organización.',
        'Se realizan actividades de bienestar que brindan espacios de esparcimiento para el empleado y su familia.',
        'Las características de mi trabajo me permiten tener un adecuado balance entre mi trabajo y mi vida personal.',
        'La empresa demuestra sensibilidad con las particularidades de la vida personal de sus empleados.',
        # IG / IL
        '¿Cómo lo soporta para lograr los resultados de su cargo?',
        '¿Cómo es la dinámica del trabajo conjunto para lograr resultados?',
        # AU–BB
        'Cuento con los recursos mínimos que requiero para realizar de forma adecuada mi trabajo.',
        'Puedo acceder a toda la información que necesito para hacer bien mi trabajo.',
        'Siento que mi cargo me brinda la oportunidad de progresar y desarrollarme.',
        'Todos en la organización tienen claro su cargo y comprenden bien el alcance sus responsabilidades.',
        'La formación que brinda la Organización realmente me ayuda a realizar mejor mi trabajo.',
        'El entrenamiento que recibo en mi cargo me ayuda progresar a nivel personal y profesional.',
        'Me evalúan por mi desempeño y tengo opciones de recibir reconocimiento si he realizado un esfuerzo por hacer mi trabajo de manera sobresaliente.',
        'En esta Organización celebramos los éxitos y nos felicitamos por los logros obtenidos.',
        # BD–BK
        'La comunicación entre áreas fluye de manera adecuada.',
        'Mi jefe se comunica de forma clara y verifica que yo haya comprendido lo que me dice.',
        'Conozco personas de diferentes áreas a la mía y se puede trabajar en un entorno agradable, de camaradería y cooperación.',
        'En mi área se facilita y promueve el trabajo en equipo, la integración y la colaboración entre compañeros.',
        'En la Organización las decisiones se toman de forma oportuna y no se aplazan decisiones importantes.',
        'Las decisiones que se toman en la organización se hacen con base en análisis completos y sistemáticos.',
        'Los líderes en la organización inspiran, motivan a los equipos y dan ejemplo con su comportamiento.',
        'Los jefes forman a sus colaboradores y les dan una adecuada y oportuna retroalimentación sobre su desempeño.',
        # IQ
        '¿Cómo es la cultura?',
        # BM–BT
        'Los jefes y directivos son cordiales y respetuosos en el trato con personas de todos los niveles en la Organización.',
        'Puedo dar mi punto de vista y sugerencias a la organización sabiendo que de ser posible se van a tener en cuenta.',
        'Siento que las relaciones en esta Organización se establecen sobre la base de la confianza en las personas.',
        'Puedo determinar formas propias de hacer mis labores mientras cumpla con las políticas de la Organización.',
        'En esta Organización el trato es justo y equitativo, sin discriminiaciones por género, edad, raza u orientación sexual.',
        'En esta Organización se evita que se utilice la intimidación y el hostigamiento para inducir la renuncia de algún empleado.',
        'Yo quiero a la Organización y me siento comprometido con mi trabajo.',
        'Me gusta mi trabajo y siento que estoy haciendo algo de valor para la Organización.'
        ,
        '¿Cómo genera un ambiente sano, limpio y seguro?',
        'Puedo desempeñar mi trabajo de forma segura y cómoda en cuanto a temperatura, sonido, espacio e iluminación.',
        'Puedo manejar el estrés que se genera en mi trabajo para evitar que se afecte mi salud.',
        'Los procesos de la organización cuentan con controles efectivos de calidad.',
        'Se revisa de forma constante la calidad de los procesos y servicios en busca de un mejoramiento continuo.',
        'Se nota una sensibilidad en la Organización por cuidar el impacto que su operación pueda generar en el medio ambiente.',
        'La Organización implementa políticas y/o procedimientos para que los empleados cuidemos el medio ambiente.',
        'La Organización contrata empleados y proveedores cumpliendo con todas los requisitos legales y pagando oportunamente.',
        'La Organización cuida el impacto que puede tener en la comunidad en donde opera y busca aportar a la misma.'
    }
    for k in CONSOLIDADO_COLS:
        v = out_row.get(k)
        if isinstance(v, float):
            if k in converted_cols_one_decimal:
                print(f'  {k}: {v:.1f}')
            else:
                print(f'  {k}: {v:.2f}')
        else:
            print(f'  {k}: {v}')
    # Ya no se imprime detalle de NPS en consola


if __name__ == '__main__':
    generar()

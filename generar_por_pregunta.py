from pathlib import Path
import argparse
import os
import shutil
import sys
import tempfile
import time
from difflib import SequenceMatcher
import unicodedata
import pandas as pd

INPUT_FILE = 'input.xlsx'
TEMPLATE_FILE = 'modelo output.xlsx'
OUTPUT_FILE = 'salida.xlsx'
DEFAULT_SHEET = 'Datos A2'
SHEET_OUT = 'Por pregunta'

# Mapear cada pregunta a su letra de columna en el input
QUESTION_TO_LETTER: dict[str, str] = {
    # IB
    '¿Cómo es trabajar en la organización?': 'IB',
    # AL..AS
    'Me siento orgulloso(a) cuando le cuento a otros que estoy trabajando en esta Organización.': 'AL',
    'Recomendaría a otros trabajar en esta Organización.': 'AM',
    'Me parece inspiradora la misión de la Organización y estoy comprometido (a) con ella.': 'AN',
    'Comprendo los objetivos de la organización y sé cómo puedo aportar a conseguirlos desde mi cargo.': 'AO',
    'Se reciben beneficios no monetarios que hacen más agradable la experiencia en la Organización.': 'AP',
    'Se realizan actividades de bienestar que brindan espacios de esparcimiento para el empleado y su familia.': 'AQ',
    'Las características de mi trabajo me permiten tener un adecuado balance entre mi trabajo y mi vida personal.': 'AR',
    'La empresa demuestra sensibilidad con las particularidades de la vida personal de sus empleados.': 'AS',
    # IG, IL
    '¿Cómo lo soporta para lograr los resultados de su cargo?': 'IG',
    '¿Cómo es la dinámica del trabajo conjunto para lograr resultados?': 'IL',
    # AU..BB
    'Cuento con los recursos mínimos que requiero para realizar de forma adecuada mi trabajo.': 'AU',
    'Puedo acceder a toda la información que necesito para hacer bien mi trabajo.': 'AV',
    'Siento que mi cargo me brinda la oportunidad de progresar y desarrollarme.': 'AW',
    'Todos en la organización tienen claro su cargo y comprenden bien el alcance de sus responsabilidades.': 'AX',
    'La formación que brinda la Organización realmente me ayuda a realizar mejor mi trabajo.': 'AY',
    'El entrenamiento que recibo en mi cargo me ayuda a progresar a nivel personal y profesional.': 'AZ',
    'Me evalúan por mi desempeño y tengo opciones de recibir reconocimiento si he realizado un esfuerzo por hacer mi trabajo de manera sobresaliente.': 'BA',
    'En esta Organización celebramos los éxitos y nos felicitamos por los logros obtenidos.': 'BB',
    # BD..BK
    'La comunicación entre áreas fluye de manera adecuada.': 'BD',
    'Mi jefe se comunica de forma clara y verifica que yo haya comprendido lo que me dice.': 'BE',
    'Conozco personas de diferentes áreas a la mía y se puede trabajar en un entorno agradable, de camaradería y cooperación.': 'BF',
    'En mi área se facilita y promueve el trabajo en equipo, la integración y la colaboración entre compañeros.': 'BG',
    'En la Organización las decisiones se toman de forma oportuna y no se aplazan decisiones importantes.': 'BH',
    'Las decisiones que se toman en la organización se hacen con base en análisis completos y sistemáticos.': 'BI',
    'Los líderes en la organización inspiran, motivan a los equipos y dan ejemplo con su comportamiento.': 'BJ',
    'Los jefes forman a sus colaboradores y les dan una adecuada y oportuna retroalimentación sobre su desempeño.': 'BK',
    # IQ
    '¿Cómo es la cultura?': 'IQ',
    # BM..BT
    'Los jefes y directivos son cordiales y respetuosos en el trato con personas de todos los niveles en la Organización.': 'BM',
    'Puedo dar mi punto de vista y sugerencias a la organización sabiendo que de ser posible se van a tener en cuenta.': 'BN',
    'Siento que las relaciones en esta Organización se establecen sobre la base de la confianza en las personas.': 'BO',
    'Puedo determinar formas propias de hacer mis labores mientras cumpla con las políticas de la Organización.': 'BP',
    'En esta Organización el trato es justo y equitativo, sin discriminaciones por género, edad, raza u orientación sexual.': 'BQ',
    'En esta Organización se evita que se utilice la intimidación y el hostigamiento para inducir la renuncia de algún empleado.': 'BR',
    'Yo quiero a la Organización y me siento comprometido con mi trabajo.': 'BS',
    'Me gusta mi trabajo y siento que estoy haciendo algo de valor para la Organización.': 'BT',
    # IV
    '¿Cómo genera un ambiente sano, limpio y seguro?': 'IV',
    # BV..CC
    'Puedo desempeñar mi trabajo de forma segura y cómoda en cuanto a temperatura, sonido, espacio e iluminación.': 'BV',
    'Puedo manejar el estrés que se genera en mi trabajo para evitar que se afecte mi salud.': 'BW',
    'Los procesos de la organización cuentan con controles efectivos de calidad.': 'BX',
    'Se revisa de forma constante la calidad de los procesos y servicios en busca de un mejoramiento continuo.': 'BY',
    'Se nota una sensibilidad en la Organización por cuidar el impacto que su operación pueda generar en el medio ambiente.': 'BZ',
    'La Organización implementa políticas y/o procedimientos para que los empleados cuidemos el medio ambiente.': 'CA',
    'La Organización contrata empleados y proveedores cumpliendo con todas los requisitos legales y pagando oportunamente.': 'CB',
    'La Organización cuida el impacto que puede tener en la comunidad en donde opera y busca aportar a la misma.': 'CC',
}

# Cohortes (variantes normalizadas)
COHORT_VARIANTS = {
    'centenialls': {
        'centenialls', 'centennials', 'centennial', 'centenials', 'centenial',
        'generacion z', 'generación z', 'gen z', 'z'
    },
    'milenialls': {
        'milenialls', 'millennials', 'millenials', 'millennial', 'millenial',
        'generacion y', 'generación y', 'gen y', 'y'
    },
    'generacion x': {'generacion x', 'generación x', 'gen x', 'x'},
    'baby boomers': {'baby boomers', 'baby boomer', 'babyboomers', 'babyboomer', 'boomers', 'boomer'}
}


def normalize_label(s: str) -> str:
    if s is None:
        return ''
    s = str(s)
    # Quitar acentos
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    # Unificar espacios
    s = ' '.join(s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').split()).strip().lower()
    # Quitar signos de puntuacion para facilitar el match (.,:;!?¿¡()[]"')
    s = ''.join(ch for ch in s if ch.isalnum() or ch.isspace())
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


def build_row_mask_all_true(n: int) -> pd.Series:
    return pd.Series([True] * n)


def open_excel_safely(p: Path) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(p, engine='openpyxl')
    except PermissionError:
        tmpdir = Path(tempfile.gettempdir())
        last_err = None
        # Reintentos con backoff, intentando copiar a temp y leer desde allí
        for i in range(8):
            try:
                tmp_path = tmpdir / f"tmp_{p.stem}_{os.getpid()}_{i}.xlsx"
                shutil.copy2(p, tmp_path)
                return pd.ExcelFile(tmp_path, engine='openpyxl')
            except PermissionError as e:
                last_err = e
                time.sleep(1.0)
        print(f"No se pudo leer '{p}' porque está en uso. Cierra el archivo en Excel/OneDrive y vuelve a intentar.", file=sys.stderr)
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


def get_aligned_series_by_letter(df: pd.DataFrame, df_raw: pd.DataFrame, letter: str | None) -> pd.Series:
    if not letter:
        return pd.Series([], dtype=object)
    idx = col_letter_to_index(letter)
    if idx < 0 or idx >= df_raw.shape[1]:
        return pd.Series([], dtype=object)
    series = df_raw.iloc[:, idx]
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


def avg_4_to_10_for_mask(
    df: pd.DataFrame,
    df_raw: pd.DataFrame,
    value_letter: str,
    select_mask: pd.Series,
    debug: bool = False,
    dbg_label: str = '',
    dbg_cohort: str = ''
) -> float:
    if select_mask is None or len(select_mask) == 0:
        return 0.0
    series = get_aligned_series_by_letter(df, df_raw, value_letter)
    if series.empty:
        return 0.0
    n = min(len(series), len(select_mask))
    if n == 0:
        return 0.0
    series = series.iloc[:n]
    sel_mask = select_mask.iloc[:n]
    total_masked = int(sel_mask.sum())
    nums_all = pd.to_numeric(series, errors='coerce')
    nums = nums_all[sel_mask.values].dropna()
    valid_count = int(nums.shape[0])
    if nums.empty:
        if debug:
            print(f"[Calc] Pregunta='{dbg_label}' Cohorte='{dbg_cohort}' Col='{value_letter}' n_mask={total_masked} n_valid=0 -> sin datos", file=sys.stderr)
        return 0.0
    nums = nums.clip(lower=1.0, upper=4.0)
    mean_4 = float(nums.mean())
    mean_4_rounded = round(mean_4, 1)
    converted = (mean_4_rounded - 1.0) * (10.0 / 3.0)
    converted_rounded = round(float(converted), 1)
    if debug:
        print(
            f"[Calc] Pregunta='{dbg_label}' Cohorte='{dbg_cohort}' Col='{value_letter}' "
            f"n_mask={total_masked} n_valid={valid_count} mean_4={mean_4:.4f} mean_4_r={mean_4_rounded:.1f} "
            f"conv10={converted:.6f} conv10_r={converted_rounded:.1f}",
            file=sys.stderr
        )
    return converted_rounded


def make_gen_mask(df: pd.DataFrame, df_raw: pd.DataFrame, gen_letter: str, cohort_key: str) -> pd.Series:
    gens = get_aligned_series_by_letter(df, df_raw, gen_letter)
    if gens.empty:
        return build_row_mask_all_true(0)
    variants = {normalize_label(v) for v in COHORT_VARIANTS.get(cohort_key.lower(), set())}
    gen_norm = gens.apply(normalize_label)
    return gen_norm.isin(variants)


def make_nps_mask(df: pd.DataFrame, df_raw: pd.DataFrame, nps_letter: str, category: str) -> pd.Series:
    if not nps_letter:
        return build_row_mask_all_true(0)
    series = get_aligned_series_by_letter(df, df_raw, nps_letter)
    if series.empty:
        return build_row_mask_all_true(0)
    s_norm = series.astype(str).apply(normalize_label)
    promotor_variants = {'entusiasta', 'entusiastas', 'promotor', 'promotores', 'promoter', 'promoters'}
    detractor_variants = {'detractor', 'detractores'}
    pasivo_variants = {'pasivo', 'pasivos', 'neutral', 'neutrales', 'indiferente', 'indiferentes'}
    cat = category.strip().lower()
    if cat in promotor_variants:
        mask = s_norm.isin(promotor_variants)
    elif cat in pasivo_variants:
        mask = s_norm.isin(pasivo_variants)
    else:
        mask = s_norm.isin(detractor_variants)
    return pd.Series(mask.values)


def load_template_rows(base_dir: Path, template_path: Path | None = None) -> pd.DataFrame | None:
    tpl_path = template_path if template_path else (base_dir / TEMPLATE_FILE)
    if not tpl_path.exists():
        return None
    try:
        xls = open_excel_safely(tpl_path)
        sheet_name = 'Por pregunta'
        if sheet_name not in xls.sheet_names:
            # buscar similar
            wn = sheet_name.lower()
            best = None
            best_score = 0.0
            for s in xls.sheet_names:
                score = SequenceMatcher(None, wn, s.strip().lower()).ratio()
                if score > best_score:
                    best = s
                    best_score = score
            sheet_name = best or xls.sheet_names[0]
        df_tpl = xls.parse(sheet_name)
        # esperamos columnas: index, Pregunta, Clima, IFE, PESO
        wanted = ['index', 'Pregunta', 'Clima', 'IFE', 'PESO']
        cols = [c for c in wanted if c in df_tpl.columns]
        if not cols:
            return None
        df_meta = df_tpl[cols].copy()
        # normalizar textos
        for c in ['index', 'Pregunta', 'Clima', 'IFE']:
            if c in df_meta.columns:
                df_meta[c] = df_meta[c].astype(str).fillna('').str.strip()
        if 'PESO' in df_meta.columns:
            df_meta['PESO'] = pd.to_numeric(df_meta['PESO'], errors='coerce')
        return df_meta
    except Exception as e:
        print(f'Advertencia: no se pudo leer la plantilla: {e}', file=sys.stderr)
        return None


def generar():
    parser = argparse.ArgumentParser(description='Generar hoja "Por pregunta" con promedios en escala 10.')
    parser.add_argument('--sheet', help='Hoja de entrada (por defecto Datos A2)')
    parser.add_argument('--id', dest='ids', nargs='*', help='Filtrar por uno o varios ID (columna A)')
    parser.add_argument('--gen-col', default='GU', help='Letra de columna para "Generación" (por defecto GU).')
    parser.add_argument('--nps-col', default='GV', help='Letra de columna para "Tipo NPS" (por defecto GV).')
    parser.add_argument('--debug', action='store_true', help='Imprimir detalles de mapeo y cálculo en consola.')
    parser.add_argument('--debug-summary', action='store_true', help='Imprimir un resumen por pregunta con todas las cohortes en una línea.')
    parser.add_argument('--template-file', dest='template_file', help='Ruta al archivo de plantilla (por defecto modelo output.xlsx)')
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent

    in_path = base_dir / INPUT_FILE
    if not in_path.exists():
        raise SystemExit(f'No se encuentra {INPUT_FILE} en {base_dir}')

    xls = open_excel_safely(in_path)
    hoja = args.sheet or DEFAULT_SHEET
    if hoja not in xls.sheet_names:
        # fuzzy match
        wn = hoja.strip().lower()
        best = None
        best_score = 0.0
        for s in xls.sheet_names:
            score = SequenceMatcher(None, wn, s.strip().lower()).ratio()
            if score > best_score:
                best = s
                best_score = score
        hoja = best or xls.sheet_names[0]

    df = xls.parse(hoja)
    df_raw = xls.parse(hoja, header=None)

    id_mask = build_id_mask(df, df_raw, args.ids)
    if id_mask is not None and len(id_mask) == len(df):
        # Filtrar df y df_raw de forma alineada (df_raw sin cabecera)
        df = df[id_mask.values].reset_index(drop=True)
        df_raw = df_raw.iloc[1:, :][id_mask.values].reset_index(drop=True)

    # Selección base: todo True del tamaño actual (tras posible filtrado)
    base_sel = build_row_mask_all_true(len(df))

    boom_mask = make_gen_mask(df, df_raw, args.gen_col, 'baby boomers')
    genx_mask = make_gen_mask(df, df_raw, args.gen_col, 'generacion x')
    mil_mask = make_gen_mask(df, df_raw, args.gen_col, 'milenialls')
    cen_mask = make_gen_mask(df, df_raw, args.gen_col, 'centenialls')

    ent_mask = make_nps_mask(df, df_raw, args.nps_col, 'entusiasta')
    pas_mask = make_nps_mask(df, df_raw, args.nps_col, 'pasivo')
    det_mask = make_nps_mask(df, df_raw, args.nps_col, 'detractor')

    # Cargar plantilla (index/Pregunta/Clima/IFE/PESO) o construir desde el mapping
    tpl_path = Path(args.template_file) if args.template_file else None
    df_meta = load_template_rows(base_dir, tpl_path)
    if df_meta is None:
        rows = []
        for label in QUESTION_TO_LETTER.keys():
            rows.append({'index': '', 'Pregunta': label, 'Clima': '', 'IFE': '', 'PESO': ''})
        df_meta = pd.DataFrame(rows, columns=['index', 'Pregunta', 'Clima', 'IFE', 'PESO'])
        print('Advertencia: se generó la hoja sin metadatos (Clima/IFE/PESO en blanco) porque no se pudo leer la plantilla.', file=sys.stderr)

    # Asegurar columnas objetivo
    out_cols = ['index', 'Pregunta', 'Clima', 'IFE', 'PESO',
                'General', 'Boomers', 'Generación X', 'Milenialls', 'Centenialls',
                'Entusiastas', 'Pasivos', 'Detractores']

    rows_out = []
    # mapa normalizado para busqueda flexible
    norm_map = {normalize_label(k): v for k, v in QUESTION_TO_LETTER.items()}
    not_matched: list[str] = []
    summaries = []
    for _, r in df_meta.iterrows():
        label = str(r.get('Pregunta', '')).strip()
        # Intento 1: exacto
        letter = QUESTION_TO_LETTER.get(label)
        match_method = None
        match_score = None
        if letter:
            match_method = 'exacto'
        # Intento 2: normalizado
        if not letter:
            nlabel = normalize_label(label)
            letter = norm_map.get(nlabel)
            if letter:
                match_method = 'normalizado'
        # Intento 3: fuzzy por normalizado
        if not letter:
            nlabel = normalize_label(label)
            best_score = 0.0
            best_letter = None
            for k_norm, v in norm_map.items():
                score = SequenceMatcher(None, nlabel, k_norm).ratio()
                if score > best_score:
                    best_score = score
                    best_letter = v
            if best_score >= 0.75:
                letter = best_letter
                match_method = 'fuzzy'
                match_score = best_score
        if not letter:
            not_matched.append(label)
            if args.debug:
                print(f"[Map] Pregunta no emparejada: '{label}'", file=sys.stderr)
        else:
            if args.debug:
                if match_method == 'fuzzy':
                    print(f"[Map] Pregunta='{label}' -> Col='{letter}' (metodo={match_method}, score={match_score:.2f})", file=sys.stderr)
                else:
                    print(f"[Map] Pregunta='{label}' -> Col='{letter}' (metodo={match_method})", file=sys.stderr)
        if letter:
            val_general = avg_4_to_10_for_mask(df, df_raw, letter, base_sel, args.debug, label, 'General')
            val_boom = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & boom_mask), args.debug, label, 'Boomers')
            val_genx = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & genx_mask), args.debug, label, 'Generación X')
            val_mil = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & mil_mask), args.debug, label, 'Milenialls')
            val_cen = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & cen_mask), args.debug, label, 'Centenialls')
            val_ent = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & ent_mask), args.debug, label, 'Entusiastas')
            val_pas = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & pas_mask), args.debug, label, 'Pasivos')
            val_det = avg_4_to_10_for_mask(df, df_raw, letter, (base_sel & det_mask), args.debug, label, 'Detractores')
        else:
            val_general = val_boom = val_genx = val_mil = val_cen = val_ent = val_pas = val_det = ''
        row = {
            'index': r.get('index', ''),
            'Pregunta': label,
            'Clima': r.get('Clima', ''),
            'IFE': r.get('IFE', ''),
            'PESO': r.get('PESO', ''),
            'General': val_general,
            'Boomers': val_boom,
            'Generación X': val_genx,
            'Milenialls': val_mil,
            'Centenialls': val_cen,
            'Entusiastas': val_ent,
            'Pasivos': val_pas,
            'Detractores': val_det,
        }
        rows_out.append(row)
        if args.debug_summary and letter:
            def fmt(x):
                return ("" if x == "" else f"{x:.1f}")
            summaries.append(
                f"[Resumen] '{label}' -> Gen={fmt(val_general)} | Boom={fmt(val_boom)} | X={fmt(val_genx)} | Mil={fmt(val_mil)} | Cen={fmt(val_cen)} | Ent={fmt(val_ent)} | Pas={fmt(val_pas)} | Det={fmt(val_det)}"
            )

    df_out = pd.DataFrame(rows_out, columns=out_cols)

    final_path = write_excel_safely(base_dir / OUTPUT_FILE, {SHEET_OUT: df_out})
    print(f'Hoja "{SHEET_OUT}" escrita en {final_path}')
    if not_matched:
        print('Advertencia: preguntas no emparejadas con columnas del input:', file=sys.stderr)
        for q in not_matched:
            print(f' - {q}', file=sys.stderr)
    if args.debug_summary and summaries:
        print('\n'.join(summaries), file=sys.stderr)


if __name__ == '__main__':
    generar()

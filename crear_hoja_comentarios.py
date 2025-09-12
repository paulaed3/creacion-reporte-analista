from pathlib import Path
import unicodedata
import argparse
import sys
from difflib import SequenceMatcher
import pandas as pd
import os
import shutil
import tempfile

INPUT_FILE = 'input.xlsx'
OUTPUT_FILE = 'salida.xlsx'  
SHEET_OUTPUT = 'comentarios'
SHEET_INPUT_PATTERN = 'Datos A2' 


TARGET_COLS = [
    'ID',
    'VARIABLE 1',
    'VARIABLE 2',
    'VARIABLE 3',
    'Tipo NPS',
    'Generación',
    'ANTIGÜEDAD',
    'TIPO CONTRATO',
    'COMENTARIOS',
]


def norm_text(s: str) -> str:
    if s is None:
        return ''
    s = str(s).strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
    s = ' '.join(s.replace('\n', ' ').replace('\r', ' ').replace('\t', ' ').split())
    return s.lower()


def find_sheet(xls: pd.ExcelFile, wanted: str | None) -> str:
    sheets = xls.sheet_names
    if not sheets:
        raise SystemExit('El archivo no contiene hojas.')
    if wanted:
        wanted_norm = norm_text(wanted)
        # Exacto/prefijo
        for s in sheets:
            s_norm = norm_text(s)
            if s_norm == wanted_norm or s_norm.startswith(wanted_norm):
                return s
        # Fuzzy por nombre de hoja
        best = None
        best_score = 0.0
        for s in sheets:
            score = SequenceMatcher(None, wanted_norm, norm_text(s)).ratio()
            if score > best_score:
                best = s
                best_score = score
        if best and best_score >= 0.6:
            return best
        # si solo hay una hoja, usarla
        if len(sheets) == 1:
            return sheets[0]
        raise SystemExit(f'No se encontró la hoja "{wanted}". Hojas disponibles: {sheets}')
    # Sin nombre deseado: si hay una sola hoja, usarla
    if len(sheets) == 1:
        return sheets[0]
    # fallback a patrón por defecto
    return find_sheet(xls, SHEET_INPUT_PATTERN)



def col_letter_to_index(letter: str) -> int:
    """Convierte letras de Excel (e.g., 'A', 'AA') a índice 0-based."""
    if not letter:
        return -1
    s = norm_text(letter).replace(' ', '')
    s = s.upper() if s else ''
    total = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'):
            continue
        total = total * 26 + (ord(ch) - ord('A') + 1)
    return total - 1  # 0-based


def crear_hoja_comentarios():
    parser = argparse.ArgumentParser(description='Generar hoja "comentarios" desde "Datos A2" filtrando opcionalmente por ID(s).')
    parser.add_argument('--id', dest='ids', nargs='*', help='Uno o varios ID a filtrar. Ej: --id 12345 67890')
    parser.add_argument('--sheet', dest='sheet', help='Nombre de la hoja de entrada si no es "Datos A2".')
    args = parser.parse_args()

    ids_norm = None
    if args.ids:
        ids_norm = {str(x).strip() for x in args.ids}

    base_dir = Path(__file__).resolve().parent
    path = (base_dir / INPUT_FILE)
    if not path.exists():
        alt = Path.cwd() / INPUT_FILE
        if alt.exists():
            path = alt
        else:
            raise SystemExit(
                f'No se encuentra {INPUT_FILE}. Buscado en: '\
                f'\n  - {base_dir}'\
                f'\n  - {Path.cwd()}'
            )

    def open_excel_safely(p: Path) -> pd.ExcelFile:
        try:
            return pd.ExcelFile(p, engine='openpyxl')
        except PermissionError:
            tmpdir = Path(tempfile.gettempdir())
            tmp_path = tmpdir / f"tmp_{p.stem}_{os.getpid()}.xlsx"
            shutil.copy2(p, tmp_path)
            return pd.ExcelFile(tmp_path, engine='openpyxl')

    xls = open_excel_safely(path)
    hoja = find_sheet(xls, args.sheet or SHEET_INPUT_PATTERN)
    df = xls.parse(hoja)
    df_raw = xls.parse(hoja, header=None)
    print(f'Usando hoja de entrada: {hoja}. Hojas disponibles: {xls.sheet_names}')

    # Si se solicitó filtrar por ID, construimos una máscara por letra (columna A)
    mask_keep = None
    if ids_norm:
        idx_id = col_letter_to_index('A')
        if 0 <= idx_id < df_raw.shape[1]:
            series_raw = df_raw.iloc[:, idx_id]
            # ajustar por posible fila de encabezado
            if len(series_raw) == len(df) + 1:
                series_raw = series_raw.iloc[1:]
            series_raw = series_raw.reset_index(drop=True).astype(str).str.strip()
            if len(series_raw) == len(df):
                mask_keep = series_raw.isin(ids_norm)
        if mask_keep is not None:
            df = df[mask_keep.values].reset_index(drop=True)
            if df.empty:
                print(f'Advertencia: no se encontraron filas para los ID proporcionados: {sorted(ids_norm)}')
        else:
            print('No se encontró columna de ID (letra A). No se aplicó filtro por ID.', file=sys.stderr)

    # Sobrescritura manual por letras de columna (proporcionadas por el usuario)
    MANUAL_LETTERS = {
        'ID': 'A',
        'VARIABLE 1': 'CM',
        'VARIABLE 2': 'CL',
        'VARIABLE 3': 'GS',
        'Tipo NPS': 'GV',
        'Generación': 'GU',
        'ANTIGÜEDAD': 'CI',
        'TIPO CONTRATO': 'CJ',
        'COMENTARIOS': 'H',
    }

    # Construir DataFrame de salida con columnas en orden y crear vacías si faltan
    out = pd.DataFrame()
    faltantes = []
    for tcol in TARGET_COLS:
        # 1) Preferir mapeo por letra si está definido
        letter = MANUAL_LETTERS.get(tcol)
        used_letter = False
        if letter:
            idx = col_letter_to_index(letter)
            if 0 <= idx < df_raw.shape[1]:
                series = df_raw.iloc[:, idx]
                # intentar alinear longitudes con df (asumiendo fila 0 = encabezados)
                if len(series) == len(df) + 1:
                    series = series.iloc[1:]
                series = series.reset_index(drop=True)
                # aplicar filtro de ID si existe
                if mask_keep is not None and len(series) > 0:
                    # cuando se creó mask_keep a partir de df previo, su longitud corresponde a la hoja sin encabezado
                    if len(series) == len(mask_keep):
                        series = series[mask_keep.values].reset_index(drop=True)
                if len(series) != len(df):
                    # fallback sin recorte
                    series = df_raw.iloc[:, idx].reset_index(drop=True)
                    # intentar alinear y filtrar de nuevo
                    if len(series) == len(df) + 1:
                        series = series.iloc[1:]
                    if mask_keep is not None and len(series) == len(mask_keep):
                        series = series[mask_keep.values].reset_index(drop=True)
                out[tcol] = series
                used_letter = True
        # 2) Si no hubo letra válida, usar mapeo por nombre
        if not used_letter:
            out[tcol] = ''
            faltantes.append(tcol)

    # Escribir/actualizar archivo de salida
    out_path = base_dir / OUTPUT_FILE
    mode = 'a' if out_path.exists() else 'w'
    writer_kwargs = dict(engine='openpyxl', mode=mode)
    if mode == 'a':
        writer_kwargs['if_sheet_exists'] = 'replace'
    try:
        with pd.ExcelWriter(out_path, **writer_kwargs) as writer:
            out.to_excel(writer, index=False, sheet_name=SHEET_OUTPUT)
        final_path = out_path
    except PermissionError:
        # Si el archivo de salida está bloqueado (Excel/OneDrive), escribir en temporal
        tmpdir = Path(tempfile.gettempdir())
        temp_out = tmpdir / f"tmp_{OUTPUT_FILE.replace('.xlsx','')}_{os.getpid()}.xlsx"
        with pd.ExcelWriter(temp_out, engine='openpyxl', mode='w') as writer:
            out.to_excel(writer, index=False, sheet_name=SHEET_OUTPUT)
        # Intentar reemplazar el archivo destino
        try:
            shutil.copy2(temp_out, out_path)
            final_path = out_path
        except PermissionError:
            # Dejar el archivo temporal y avisar
            final_path = temp_out
            print(f'El archivo de salida estaba en uso. Se creó una copia temporal en: {temp_out}', file=sys.stderr)

    print(f'Hoja "{SHEET_OUTPUT}" creada en {final_path.name} con {len(out)} filas.')
    if faltantes:
        print('Advertencia: No se encontraron en el origen y se dejaron vacías:', ', '.join(faltantes))
    # Mostrar resumen de mapeo fijo por letras
    print('\nResumen de mapeo (letras):')
    for k, v in MANUAL_LETTERS.items():
        print(f'  {k} <- LETTER:{v}')


if __name__ == '__main__':
    crear_hoja_comentarios()

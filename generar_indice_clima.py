from __future__ import annotations
from pathlib import Path
import argparse
import json
import os
import shutil
import sys
import tempfile
import time
from difflib import SequenceMatcher
import pandas as pd

INPUT_FILE = 'input.xlsx'
OUTPUT_FILE = 'salida.xlsx'
DEFAULT_SHEET = 'Datos A2'
SHEET_OUT = 'Indice de clima'
MAPPING_FILE = 'indice_clima_mapping.json'

# Utilidades de IO (mismo enfoque robusto que otros scripts)
def open_excel_safely(p: Path) -> pd.ExcelFile:
    try:
        return pd.ExcelFile(p, engine='openpyxl')
    except PermissionError:
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


# Conversión 1–4 a 10 con redondeos solicitados
def convert_4_to_10(avg4: float) -> float:
    # round to one decimal in 1-4 scale first
    avg4r = round(float(avg4), 1)
    # then convert to 10
    conv10 = (avg4r - 1.0) * (10.0 / 3.0)
    return round(conv10, 1)


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


def avg_from_letters(df: pd.DataFrame, df_raw: pd.DataFrame, letters: list[str]) -> tuple[float, float, float]:
    # retorna (real_prom, escala4_r, escala10_r)
    vals = []
    for lt in letters:
        s = get_aligned_series_by_letter(df, df_raw, lt)
        if s.empty:
            continue
        nums = pd.to_numeric(s, errors='coerce').dropna()
        if nums.empty:
            continue
        nums = nums.clip(lower=1.0, upper=4.0)
        vals.append(float(nums.mean()))
    if not vals:
        return (0.0, 0.0, 0.0)
    avg4 = sum(vals) / len(vals)
    avg4_r = round(avg4, 1)
    conv10_r = convert_4_to_10(avg4_r)
    return (avg4, avg4_r, conv10_r)


def load_mapping(base_dir: Path) -> dict:
    mp = base_dir / MAPPING_FILE
    if not mp.exists():
        raise SystemExit(f"No se encuentra {MAPPING_FILE} en {base_dir}")
    with mp.open('r' , encoding='utf-8') as f:
        data = json.load(f)
    return data


def linearize_rows(struct: dict) -> list[dict]:
    # Devuelve una lista lineal con filas de salida en el orden del mapping
    out = []
    label_to_row = {}

    def visit(node_label: str, nodes: dict[str, dict], parent: str | None = None):
        node = nodes[node_label]
        if node.get('kind') == 'group':
            out.append({'label': node_label, 'kind': 'group', 'columns': node.get('columns', [])})
            for ch in node.get('children', []):
                visit(ch, nodes, node_label)
        else:
            out.append({'label': node_label, 'kind': 'value', 'columns': node.get('columns', [])})

    # Construir índice de nodos por label
    nodes = {}
    for r in struct['rows']:
        nodes[r['label']] = r
    # Raíces son las que no aparecen como children de otras
    children = set()
    for r in struct['rows']:
        for ch in r.get('children', []):
            children.add(ch)
    roots = [r['label'] for r in struct['rows'] if r['label'] not in children]

    for root in roots:
        visit(root, nodes)
    return out


def generar():
    parser = argparse.ArgumentParser(description='Generar hoja "Indice de clima" (Real, Escala 4, Escala 10).')
    parser.add_argument('--sheet', help='Hoja de entrada (por defecto Datos A2)')
    parser.add_argument('--id', dest='ids', nargs='*', help='Filtrar por uno o varios ID (columna A)')
    parser.add_argument('--debug', action='store_true')
    parser.add_argument('--out', help='Archivo de salida (por defecto salida.xlsx en esta carpeta)')
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent

    mapping = load_mapping(base_dir)

    in_path = base_dir / INPUT_FILE
    if not in_path.exists():
        raise SystemExit(f'No se encuentra {INPUT_FILE} en {base_dir}')

    xls = open_excel_safely(in_path)
    hoja = args.sheet or mapping.get('defaults', {}).get('sheet', DEFAULT_SHEET)
    if hoja not in xls.sheet_names:
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

    # filtrar por id si corresponde
    if args.ids:
        idx_id = col_letter_to_index('A')
        if 0 <= idx_id < df_raw.shape[1]:
            series_id = df_raw.iloc[:, idx_id]
            if len(series_id) == len(df) + 1:
                series_id = series_id.iloc[1:]
            series_id = series_id.reset_index(drop=True).astype(str).str.strip()
            ids_norm = {str(x).strip() for x in args.ids}
            mask = series_id.isin(ids_norm)
            df = df[mask.values].reset_index(drop=True)
            df_raw = df_raw.iloc[1:, :].reset_index(drop=True)

    # Expandir estructura a filas lineales
    flat = linearize_rows(mapping)

    # Índice rápido de nodos por etiqueta
    nodes_index = {r['label']: r for r in mapping['rows']}

    # Memoización de valores calculados por nodo: label -> (avg4_real, escala4_r, escala10_r)
    computed: dict[str, tuple[float, float, float]] = {}

    def parse_letters(letters_any) -> list[str]:
        letters = letters_any or []
        if isinstance(letters, str):
            letters = [x.strip() for x in letters.split(',') if x.strip()]
        return letters

    def compute_node(label: str) -> tuple[float, float, float]:
        if label in computed:
            return computed[label]
        node = nodes_index[label]
        kind = node.get('kind')
        if kind == 'value':
            letters = parse_letters(node.get('columns'))
            avg_real, escala4, escala10 = avg_from_letters(df, df_raw, letters)
            computed[label] = (avg_real, escala4, escala10)
            return computed[label]
        # kind == 'group'
        children = node.get('children') or []
        group_agg = (node.get('groupAggregation') or mapping.get('defaults', {}).get('groupAggregation') or 'unweighted').lower()
        if children and group_agg == 'unweighted':
            child_vals_real = []
            child_vals_4 = []
            child_vals_10 = []
            for ch in children:
                c_avg_real, c_esc4, c_esc10 = compute_node(ch)
                # Considerar valores > 0.0 como válidos (0.0 indica sin datos por nuestro pipeline)
                if c_avg_real and c_avg_real > 0.0:
                    child_vals_real.append(float(c_avg_real))
                if c_esc4 and c_esc4 > 0.0:
                    child_vals_4.append(float(c_esc4))
                if c_esc10 and c_esc10 > 0.0:
                    child_vals_10.append(float(c_esc10))
            if child_vals_real or child_vals_4 or child_vals_10:
                avg4_real = sum(child_vals_real) / len(child_vals_real) if child_vals_real else 0.0
                escala4 = round(sum(child_vals_4) / len(child_vals_4), 1) if child_vals_4 else 0.0
                escala10 = round(sum(child_vals_10) / len(child_vals_10), 1) if child_vals_10 else 0.0
                computed[label] = (avg4_real, escala4, escala10)
                return computed[label]
        # Fallback: usar las columnas declaradas del grupo si no hay hijos válidos
        letters = parse_letters(node.get('columns'))
        avg_real, escala4, escala10 = avg_from_letters(df, df_raw, letters)
        computed[label] = (avg_real, escala4, escala10)
        return computed[label]

    # Construir dataframe de salida usando valores computados
    rows = []
    for node in flat:
        avg_real, escala4, escala10 = compute_node(node['label'])
        rows.append({
            'Eje': node['label'],
            'Real': round(avg_real, 6) if avg_real else '',
            'Escala 4': escala4 if escala4 else '',
            'Escala 10': escala10 if escala10 else ''
        })

    df_out = pd.DataFrame(rows, columns=['Eje', 'Real', 'Escala 4', 'Escala 10'])

    out_file = args.out or OUTPUT_FILE
    final_path = write_excel_safely(base_dir / out_file, {SHEET_OUT: df_out})
    print(f'Hoja "{SHEET_OUT}" escrita en {final_path}')


if __name__ == '__main__':
    generar()

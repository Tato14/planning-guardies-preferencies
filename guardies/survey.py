"""Importació de la taula de preferències (enquesta) a la pestanya Radiòlegs."""
from __future__ import annotations

import difflib
import io
import re

import openpyxl

from .io_entrada import pref_value
from .model import ROL_GUARDIES, Radiologist
from .periods import fold

ES_COLS = {
    'lun tarde': 'dl_tarda', 'lun noche': 'dl_nit', 'mar tarde': 'dt_tarda', 'mar noche': 'dt_nit',
    'mie tarde': 'dc_tarda', 'mie noche': 'dc_nit', 'jue tarde': 'dj_tarda', 'jue noche': 'dj_nit',
    'vie tarde': 'dv_tarda', 'vie noche': 'dv_nit', 'sab dia': 'ds_dia', 'sab noche': 'ds_nit',
    'dom dia': 'dg_dia', 'dom noche': 'dg_nit', 'fest dia': 'fest_dia', 'fest noche': 'fest_nit',
}


def read_survey(src) -> list:
    wb = openpyxl.load_workbook(io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src, data_only=True)
    ws = wb.worksheets[0]
    hdr = {fold(ws.cell(row=1, column=c).value): c for c in range(1, ws.max_column + 1)}

    def col(prefix):
        return next((c for h, c in hdr.items() if h.startswith(prefix)), None)
    c_name, c_min, c_max = col('radiolog'), col('min'), col('max')
    c_mati, c_dies = col('manana'), col('dias puntuales')
    rows = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=c_name).value
        if not isinstance(name, str) or not name.strip():
            continue
        def num(c):
            v = ws.cell(row=r, column=c).value if c else None
            try:
                return float(v) if v is not None and str(v).strip() != '' else None
            except ValueError:
                return None
        prefs = {}
        for h, key in ES_COLS.items():
            c = hdr.get(h)
            if c:
                pv = pref_value(ws.cell(row=r, column=c).value)
                if pv:
                    prefs[key] = pv
        rows.append({'name': name.strip(), 'min': num(c_min), 'max': num(c_max), 'prefs': prefs,
                     'mati': pref_value(ws.cell(row=r, column=c_mati).value) if c_mati else None,
                     'dies': str(ws.cell(row=r, column=c_dies).value or '').strip() if c_dies else ''})
    return rows


def _tokens(name):
    return [t for t in re.split(r'[\s\-]+', fold(name.split(',')[0])) if len(t) >= 3]


def _tok_eq(a, b):
    return a == b or difflib.SequenceMatcher(None, a, b).ratio() >= 0.82


def _first_initial(name):
    parts = name.split(',', 1)
    return fold(parts[1]).strip()[:1] if len(parts) == 2 and parts[1].strip() else ''


def match_status(survey_name, master_names) -> tuple:
    """(nom, estat) amb estat 'ok' | 'cap' | 'ambigu'.
    Cognoms: el primer ha de coincidir i també tots els que tinguin tots dos noms. Si tots dos porten
    nom de pila, la inicial ha de coincidir. Si queden dues opcions igual de bones, és ambigu."""
    ta, ia = _tokens(survey_name), _first_initial(survey_name)
    best = []
    for m in master_names:
        tb, ib = _tokens(m), _first_initial(m)
        if not ta or not tb or not _tok_eq(ta[0], tb[0]):
            continue
        if ia and ib and ia != ib:
            continue
        k = min(len(ta), len(tb))
        if all(_tok_eq(ta[i], tb[i]) for i in range(k)):
            score = sum(difflib.SequenceMatcher(None, ta[i], tb[i]).ratio() for i in range(k)) / k
            score += 0.5 * (min(len(ta), len(tb)) / max(len(ta), len(tb)))   # més cognoms coincidents, millor
            best.append((round(score, 6), m))
    if not best:
        return None, 'cap'
    best.sort(reverse=True)
    if len(best) > 1 and best[0][0] == best[1][0]:
        return None, 'ambigu'
    return best[0][1], 'ok'


def match(survey_name, master_names):
    return match_status(survey_name, master_names)[0]


def merge(rads: dict, rows: list) -> tuple:
    """Aplica l'enquesta a la llista mestra. Retorna (llista nova, informe)."""
    out = {n: r for n, r in rads.items()}
    report = {'casats': [], 'nous': [], 'sense_resposta': [], 'ambigus': [], 'duplicats': []}
    used = set()
    for row in rows:
        m, st = match_status(row['name'], list(rads))
        if st == 'ambigu':
            report['ambigus'].append(row['name'])      # no s'afegeix: cal decidir-ho a mà
            continue
        if m is not None and m in used:
            report['duplicats'].append((row['name'], m))
            continue
        if m is None:
            if row['name'] in out:
                report['duplicats'].append((row['name'], row['name']))
                continue
            r = Radiologist(name=row['name'], role=ROL_GUARDIES)
            out[row['name']] = r
            report['nous'].append(row['name'])
        else:
            r = out[m]
            used.add(m)
            report['casats'].append((row['name'], m))
        r.min_m, r.max_m = row['min'], (abs(row['max']) if row['max'] is not None else None)
        if r.max_m is None and r.min_m:
            r.max_m = r.min_m
        r.prefs = dict(row['prefs'])
        r.mati = row['mati']
        r.dies_puntuals = row['dies']
    report['sense_resposta'] = sorted(n for n in rads if n not in used)
    return out, report

"""Lectura del planning del mes anterior (el que es va publicar, amb els canvis manuals).

Se'n treu:
  - quantes guàrdies ha fet cadascú, per bossa (laborables / cap de setmana i festius);
  - la data de l'última guàrdia de cadascú (per al descans a l'inici del període);
  - si existeix, la pestanya «Historic» que va escriure l'app: deute i quota de cada
    persona a l'inici d'aquell període, per encadenar l'equitat mes a mes.
Funciona també amb els plannings de l'app anterior (v3), que no tenen pestanya Historic.
"""
from __future__ import annotations

import datetime as dt
import io
import re
from dataclasses import dataclass, field

import openpyxl

from .model import POOL_C, POOL_L
from .periods import as_date, fold

SHIFT_ROWS = {'8-16': (4, 9), '16-20': (11, 16), '20-08': (18, 23)}
HIST_SHEET = 'Historic'
DEUTE_MAX = 2.0


@dataclass
class History:
    start: dt.date | None = None
    end: dt.date | None = None
    code: str = ''
    guards: dict = field(default_factory=dict)        # nom -> {'L': n, 'C': n}
    guard_dates: dict = field(default_factory=dict)   # nom -> [dates]
    debt_start: dict = field(default_factory=dict)    # nom -> {'L': x, 'C': x}  (de la pestanya Historic)
    share: dict = field(default_factory=dict)         # nom -> {'L': x, 'C': x}
    hist_last: dict = field(default_factory=dict)     # nom -> data de l'última guàrdia segons Historic
    has_hist_sheet: bool = False
    unmatched: list = field(default_factory=list)
    warnings: list = field(default_factory=list)

    def had_guard(self, name) -> bool:
        g = self.guards.get(name, {})
        return (g.get(POOL_L, 0) + g.get(POOL_C, 0)) > 0

    def last_guard(self, name):
        ds = list(self.guard_dates.get(name, []))
        if self.hist_last.get(name):
            ds.append(self.hist_last[name])
        return max(ds) if ds else None

    def debt_now(self, name, pool) -> float:
        """Deute acumulat a l'inici del període nou: (deute inicial + quota − fetes)."""
        if name not in self.share:
            return 0.0
        v = (self.debt_start.get(name, {}).get(pool, 0.0) + self.share[name].get(pool, 0.0)
             - self.guards.get(name, {}).get(pool, 0))
        return max(-DEUTE_MAX, min(DEUTE_MAX, v))


def _match_name(raw, names_fold):
    """Nom exacte; si no, només una coincidència inequívoca (si n'hi ha més d'una, no es casa)."""
    f = fold(raw)
    if f in names_fold:
        return names_fold[f]
    pref = {v for k, v in names_fold.items() if k.startswith(f) or f.startswith(k)}
    if len(pref) == 1:
        return pref.pop()
    if len(pref) > 1:
        return None
    sur = f.split(',')[0].strip()
    cands = {v for k, v in names_fold.items() if k.split(',')[0].strip() == sur}
    return cands.pop() if len(cands) == 1 else None


LINE_RE = re.compile(r'^(Grup\s*[12]|Grupo\s*[12]|N\s+i\s+B|NB|N|B)\s*[:.]\s*(.+)$', re.IGNORECASE)


def _is_fixed_cell(d, label, fest, cfg):
    """Torns fixos segons la data i la franja (no segons el text, que es pot haver editat a mà)."""
    if fest or d.weekday() >= 5:
        return False                                   # caps de setmana i festius: tot és guàrdia
    if label == '8-16':
        return True                                    # matins entre setmana
    if label == '16-20':
        if d.weekday() in (0, 1):
            return True                                # tardes de dilluns i dimarts
        if cfg is not None:
            if any(k[0] == d.weekday() and k[1] == '16-20' for k in cfg.fix_slots):
                return True
            if d.weekday() == 2 and d.day <= 7 and cfg.first_wed is not None:
                return True
        elif d.weekday() == 2 and d.day <= 7:
            return True                                # per defecte, el 1r dimecres és fix
    return False


def parse_planning(src, names, cfg=None) -> History:
    """names: noms de la llista mestra actual (per casar-hi els del planning).
    cfg: Configuració actual, per saber quines tardes són fixes."""
    wb = src if isinstance(src, openpyxl.Workbook) else openpyxl.load_workbook(
        io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src, data_only=True)
    names_fold = {fold(n): n for n in names}
    h = History()
    sheets = [s for s in wb.sheetnames if 'setmana' in s.lower()]
    if not sheets:
        raise ValueError("El planning anterior no té cap pestanya «Setmana».")
    all_dates, unmatched, seen_dates, dup_dates = [], set(), set(), set()
    for sn in sheets:
        ws = wb[sn]
        for col in range(2, 9):
            d = as_date(ws.cell(row=3, column=col).value)
            if d is None:
                continue
            first = str(ws.cell(row=4, column=col).value or '')
            if first.startswith('(fora del període'):
                continue
            if d in seen_dates:
                dup_dates.add(d)
                continue
            seen_dates.add(d)
            all_dates.append(d)
            fest_cols = any('FESTIU' in str(ws.cell(row=r, column=col).value or '').upper() for r in range(4, 24))
            for label, (r0, r1) in SHIFT_ROWS.items():
                for r in range(r0, r1 + 1):
                    v = ws.cell(row=r, column=col).value
                    if not isinstance(v, str):
                        continue
                    s = v.strip()
                    if not s or s.startswith(('(', '[', 'Dissabte', 'Diumenge', 'FESTIU')):
                        continue
                    m = LINE_RE.match(s)
                    if not m:
                        continue
                    tag = fold(m.group(1)).replace(' ', '')
                    name_raw = re.sub(r'\s*\([^)]*\)\s*$', '', m.group(2)).strip()
                    if not name_raw or name_raw.startswith('[') or name_raw in ('None', '—'):
                        continue
                    weekend_or_fest = d.weekday() >= 5 or fest_cols
                    if tag in ('n', 'b') or _is_fixed_cell(d, label, fest_cols, cfg):
                        continue
                    if label == '16-20' and weekend_or_fest:
                        continue
                    nm = _match_name(name_raw, names_fold)
                    if nm is None:
                        unmatched.add(name_raw)
                        continue
                    pool = POOL_C if weekend_or_fest else POOL_L
                    g = h.guards.setdefault(nm, {POOL_L: 0, POOL_C: 0})
                    g[pool] += 1
                    h.guard_dates.setdefault(nm, []).append(d)
    if all_dates:
        h.start, h.end = min(all_dates), max(all_dates)
    h.unmatched = sorted(unmatched)
    if dup_dates:
        h.warnings.append("El planning anterior té dates repetides en dues pestanyes (p. ex. una còpia de setmana); "
                          "s'ha llegit només la primera: " + ', '.join(f"{d:%d/%m}" for d in sorted(dup_dates)))
    if HIST_SHEET in wb.sheetnames:
        ws = wb[HIST_SHEET]
        h.has_hist_sheet = True
        hdr_row = None
        for r in range(1, 12):
            if fold(ws.cell(row=r, column=1).value) == 'nom':
                hdr_row = r
                break
        if fold(ws.cell(row=2, column=1).value) == 'periode':
            h.code = str(ws.cell(row=2, column=2).value or '')
        if hdr_row:
            hdr = {fold(ws.cell(row=hdr_row, column=c).value): c for c in range(1, ws.max_column + 1)}
            def num(r, key):
                c = hdr.get(key)
                try:
                    return float(ws.cell(row=r, column=c).value or 0) if c else 0.0
                except (TypeError, ValueError):
                    return 0.0
            for r in range(hdr_row + 1, ws.max_row + 1):
                raw = ws.cell(row=r, column=1).value
                if not isinstance(raw, str) or not raw.strip():
                    continue
                nm = _match_name(raw.strip(), names_fold)
                if nm is None:
                    continue
                c_last = hdr.get('ultima guardia')
                if c_last:
                    d_last = as_date(ws.cell(row=r, column=c_last).value)
                    if d_last:
                        h.hist_last[nm] = d_last
                if fold(ws.cell(row=r, column=2).value) != 'p':
                    continue
                h.debt_start[nm] = {POOL_L: num(r, 'deute l (inici)'), POOL_C: num(r, 'deute c (inici)')}
                h.share[nm] = {POOL_L: num(r, 'quota l'), POOL_C: num(r, 'quota c')}
    return h

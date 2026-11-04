"""Lectura del fitxer d'entrada del mes (Configuració + Radiòlegs + Vacances)."""
from __future__ import annotations

import calendar
import datetime as dt
import io
import re

import openpyxl

from .model import (CODI_INACTIU, CODI_REINCORPORACIO, CODIS_AMB_REINCORPORACIO, CODIS_BLOQUEJANTS,
                    CODIS_FULL, DIES_REINCORPORACIO, MOTIU_TEXT, PREF_HEADERS, PREF_KEYS, PREF_NO,
                    PREF_SI, ROL_ALIES, ROL_CS_DIA, ROL_DG_NIT, ROL_GUARDIES, ROLS, Config, Entrada,
                    Period, Radiologist)
from .periods import as_date, fmt_dia, fold, month_label, parse_month

DOW_MAP = {'dilluns': 0, 'dimarts': 1, 'dimecres': 2, 'dijous': 3, 'divendres': 4,
           'dissabte': 5, 'diumenge': 6}
ORD_MAP = {'1r': 1, '1': 1, 'primer': 1, '2n': 2, '2': 2, 'segon': 2, '3r': 3, '3': 3, 'tercer': 3,
           '4t': 4, '4': 4, 'quart': 4, '5e': 5, '5': 5, 'cinque': 5}


def _open(src):
    if isinstance(src, openpyxl.Workbook):
        return src
    if isinstance(src, (bytes, bytearray)):
        return openpyxl.load_workbook(io.BytesIO(src), data_only=True)
    return openpyxl.load_workbook(src, data_only=True)


class _Bad(float):
    pass


def _num(v):
    if v is None or (isinstance(v, str) and not v.strip()):
        return None
    try:
        return float(str(v).replace(',', '.').strip())
    except ValueError:
        return _Bad('nan')


def pref_value(v):
    f = fold(v)
    if not f:
        return None
    if f in ('si', 's', 'x', 'p', 'yes', 'sí', 'preferit', 'preferida'):
        return PREF_SI
    if f in ('no', 'n'):
        return PREF_NO
    return None


def es_actiu(v) -> bool:
    f = fold(v)
    return f not in ('no', 'n', 'fals', 'false', '0')


def es_fila_llegenda(name: str) -> bool:
    f = fold(name).rstrip(':').strip()
    return f.startswith('llegenda') or f in {fold(r) for r in list(ROL_ALIES) + list(ROLS)}


def normalitza_rol(v) -> str:
    f = fold(v)
    if not f:
        return ROL_GUARDIES
    for k, r in ROL_ALIES.items():
        if fold(k) == f:
            return r
    return f   # desconegut: es valida més avall


# ----------------------------------------------------------------------
# Radiòlegs
# ----------------------------------------------------------------------
def _find_header_row(ws, first_prefix='nom', max_row=8):
    for r in range(1, max_row + 1):
        if fold(ws.cell(row=r, column=1).value).startswith(first_prefix):
            return r
    return 3


def read_radiologists(ws, warnings) -> dict:
    hr = _find_header_row(ws)
    cols = {}
    for c in range(1, ws.max_column + 1):
        h = fold(ws.cell(row=hr, column=c).value)
        if h:
            cols.setdefault(h, c)

    def col(*prefixes):
        for p in prefixes:
            for h, c in cols.items():
                if h.startswith(p):
                    return c
        return None

    c_rol, c_mail, c_act = col('rol'), col('correu', 'email', 'notes'), col('actiu')
    c_min, c_max = col('min'), col('max')
    c_mati, c_dies, c_obs = col('mati', 'manana'), col('dies puntuals', 'dias puntuales'), col('observacions')
    c_pref = {k: col(fold(h)) for k, h in PREF_HEADERS.items()}
    if c_act is None and ws.max_column >= 5:
        c_act = 5   # format v2-v3: Actiu a la columna E

    out = {}
    for r in range(hr + 1, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        if not isinstance(name, str) or not name.strip():
            continue
        name = name.strip()
        if es_fila_llegenda(name):
            continue
        rol = normalitza_rol(ws.cell(row=r, column=c_rol).value if c_rol else None)
        if rol not in ROLS:
            warnings.append(f"Radiòlegs, fila {r}: rol «{ws.cell(row=r, column=c_rol).value}» desconegut per a {name}; "
                            "es tracta com a «guàrdies».")
            rol = ROL_GUARDIES
        mail = ws.cell(row=r, column=c_mail).value if c_mail else ''
        mn = _num(ws.cell(row=r, column=c_min).value) if c_min else None
        mx = _num(ws.cell(row=r, column=c_max).value) if c_max else None
        for lbl, val, col_ in (('mínim', mn, c_min), ('màxim', mx, c_max)):
            if isinstance(val, _Bad):
                warnings.append(f"Radiòlegs: el {lbl} de {name} («{ws.cell(row=r, column=col_).value}») no és un número; "
                                "es deixa buit.")
        mn = None if isinstance(mn, _Bad) else mn
        mx = None if isinstance(mx, _Bad) else mx
        if mx is not None and mx < 0:
            warnings.append(f"Radiòlegs: {name} té màxim {mx:g}; es pren {abs(mx):g}.")
            mx = abs(mx)
        if mx is None and mn is not None:
            mx = mn
        if mx is not None and mx == 0 and rol not in ('fix',):
            warnings.append(f"Radiòlegs: {name} té màxim 0; es tracta com qui no vol fer guàrdies (1 cada 2 mesos).")
        if mx is not None and 0 < mx < 1:
            warnings.append(f"Radiòlegs: {name} té un màxim inferior a 1 ({mx:g}); se li assegura 1 guàrdia cada 2 mesos.")
        if mn is not None and mx is not None and mn > mx:
            warnings.append(f"Radiòlegs: {name} té mínim ({mn:g}) més gran que el màxim ({mx:g}).")
        prefs = {}
        for k, c in c_pref.items():
            if c:
                pv = pref_value(ws.cell(row=r, column=c).value)
                if pv:
                    prefs[k] = pv
        if name in out:
            warnings.append(f"Radiòlegs: {name} hi surt dues vegades; es fa servir la primera.")
            continue
        out[name] = Radiologist(
            name=name, role=rol, email=str(mail or '').strip() if mail and '@' in str(mail) else '',
            active=es_actiu(ws.cell(row=r, column=c_act).value) if c_act else True,
            min_m=mn, max_m=mx, prefs=prefs,
            mati=pref_value(ws.cell(row=r, column=c_mati).value) if c_mati else None,
            dies_puntuals=str(ws.cell(row=r, column=c_dies).value or '').strip() if c_dies else '',
            observacions=str(ws.cell(row=r, column=c_obs).value or '').strip() if c_obs else '',
            row=r)
    return out


# ----------------------------------------------------------------------
# Configuració
# ----------------------------------------------------------------------
def read_config(ws, warnings) -> Config:
    cfg = Config()
    old_friday = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(row=r, column=1).value
        if not isinstance(a, str) or not a.strip():
            continue
        fa = fold(a)
        b = fold(ws.cell(row=r, column=2).value)
        c = str(ws.cell(row=r, column=3).value or '').strip()
        name = str(ws.cell(row=r, column=4).value or '').strip()
        m = re.match(r'^divendres\s*[-–]\s*(\S+)', fa)
        if m and m.group(1) in ORD_MAP:
            if name:
                cfg.friday[ORD_MAP[m.group(1)]] = name
            continue
        if 'dimecres-1r' in fa or '1r-dimecres' in fa or 'primer dimecres' in fa:
            cfg.first_wed = name          # '' = fila present sense titular (sortirà pendent)
            continue
        if fa in DOW_MAP and b in ('8-16', '16-20') and c:
            dow = DOW_MAP[fa]
            if dow == 4 and b == '8-16' and fold(c) in ('n i b', 'nib'):
                old_friday = name          # format antic: un sol titular per a tots els divendres
                continue
            cfg.fix_slots[(dow, b, 'N i B' if fold(c) in ('n i b', 'nib') else c.upper())] = name
            continue
        if fa == 'sun-nit-only' and name:
            cfg.sun_nit_only.append(name)
        elif fa == 'weekend-day-only' and name:
            cfg.weekend_day_only.append(name)
        elif fa.startswith('descans'):
            v = _num(ws.cell(row=r, column=2).value)
            if v is not None and not isinstance(v, _Bad):
                cfg.rest_days = max(0, int(v))
        elif fa.startswith('minim per a tothom') or fa.startswith('1 guardia cada'):
            v = _num(ws.cell(row=r, column=2).value)
            if v is not None and not isinstance(v, _Bad) and v >= 1:
                cfg.min_every_months = int(v)
    if old_friday:
        for k in range(1, 6):
            cfg.friday.setdefault(k, old_friday)
    return cfg


# ----------------------------------------------------------------------
# Vacances (període i absències)
# ----------------------------------------------------------------------
def _find_label_row(ws, prefix, max_row=15):
    pf = fold(prefix)
    for r in range(1, min(ws.max_row, max_row) + 1):
        if fold(ws.cell(row=r, column=1).value).startswith(pf):
            return r
    return None


def _value_right_of(ws, row, prefix, max_col=60):
    pf = fold(prefix)
    last = min(ws.max_column, max_col)
    for c in range(1, last + 1):
        if fold(ws.cell(row=row, column=c).value).startswith(pf):
            for c2 in range(c + 1, last + 1):
                v = ws.cell(row=row, column=c2).value
                if v not in (None, ''):
                    return v
            return None
    return None


def read_vacances(ws, names: set, warnings) -> tuple:
    header_row = _find_label_row(ws, 'radiol') or 3
    date_cols, int_cols = {}, {}
    for c in range(2, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        d = as_date(v) if not isinstance(v, str) else None
        if d is not None:
            date_cols[c] = d
        elif isinstance(v, (int, float)) and not isinstance(v, bool) and float(v).is_integer() and 1 <= int(v) <= 31:
            int_cols[c] = int(v)

    if date_cols:
        raw_label = _value_right_of(ws, 1, 'mes del planning') or ws['B1'].value
        all_dates = sorted(set(date_cols.values()))
        start = as_date(_value_right_of(ws, 2, 'periode planificat')) or all_dates[0]
        end = as_date(_value_right_of(ws, 2, 'fins al')) or all_dates[-1]
        if end < start:
            raise ValueError(f"El període està invertit: {fmt_dia(start, True)} → {fmt_dia(end, True)}")
        planned = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
        falten = [d for d in planned if d not in set(all_dates)]
        if falten:
            raise ValueError("A la capçalera de Vacances hi falten dates del període: "
                             + ', '.join(fmt_dia(d) for d in falten[:10]))
        try:
            year, month = parse_month(raw_label)
        except ValueError:
            cnt = {}
            for d in planned:
                cnt[(d.year, d.month)] = cnt.get((d.year, d.month), 0) + 1
            year, month = max(cnt.items(), key=lambda kv: kv[1])[0]
            warnings.append(f"No s'ha entès el mes «{raw_label}»; s'ha pres {month_label(year, month)}.")
        fest_row = _find_label_row(ws, 'festiu')
        festius = set()
        if fest_row:
            for c, d in date_cols.items():
                if start <= d <= end and str(ws.cell(row=fest_row, column=c).value or '').strip():
                    festius.add(d)
        col_date = {c: d for c, d in date_cols.items() if d <= end}
        if start.weekday() != 0:
            warnings.append(f"El període comença en {fmt_dia(start, True)}, no en dilluns (dia pont).")
        if (start.year, start.month) < (year, month):
            warnings.append(f"ATENCIÓ: el període comença el {fmt_dia(start, True)}, abans del mes del planning "
                            f"({month_label(year, month)}). Comproveu la data de la fila 2.")
        first_name_row = max(header_row, fest_row or 0) + 1
    else:
        year, month = parse_month(ws['B1'].value)
        n = calendar.monthrange(year, month)[1]
        planned = [dt.date(year, month, d) for d in range(1, n + 1)]
        festius = {dt.date(year, month, int(p)) for p in re.split(r'[,;\s]+', str(ws['F1'].value or ''))
                   if p.isdigit() and 1 <= int(p) <= n}
        col_date = {c: dt.date(year, month, d) for c, d in int_cols.items() if d <= n}
        warnings.append("Fitxer de Vacances en format antic (mes natural): el planning no anirà per setmanes completes.")
        first_name_row = header_row + 1

    period = Period(year=year, month=month, dates=planned, festius=festius)
    absences, seen = {}, set()
    for r in range(first_name_row, ws.max_row + 1):
        raw = ws.cell(row=r, column=1).value
        name = raw.strip() if isinstance(raw, str) else ''
        cons, bad = {}, []
        for c, d in col_date.items():
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().upper() in CODIS_FULL:
                cons[d] = v.strip().upper()
            elif v not in (None, '') and str(v).strip():
                bad.append(f"{d:%d/%m}={str(v).strip()}")
        if bad and name:
            warnings.append(f"Vacances: {name} té valors no reconeguts (no compten com a absència): "
                            + ', '.join(bad[:6]) + (' …' if len(bad) > 6 else '') + ". Feu servir V, B, C, G o X.")
        if not name:
            if cons:
                warnings.append(f"Vacances, fila {r}: hi ha marques però no hi ha nom (s'ignoren).")
            continue
        if name not in names:
            if cons:
                warnings.append(f"Vacances, fila {r}: «{name}» no és a Radiòlegs; les marques s'ignoren.")
            continue
        seen.add(name)
        absences.setdefault(name, {}).update(cons)
    return period, absences, seen


def apply_derived_codes(absences: dict, inactius: set, period: Period) -> None:
    for name in inactius:
        cons = absences.setdefault(name, {})
        for d in period.dates:
            cons[d] = CODI_INACTIU
    extra = max(0, DIES_REINCORPORACIO - 1)
    for name, cons in absences.items():
        if name in inactius or not cons:
            continue
        fins = [d for d, code in cons.items() if code in CODIS_AMB_REINCORPORACIO
                and cons.get(d + dt.timedelta(days=1)) not in CODIS_AMB_REINCORPORACIO]
        for d in fins:
            for k in range(1, extra + 1):
                nd = d + dt.timedelta(days=k)
                if period.contains(nd) and cons.get(nd) in (None, 'X'):
                    cons[nd] = CODI_REINCORPORACIO


def can_work(name: str, d: dt.date, absences: dict) -> bool:
    cons = absences.get(name, {})
    if cons.get(d) in CODIS_BLOQUEJANTS:
        return False
    one = dt.timedelta(days=1)
    return not (cons.get(d - one) == 'G' or cons.get(d + one) == 'G')


def motiu(name: str, d: dt.date, absences: dict) -> str:
    cons = absences.get(name, {})
    code = cons.get(d)
    if code in CODIS_BLOQUEJANTS:
        return MOTIU_TEXT.get(code, code)
    one = dt.timedelta(days=1)
    if cons.get(d - one) == 'G' or cons.get(d + one) == 'G':
        return 'guàrdia externa adjacent'
    return 'no disponible'


# ----------------------------------------------------------------------
# Entrada completa
# ----------------------------------------------------------------------
def load_entrada(src) -> Entrada:
    wb = _open(src)
    warnings = []
    for req in ('Radiòlegs', 'Configuració', 'Vacances'):
        if req not in wb.sheetnames:
            raise ValueError(f"Falta la pestanya «{req}» al fitxer d'entrada.")
    rads = read_radiologists(wb['Radiòlegs'], warnings)
    cfg = read_config(wb['Configuració'], warnings)
    # Perfils especials de Configuració (compatibilitat): s'apliquen com a rol
    for name in cfg.sun_nit_only:
        if name in rads and rads[name].role == ROL_GUARDIES:
            rads[name].role = ROL_DG_NIT
    for name in cfg.weekend_day_only:
        if name in rads and rads[name].role == ROL_GUARDIES:
            rads[name].role = ROL_CS_DIA
    period, absences, seen = read_vacances(wb['Vacances'], set(rads), warnings)
    inactius = {n for n, r in rads.items() if not r.active}
    for n in rads:
        absences.setdefault(n, {})
    apply_derived_codes(absences, inactius, period)

    # Coherència
    for key, n in list(cfg.fix_slots.items()) + [(('1r dimecres',), cfg.first_wed or '')] + \
            [((f'divendres {k}',), v) for k, v in cfg.friday.items()]:
        if n and n not in rads:
            warnings.append(f"Configuració: «{n}» no és a la pestanya Radiòlegs.")
    buits = sorted({f"{['dl', 'dt', 'dc', 'dj', 'dv'][k[0]]} {k[1]} {k[2]}" for k, n in cfg.fix_slots.items() if not n})
    if buits:
        warnings.append("Configuració: torns fixos sense titular (sortiran pendents): " + ', '.join(buits))
    if cfg.first_wed == '':
        warnings.append("Configuració: el 1r dimecres no té titular (sortirà pendent).")
    falten = [k for k in range(1, 6) if not cfg.friday.get(k)]
    if falten:
        warnings.append("Configuració: falten titulars dels divendres al matí ("
                        + ', '.join(f"{k}r" if k in (1, 3) else f"{k}n" if k == 2 else f"{k}t" if k == 4 else '5è'
                                    for k in falten) + "). Sortiran com a pendents.")
    sense_fila = sorted((n for n, r in rads.items() if r.active and n not in seen), key=fold)
    if sense_fila:
        warnings.append("Sense fila a Vacances (no se'ls poden marcar absències): " + ', '.join(sense_fila))
    return Entrada(period=period, radiologists=rads, config=cfg, absences=absences, warnings=warnings)

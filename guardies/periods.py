"""Períodes de setmanes completes i festius.

Cada planning mensual està format per setmanes completes (dilluns → diumenge).
Una setmana pertany al mes en què cau el seu DILLUNS:
    Desembre 2026 = dl 07/12/2026 → dg 03/01/2027
"""
from __future__ import annotations

import datetime as dt
import re
import unicodedata

MESOS_CA = ['gener', 'febrer', 'març', 'abril', 'maig', 'juny', 'juliol',
            'agost', 'setembre', 'octubre', 'novembre', 'desembre']
MESOS_CA_ABREV = ['gen', 'febr', 'març', 'abr', 'maig', 'juny', 'jul', 'ag', 'set', 'oct', 'nov', 'des']
DIES_CA = ['dl', 'dt', 'dc', 'dj', 'dv', 'ds', 'dg']
DIES_CA_LLARG = ['Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres', 'Dissabte', 'Diumenge']
ORDINALS = {1: '1r', 2: '2n', 3: '3r', 4: '4t', 5: '5è'}


def fold(s) -> str:
    if s is None:
        return ''
    return unicodedata.normalize('NFKD', str(s)).encode('ascii', 'ignore').decode().lower().strip()


def as_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        m = re.match(r'^\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*$', v)
        if m:
            try:
                return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
        try:
            return dt.date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


def first_monday(year: int, month: int) -> dt.date:
    d = dt.date(year, month, 1)
    return d + dt.timedelta(days=(7 - d.weekday()) % 7)


def next_month(year: int, month: int) -> tuple:
    return (year + 1, 1) if month == 12 else (year, month + 1)


def prev_month(year: int, month: int) -> tuple:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def period_for_month(year: int, month: int) -> tuple:
    start = first_monday(year, month)
    ny, nm = next_month(year, month)
    return start, first_monday(ny, nm) - dt.timedelta(days=1)


def period_month_of(d: dt.date) -> tuple:
    monday = d - dt.timedelta(days=d.weekday())
    return monday.year, monday.month


def month_label(year: int, month: int) -> str:
    return f"{MESOS_CA[month - 1].capitalize()} {year}"


def parse_month(s) -> tuple:
    if isinstance(s, (dt.datetime, dt.date)):
        return s.year, s.month
    s = str(s or '').strip()
    noms = {'gener': 1, 'febrer': 2, 'marc': 3, 'abril': 4, 'maig': 5, 'juny': 6, 'juliol': 7,
            'agost': 8, 'setembre': 9, 'octubre': 10, 'novembre': 11, 'desembre': 12,
            'enero': 1, 'febrero': 2, 'marzo': 3, 'mayo': 5, 'junio': 6, 'julio': 7,
            'septiembre': 9, 'noviembre': 11, 'diciembre': 12}
    f = fold(s)
    year = month = None
    m = re.search(r'(20\d{2})', s)
    if m:
        year = int(m.group(1))
    for k, v in noms.items():
        if k in f:
            month = v
            break
    if month is None:
        m = re.search(r'\b(20\d{2})-(0?[1-9]|1[0-2])\b', s) or re.search(r'\b(0?[1-9]|1[0-2])[/\-](20\d{2})\b', s)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            year, month = (a, b) if a > 12 else (b, a)
    if year is None or month is None:
        raise ValueError(f"No s'ha pogut interpretar el mes: «{s}»")
    return year, month


def fmt_dia(d: dt.date, with_year: bool = False) -> str:
    s = f"{DIES_CA[d.weekday()]} {d.day:02d}/{d.month:02d}"
    return f"{s}/{d.year}" if with_year else s


def week_sheet_name(monday: dt.date) -> str:
    """'Setmana 49 (30 nov-06 des)'. Sempre conté 'Setmana' (l'Office Script de M365 la busca)."""
    sunday = monday + dt.timedelta(days=6)
    a = f"{monday.day:02d} {MESOS_CA_ABREV[monday.month - 1]}"
    b = f"{sunday.day:02d} {MESOS_CA_ABREV[sunday.month - 1]}"
    return f"Setmana {monday.isocalendar()[1]:02d} ({a}-{b})"


def weekday_ordinal_in_month(d: dt.date) -> int:
    """1 per al primer dilluns/dimarts/... del mes natural, 2 per al segon, etc."""
    return (d.day - 1) // 7 + 1


# ------------------------------------------------------------------
# Festius de Catalunya
# ------------------------------------------------------------------
# Calendaris oficials verificats (Departament d'Empresa i Treball):
#   2026: Ordre EMT/66/2025 · 2027: Ordre EMT/52/2026 (DOGC 9637, 1/4/2026)
FESTIUS_OFICIALS = {
    2026: [(1, 1), (1, 6), (4, 3), (4, 6), (5, 1), (6, 24), (8, 15), (9, 11), (10, 12), (12, 8), (12, 25), (12, 26)],
    2027: [(1, 1), (1, 6), (3, 26), (3, 29), (5, 1), (6, 24), (9, 11), (10, 12), (11, 1), (12, 6), (12, 8), (12, 25)],
}


def _pasqua(year: int) -> dt.date:
    a = year % 19
    b, c = divmod(year, 100)
    d, e = divmod(b, 4)
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = divmod(c, 4)
    l_ = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l_) // 451
    month = (h + l_ - 7 * m + 114) // 31
    day = ((h + l_ - 7 * m + 114) % 31) + 1
    return dt.date(year, month, day)


def festius_tradicionals(year: int) -> list:
    """Els 14 festius habituals de Catalunya (inclosos els que cauen en diumenge)."""
    p = _pasqua(year)
    fixos = [(1, 1), (1, 6), (5, 1), (6, 24), (8, 15), (9, 11), (10, 12), (11, 1), (12, 6), (12, 8), (12, 25), (12, 26)]
    out = [dt.date(year, m, d) for m, d in fixos]
    out += [p - dt.timedelta(days=2), p + dt.timedelta(days=1)]
    return sorted(out)


def festius_any(year: int) -> tuple:
    """(festius, verificat): calendari oficial si està publicat; si no, els 14 habituals."""
    if year in FESTIUS_OFICIALS:
        return sorted(dt.date(year, m, d) for m, d in FESTIUS_OFICIALS[year]), True
    return festius_tradicionals(year), False


def festius_entre(start: dt.date, end: dt.date) -> tuple:
    out, ok = [], True
    for y in range(start.year, end.year + 1):
        fy, verified = festius_any(y)
        f = [d for d in fy if start <= d <= end]
        out += f
        if f and not verified:
            ok = False
    return sorted(out), ok

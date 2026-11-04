import collections
import datetime as dt
import io
import random

import openpyxl
import pytest

from guardies import demo, engine, templates, writer
from guardies.history import parse_planning
from guardies.io_entrada import can_work, load_entrada
from guardies.model import K_NOU, K_PROP, K_SENSE, POOL_C, POOL_L, ROL_CS_DIA, ROL_DG_NIT, Period, Slot
from guardies.periods import (FESTIUS_OFICIALS, as_date, festius_tradicionals, next_month, period_for_month)
import officescript_port as osp

Y, M = 2027, 3      # període de 5 setmanes, amb Divendres Sant i Dilluns de Pasqua


def _bytes(wb):
    return templates.to_bytes(wb)


def _with_absences(wb_bytes, seed, n_vac=8):
    rnd = random.Random(seed)
    wb = openpyxl.load_workbook(io.BytesIO(wb_bytes))
    ws = wb['Vacances']
    cols = {ws.cell(row=4, column=c).value.date(): c for c in range(2, ws.max_column + 1)
            if isinstance(ws.cell(row=4, column=c).value, dt.datetime)}
    dates = sorted(cols)
    rows = [r for r in range(6, ws.max_row + 1) if ws.cell(row=r, column=1).value]
    for r in rnd.sample(rows, n_vac):
        i = rnd.randrange(0, len(dates) - 7)
        for d in dates[i:i + 7]:
            ws.cell(row=r, column=cols[d]).value = rnd.choice('VVVB')
    for r in rnd.sample(rows, 3):
        ws.cell(row=r, column=cols[rnd.choice(dates[3:])]).value = 'G'
    return _bytes(wb)


@pytest.fixture(scope='module')
def solved():
    inp = _with_absences(_bytes(demo.demo_input(Y, M)), 1)
    e = load_entrada(inp)
    return e, engine.solve(e, None)


def test_periodes_continus_i_festius():
    y, m = 2026, 11
    s0, e0 = period_for_month(y, m)
    for _ in range(36):
        y, m = next_month(y, m)
        s1, e1 = period_for_month(y, m)
        assert s1 == e0 + dt.timedelta(days=1) and s1.weekday() == 0 and e1.weekday() == 6
        e0 = e1
    for year, off in FESTIUS_OFICIALS.items():
        t, o = set(festius_tradicionals(year)), {dt.date(year, mm, dd) for mm, dd in off}
        assert o <= t and all(d.weekday() == 6 for d in t - o)


def test_plantilla_buida_sense_radiolegs_falsos():
    e = load_entrada(_bytes(templates.blank_input(Y, M)))
    assert e.radiologists == {}


def test_regles_dures(solved):
    e, res = solved
    per = e.period
    assert res.status == 'OPTIMAL' and not res.uncovered
    perday = collections.Counter((n, s.date) for s, n in res.assignments.items())
    assert max(perday.values()) == 1
    fixed_days = collections.defaultdict(set)
    for fa in res.fixed:
        if fa.status == 'ok':
            fixed_days[fa.name].add(fa.date)
    for s, n in res.assignments.items():
        r = e.radiologists[n]
        assert can_work(n, s.date, e.absences), (n, s)
        assert engine.pref_of(r, s, per) != 'N', (n, s)
        assert s.date not in fixed_days[n], (n, s)
        if r.role == ROL_DG_NIT:
            assert s.dow == 6 and s.shift == 'nit'
        if r.role == ROL_CS_DIA:
            assert s.dow >= 5 and s.shift == 'dia'
        if r.klass == K_NOU:
            assert s.dow == 2 and s.shift == '16-20' and not Period.is_first_wed(s.date)
    for r in e.active(K_PROP):
        assert res.count(r.name) <= r.max_m
    for r in e.active(K_SENSE):
        assert res.count(r.name) <= 1


def test_determinista(solved):
    e, res = solved
    res2 = engine.solve(e, None)
    assert res2.assignments == res.assignments


def test_cadena_sense_resposta_i_terra():
    """4 mesos encadenats: qui no ha respost fa exactament 1 guàrdia cada 2 mesos;
    qui té màxim 1 en fa almenys 1 cada 2 mesos (terra)."""
    ym, prev, src = (2027, 1), None, _bytes(demo.demo_input(2027, 1))
    per_month = collections.defaultdict(list)
    for k in range(4):
        e = load_entrada(_with_absences(src, 10 + k, n_vac=4))
        h = parse_planning(prev, list(e.radiologists)) if prev else None
        res = engine.solve(e, h)
        assert not res.uncovered
        for r in e.active():
            per_month[r.name].append(res.count(r.name))
        prev = _bytes(writer.write_planning(res))
        ym = next_month(*ym)
        src = _bytes(templates.build_input_from(src, *ym))
    e = load_entrada(src)
    for r in e.active(K_SENSE):
        v = per_month[r.name]
        assert all(v[i] + v[i + 1] == 1 for i in range(3)), (r.name, v)
    for r in e.active(K_PROP):
        if r.max_m == 1 and r.role not in (ROL_DG_NIT, ROL_CS_DIA):
            assert sum(per_month[r.name]) >= 2, (r.name, per_month[r.name])


def test_historial_llegeix_canvis_manuals(solved):
    e, res = solved
    wb = writer.write_planning(res)
    h = parse_planning(wb, list(e.radiologists))
    for r in e.active():
        if r.klass in (K_PROP, K_SENSE, K_NOU):
            g = h.guards.get(r.name, {POOL_L: 0, POOL_C: 0})
            assert g[POOL_L] == res.count(r.name, POOL_L) and g[POOL_C] == res.count(r.name, POOL_C), r.name
    # Canvi manual: una guàrdia de nit passa a una altra persona
    s, a = next((s, n) for s, n in res.assignments.items() if s.shift == 'nit' and s.group == 'NB' and s.dow == 0)
    b = next(r.name for r in e.active(K_SENSE) if r.name != a)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=18, max_row=23):
            for c in row:
                if c.value == f"N i B: {a}" and as_date(ws.cell(row=3, column=c.column).value) == s.date:
                    c.value = f"N i B: {b}"
    h2 = parse_planning(wb, list(e.radiologists))
    tot = lambda hh, n: sum(hh.guards.get(n, {}).values())
    assert tot(h2, a) == tot(h, a) - 1 and tot(h2, b) == tot(h, b) + 1
    assert h2.has_hist_sheet and h2.share


def test_compatible_office_script(solved, tmp_path):
    e, res = solved
    p = tmp_path / 'p.xlsx'
    writer.write_planning(res, p)
    rows, month = osp.parse(p)
    assert month == e.period.month
    names = collections.Counter(n for d, lbl, role, n in rows if role in ('NB', 'G1', 'G2'))
    for s, n in res.assignments.items():
        assert names[n] > 0
    assert min(d for d, *_ in rows) >= e.period.start and max(d for d, *_ in rows) <= e.period.end


def test_dia_impossible_queda_sense_cobrir():
    wb = demo.demo_input(Y, M)
    ws = wb['Vacances']
    target = next(c for c in range(2, ws.max_column + 1)
                  if as_date(ws.cell(row=4, column=c).value)
                  and as_date(ws.cell(row=4, column=c).value).weekday() == 1
                  and as_date(ws.cell(row=4, column=c).value) >= dt.date(Y, M, 8))
    for r in range(6, ws.max_row + 1):
        if ws.cell(row=r, column=1).value:
            ws.cell(row=r, column=target).value = 'C'
    e = load_entrada(_bytes(wb))
    res = engine.solve(e, None)
    d = as_date(ws.cell(row=4, column=target).value)
    assert [s.date for s in res.uncovered] == [d]


def test_format_antic_de_vacances():
    wb = demo.demo_input(Y, M)
    ws = wb['Vacances']
    wb.remove(ws)
    ws = wb.create_sheet('Vacances')
    ws['A1'], ws['B1'], ws['F1'] = 'Mes/Any:', 'Març 2027', '26'
    ws['A3'] = 'Radiòleg'
    for d in range(1, 32):
        ws.cell(row=3, column=1 + d, value=d)
    ws['A4'] = 'Alfa, Anna'
    e = load_entrada(_bytes(wb))
    assert e.period.start == dt.date(2027, 3, 1) and e.period.end == dt.date(2027, 3, 31)
    assert dt.date(2027, 3, 26) in e.period.festius

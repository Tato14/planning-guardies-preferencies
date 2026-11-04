"""Proves de regressió de la revisió independent (H = alta, M = mitjana, L = baixa)."""
import copy
import datetime as dt
import io

import openpyxl

from guardies import demo, engine, survey, templates, writer
from guardies.history import parse_planning
from guardies.io_entrada import load_entrada
from guardies.model import K_SENSE, POOL_C, POOL_L, PREF_NO, Radiologist, ROL_NOU
from guardies.periods import as_date, fold

Y, M = 2027, 3


def _b(wb):
    return templates.to_bytes(wb)


def _entrada(rads=None, cfg=None, y=Y, m=M):
    rads = rads if rads is not None else demo.demo_radiologists()
    cfg = cfg if cfg is not None else demo.demo_config(rads)
    return load_entrada(_b(templates.build_input(rads, cfg, y, m)))


def test_h1_maxim_zero_es_sense_resposta():
    rads = demo.demo_radiologists()
    rads[20].max_m, rads[20].min_m = 0, 0
    e = _entrada(rads)
    assert e.radiologists[rads[20].name].klass == K_SENSE
    assert any('màxim 0' in w for w in e.warnings)


def test_h2_cognom_compartit_no_es_casa():
    rads = demo.demo_radiologists()
    a, b = rads[20].name, rads[21].name
    rads[21].name = a.split(',')[0] + ', Zoe'            # dos amb el mateix cognom
    e = _entrada(rads)
    res = engine.solve(e, None)
    wb = writer.write_planning(res)
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=18, max_row=23, min_col=2, max_col=8):
            for c in row:
                if isinstance(c.value, str) and c.value.startswith('N i B: '):
                    c.value = 'N i B: ' + a.split(',')[0]       # només el cognom
                    break
    h = parse_planning(wb, list(e.radiologists), e.config)
    assert a.split(',')[0] in h.unmatched


def test_h3_tarda_fixa_de_dijous_no_genera_guardies_invisibles():
    rads = demo.demo_radiologists()
    cfg = demo.demo_config(rads)
    cfg.fix_slots[(3, '16-20', 'N')] = rads[2].name
    e = _entrada(rads, cfg)
    res = engine.solve(e, None)
    assert not any(s.dow == 3 and s.shift == '16-20' for s in res.slots)
    h = parse_planning(writer.write_planning(res), list(e.radiologists), e.config)
    for r in e.active():
        if r.klass in ('P', 'S', 'NOU'):
            assert sum(h.guards.get(r.name, {}).values()) == res.count(r.name), r.name


def test_m1_nou_incorporat_respecta_el_no_i_el_1r_dimecres():
    rads = demo.demo_radiologists()
    nou = next(r for r in rads if r.role == ROL_NOU)
    nou.prefs['dc_tarda'] = PREF_NO
    res = engine.solve(_entrada(rads), None)
    assert res.count(nou.name) == 0


def test_m2_titulars_buits_surten_pendents():
    rads = demo.demo_radiologists()
    cfg = demo.demo_config(rads)
    cfg.fix_slots[(3, '8-16', 'N')] = ''
    cfg.friday = {}
    e = _entrada(rads, cfg)
    res = engine.solve(e, None)
    pend = [f for f in res.fixed if f.status == 'pendent']
    assert any(f.date.weekday() == 3 and f.role == 'N' for f in pend)
    fridays = [d for d in e.period.dates if d.weekday() == 4 and d not in e.period.festius]
    assert sum(1 for f in pend if f.date.weekday() == 4) == len(fridays)


def test_m4_entrada_seguent_porta_les_absencies_de_final_de_mes():
    rads = demo.demo_radiologists()
    wb = templates.build_input(rads, demo.demo_config(rads), 2027, 1)
    ws = wb['Vacances']
    cols = {as_date(ws.cell(row=4, column=c).value): c for c in range(2, ws.max_column + 1)
            if as_date(ws.cell(row=4, column=c).value)}
    name = rads[22].name
    row = next(r for r in range(6, ws.max_row + 1) if ws.cell(row=r, column=1).value == name)
    for d in (dt.date(2027, 1, 30), dt.date(2027, 1, 31)):
        ws.cell(row=row, column=cols[d]).value = 'V'
    nxt = load_entrada(_b(templates.build_input_from(_b(wb), 2027, 2)))
    assert nxt.absences[name].get(dt.date(2027, 2, 1)) == 'R' and nxt.absences[name].get(dt.date(2027, 2, 2)) == 'R'


def test_m5_tarda_fixa_de_dilluns_no_compta_com_a_guardia():
    e = _entrada()
    res = engine.solve(e, None)
    wb = writer.write_planning(res)
    target = None
    for ws in wb.worksheets:
        if 'Setmana' not in ws.title:
            continue
        for col in range(2, 9):
            d = as_date(ws.cell(row=3, column=col).value)
            if d and d.weekday() == 0 and e.period.contains(d) and d not in e.period.festius:
                ws.cell(row=11, column=col).value = f"N i B: {e.active(K_SENSE)[0].name}"
                target = e.active(K_SENSE)[0].name
                break
        if target:
            break
    h = parse_planning(wb, list(e.radiologists), e.config)
    assert sum(h.guards.get(target, {}).values()) == res.count(target)


def test_m7_no_de_dimarts_val_per_al_dimarts_festiu():
    rads = demo.demo_radiologists()
    for r in rads:
        if r.klass in ('P', 'S'):
            r.prefs['dt_nit'] = PREF_NO if r.name != rads[25].name else None
    e = _entrada(rads, y=2026, m=12)            # dt 08/12/2026 és festiu
    res = engine.solve(e, None)
    s = next(s for s in res.assignments if s.date == dt.date(2026, 12, 8) and s.shift == 'nit')
    assert e.radiologists[res.assignments[s]].prefs.get('dt_nit') != PREF_NO


def test_m8_enquesta_inicials_empats_i_duplicats():
    rads = {n: Radiologist(n) for n in ['Alzina Roure, Anna', 'Alzina Roure, Pere', 'Freixe, Marta']}
    rows = [{'name': 'Alzina Roure, Pere', 'min': 1, 'max': 3, 'prefs': {}, 'mati': None, 'dies': ''},
            {'name': 'Alzina Roure', 'min': 2, 'max': 4, 'prefs': {}, 'mati': None, 'dies': ''},
            {'name': 'Freixe, Marta', 'min': 1, 'max': 2, 'prefs': {}, 'mati': None, 'dies': ''},
            {'name': 'Freixe', 'min': 5, 'max': 6, 'prefs': {}, 'mati': None, 'dies': ''},
            {'name': 'Freixe, Jordi', 'min': 1, 'max': 1, 'prefs': {}, 'mati': None, 'dies': ''}]
    out, rep = survey.merge(rads, rows)
    assert out['Alzina Roure, Pere'].max_m == 3 and out['Alzina Roure, Anna'].max_m is None
    assert 'Alzina Roure' in rep['ambigus']
    assert out['Freixe, Marta'].max_m == 2 and ('Freixe', 'Freixe, Marta') in rep['duplicats']
    assert 'Freixe, Jordi' in rep['nous']


def test_l1_un_mes_inactiu_no_esborra_el_deute():
    rads = demo.demo_radiologists()
    e1 = _entrada(rads, y=2027, m=1)
    r1 = engine.solve(e1, None)
    p1 = writer.write_planning(r1)
    h1 = parse_planning(p1, list(e1.radiologists), e1.config)
    who = next(n for n in r1.shares if abs(h1.debt_now(n, POOL_L)) + abs(h1.debt_now(n, POOL_C)) > 0.05)
    debt = {k: h1.debt_now(who, k) for k in (POOL_L, POOL_C)}
    rads2 = copy.deepcopy(rads)
    next(r for r in rads2 if r.name == who).active = False
    e2 = _entrada(rads2, y=2027, m=2)
    r2 = engine.solve(e2, h1)
    h2 = parse_planning(writer.write_planning(r2), list(e2.radiologists), e2.config)
    for k in (POOL_L, POOL_C):
        assert abs(h2.debt_now(who, k) - debt[k]) < 1e-3


def test_l3_pestanya_duplicada_no_compta_dues_vegades():
    e = _entrada()
    res = engine.solve(e, None)
    wb = writer.write_planning(res)
    first = next(ws for ws in wb.worksheets if 'Setmana' in ws.title)
    cp = wb.copy_worksheet(first)
    cp.title = first.title[:20] + ' (2)'
    h = parse_planning(wb, list(e.radiologists), e.config)
    assert sum(sum(v.values()) for v in h.guards.values()) == len(res.assignments)
    assert h.warnings

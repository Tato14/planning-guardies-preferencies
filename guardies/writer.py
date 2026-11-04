"""Escriptura del planning: pestanyes setmanals (format compatible amb l'Office Script de M365),
Resum, Pendents i Historic (estat per al mes següent)."""
from __future__ import annotations

import datetime as dt

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

from . import __version__
from .engine import Result, pref_of
from .history import HIST_SHEET
from .model import K_FIX, K_NOU, K_PROP, K_SENSE, POOL_C, POOL_L, PREF_SI
from .periods import fmt_dia, fold, week_sheet_name

SHIFT_ROWS = {'8-16': (4, 9), '16-20': (11, 16), '20-08': (18, 23)}
BLAU = '4472C4'
F_BOLD = Font(bold=True, size=10)
F_SMALL = Font(size=10)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
FILL_FIX = PatternFill('solid', fgColor='E8F1FB')
FILL_ROT = PatternFill('solid', fgColor='F4FBF0')
FILL_FEST = PatternFill('solid', fgColor='FFF3CD')
FILL_SUB = PatternFill('solid', fgColor='FBEFD8')
FILL_OUT = PatternFill('solid', fgColor='F2F2F2')
FILL_HDR = PatternFill('solid', fgColor=BLAU)
THIN = Side(style='thin', color='BFBFBF')
CLASSE_TXT = {K_PROP: 'Proporcional', K_SENSE: 'Sense resposta (1 cada 2 mesos)', K_FIX: 'Només fixos',
              K_NOU: 'Nou-incorporat'}


def _week_sheet(wb, monday, res):
    per = res.entrada.period
    ws = wb.create_sheet(week_sheet_name(monday))
    sunday = monday + dt.timedelta(days=6)
    ws['A1'] = f"Setmana {monday.isocalendar()[1]:02d} · {fmt_dia(monday)} – {fmt_dia(sunday, True)}"
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = f"Planning de {per.label}: {per.text()} ({per.n_weeks} setmanes)"
    ws['A2'].font = Font(italic=True, size=9, color='595959')
    ws['A3'] = 'Franja'
    ws['A3'].font, ws['A3'].fill, ws['A3'].alignment = Font(bold=True, color='FFFFFF'), FILL_HDR, CENTER
    for (a, b), txt in zip(SHIFT_ROWS.values(), ('8:00-16:00', '16:00-20:00', '20:00-08:00')):
        ws.merge_cells(start_row=a, start_column=1, end_row=b, end_column=1)
        c = ws.cell(row=a, column=1, value=txt)
        c.font, c.alignment = Font(bold=True), CENTER
        c.fill = PatternFill('solid', fgColor='D5E3F0')
        for r in range(a, b + 1):
            for col in range(2, 9):
                ws.cell(row=r, column=col).border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
    for i in range(7):
        d = monday + dt.timedelta(days=i)
        h = ws.cell(row=3, column=2 + i, value=d)
        h.number_format, h.alignment = '[$-403]ddd dd/mm/yyyy', CENTER
        if not per.contains(d):
            h.font, h.fill = Font(bold=True, size=10, color='808080'), PatternFill('solid', fgColor='D9D9D9')
            note = '(fora del període — vegeu el planning anterior)' if d < per.start else '(fora del període — vegeu el planning següent)'
            c = ws.cell(row=4, column=2 + i, value=note)
            c.font, c.alignment = Font(italic=True, size=9, color='808080'), Alignment(wrap_text=True, vertical='top')
            for r in range(4, 24):
                ws.cell(row=r, column=2 + i).fill = FILL_OUT
        elif d in per.festius:
            h.font, h.fill = Font(bold=True, size=10, color='C0392B'), FILL_FEST
        else:
            h.font, h.fill = Font(bold=True, size=10, color='FFFFFF'), FILL_HDR
    ws.row_dimensions[3].height = 30
    ws.column_dimensions['A'].width = 14
    for c in range(2, 9):
        ws.column_dimensions[L(c)].width = 32
    return ws


def _put(ws, col, label, lines, fill):
    a, b = SHIFT_ROWS[label]
    for k, line in enumerate(lines):
        if a + k > b:
            break
        c = ws.cell(row=a + k, column=col, value=line)
        c.alignment, c.fill = LEFT, fill
        c.font = F_BOLD if line.startswith(('N:', 'B:', 'N i B', 'Grup', '(', 'Dissabte', 'Diumenge', 'FESTIU', '[')) else F_SMALL


def write_planning(res: Result, path=None):
    e, per = res.entrada, res.entrada.period
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _instruccions(wb.create_sheet('Instruccions'), res)
    mondays = per.week_mondays()
    sheets = {m: _week_sheet(wb, m, res) for m in mondays}
    asg = res.assignments
    unc = set(res.uncovered)

    def who(s):
        if s in unc:
            return '[SENSE COBRIR]'
        return asg.get(s, '—')

    from .model import Slot
    fixed_by = {}
    for fa in res.fixed:
        fixed_by.setdefault((fa.date, fa.shift), []).append(fa)

    def fixed_lines(d, shift):
        out = []
        for fa in fixed_by.get((d, shift), []):
            if fa.status == 'ok':
                nm = f"{fa.name} ({fa.note})" if fa.note else fa.name
            elif fa.status == 'substituir':
                nm = f"[SUBSTITUIR — {fa.name}: {fa.reason}]"
            else:
                nm = f"[PENDENT — {fa.note or 'torn fix'}: {fa.reason}]"
            out.append(f"{fa.role}: {nm}")
        return out

    for d in per.dates:
        dow = d.weekday()
        ws = sheets[d - dt.timedelta(days=dow)]
        col = 2 + dow
        fest = d in per.festius
        S = lambda sh, g: who(Slot(d, sh, g))
        # 8-16
        if fest and dow <= 2:
            _put(ws, col, '8-16', ['FESTIU — Torn dia 8-20:', f"N i B: {S('dia', 'NB')}"], FILL_FEST)
        elif fest and dow in (3, 4):
            _put(ws, col, '8-16', ['FESTIU — Torn dia 8-20:', f"Grup 1: {S('dia', 'G1')}", f"Grup 2: {S('dia', 'G2')}"], FILL_FEST)
        elif dow == 5:
            _put(ws, col, '8-16', ['Dissabte — Torn dia 8-20:', f"Grup 1: {S('dia', 'G1')}", f"Grup 2: {S('dia', 'G2')}"], FILL_ROT)
        elif dow == 6:
            _put(ws, col, '8-16', ['Diumenge — Torn dia 8-20:', f"N i B: {S('dia', 'NB')}"], FILL_ROT)
        else:
            fx = fixed_lines(d, '8-16')
            if fx:
                _put(ws, col, '8-16', fx, FILL_SUB if any('[' in l for l in fx) else FILL_FIX)
        # 16-20
        if fest:
            _put(ws, col, '16-20', ['(inclòs torn dia)'], FILL_FEST)
        elif dow >= 5:
            _put(ws, col, '16-20', ['(inclòs torn dia)'], FILL_ROT)
        else:
            fx = fixed_lines(d, '16-20')
            if fx:
                _put(ws, col, '16-20', fx, FILL_SUB if any('[' in l for l in fx) else
                     (FILL_FEST if (dow == 2 and d.day <= 7) else FILL_FIX))
            elif dow == 2:
                _put(ws, col, '16-20', [f"N i B: {S('16-20', 'NB')}"], FILL_ROT)
            elif dow in (3, 4):
                _put(ws, col, '16-20', [f"Grup 1: {S('16-20', 'G1')}", f"Grup 2: {S('16-20', 'G2')}"], FILL_ROT)
        # 20-08
        if fest and dow <= 2:
            _put(ws, col, '20-08', ['FESTIU — Torn nit 20-08:', f"N i B: {S('nit', 'NB')}"], FILL_FEST)
        elif fest and dow in (3, 4):
            _put(ws, col, '20-08', ['FESTIU — Torn nit 20-08:', f"Grup 1: {S('nit', 'G1')}", f"Grup 2: {S('nit', 'G2')}"], FILL_FEST)
        elif dow in (0, 1, 2, 6):
            _put(ws, col, '20-08', [f"N i B: {S('nit', 'NB')}"], FILL_ROT)
        else:
            _put(ws, col, '20-08', [f"Grup 1: {S('nit', 'G1')}", f"Grup 2: {S('nit', 'G2')}"], FILL_ROT)

    _resum(wb.create_sheet('Resum'), res)
    _pendents(wb.create_sheet('Pendents'), res)
    _historic(wb.create_sheet(HIST_SHEET), res)
    wb.active = 1
    if path:
        wb.save(path)
    return wb


def _header(ws, row, cols, widths=None):
    for c, h in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font, cell.fill = Font(bold=True, color='FFFFFF'), FILL_HDR
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if widths:
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[L(c)].width = w


def _instruccions(ws, res):
    per = res.entrada.period
    lines = [
        (f"PLANNING DE GUÀRDIES — Telediagnòstic IDI · {per.label}", True),
        (f"Període: {per.text()} ({per.n_weeks} setmanes). Generat amb l'app v{__version__}.", False),
        ("", False),
        ("Pestanyes", True),
        ("• Setmana NN: una per setmana, de dilluns a diumenge. Els dies fora del període surten en gris.", False),
        ("• Resum: guàrdies de cadascú comparades amb el seu objectiu (capacitat declarada i deute acumulat).", False),
        ("• Pendents: torns fixos que cal substituir (amb propostes), guàrdies sense cobrir i observacions individuals.", False),
        ("• Historic: estat per calcular el mes següent. No l'editeu ni l'esborreu.", False),
        ("", False),
        ("Canvis manuals", True),
        ("• Podeu intercanviar noms directament a les pestanyes setmanals: el mes següent l'app llegeix el que", False),
        ("  hi hagi escrit i compensa les diferències. Mantingueu el format «N i B: Nom», «Grup 1: Nom»…", False),
        ("• No canvieu els noms de les pestanyes setmanals: els fluxos de Microsoft 365 les busquen per «Setmana».", False),
    ]
    for i, (t, b) in enumerate(lines, start=1):
        ws.cell(row=i, column=1, value=t).font = Font(bold=b, size=12 if i == 1 else 11)
    ws.column_dimensions['A'].width = 110


def _resum(ws, res):
    e, per = res.entrada, res.entrada.period
    ws['A1'] = f"RESUM — {per.label} ({per.text()})"
    ws['A1'].font = Font(bold=True, size=12)
    tot_slots = len(res.slots)
    ws['A2'] = (f"{tot_slots} guàrdies · {tot_slots - len(res.uncovered)} cobertes · "
                f"{sum(1 for s, n in res.assignments.items() if pref_of(e.radiologists[n], s, per) == PREF_SI)} en torns «Sí» "
                f"· L = laborables (dl-dv) · C = caps de setmana i festius")
    ws['A2'].font = Font(italic=True, size=9, color='595959')
    cols = ['Nom', 'Rol', 'Classe', 'Actiu', 'Màx./mes', 'Objectiu L', 'Objectiu C', 'Guàrdies L', 'Guàrdies C',
            'Total', 'Diferència', 'Torns «Sí»', 'Dates de guàrdia', 'Fixos', 'Observacions']
    _header(ws, 4, cols, [32, 18, 26, 7, 8, 9, 9, 9, 9, 7, 9, 9, 40, 7, 50])
    fixos = {}
    for fa in res.fixed:
        if fa.status == 'ok':
            fixos[fa.name] = fixos.get(fa.name, 0) + 1
    r = 5
    for rad in sorted(e.radiologists.values(), key=lambda x: fold(x.name)):
        n = rad.name
        tL, tC = res.targets.get(n, {}).get(POOL_L), res.targets.get(n, {}).get(POOL_C)
        gL, gC = res.count(n, POOL_L), res.count(n, POOL_C)
        if rad.klass == K_SENSE:
            tL, tC = None, None
            obj = 1 if n in res.due else 0
        else:
            obj = (tL or 0) + (tC or 0) if tL is not None else None
        dates = ', '.join(f"{s.date:%d/%m}" for s in sorted((s for s, nm in res.assignments.items() if nm == n),
                                                               key=lambda s: s.date))
        si = sum(1 for s, nm in res.assignments.items() if nm == n and pref_of(rad, s, per) == PREF_SI)
        vals = [n, rad.role, CLASSE_TXT.get(rad.klass, rad.klass), 'Sí' if rad.active else 'No', rad.max_m,
                round(tL, 2) if tL is not None else None, round(tC, 2) if tC is not None else None,
                gL, gC, gL + gC, round(gL + gC - obj, 2) if obj is not None else None, si or None,
                dates or None, fixos.get(n), ' · '.join(x for x in (rad.dies_puntuals, rad.observacions) if x) or None]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            if not rad.active:
                cell.font = Font(italic=True, color='A6A6A6')
        r += 1
    ws.freeze_panes = 'B5'
    ws.auto_filter.ref = f"A4:{L(len(cols))}{r - 1}"


def _pendents(ws, res):
    e = res.entrada
    r = 1
    ws.cell(row=r, column=1, value='PENDENTS I OBSERVACIONS').font = Font(bold=True, size=12)
    r += 2
    ws.cell(row=r, column=1, value='Torns fixos que cal cobrir a mà').font = Font(bold=True, size=11)
    r += 1
    _header(ws, r, ['Data', 'Franja', 'Rol', 'Titular', 'Motiu', 'Propostes (per ordre)'], [14, 8, 7, 30, 34, 70])
    r += 1
    pend = [fa for fa in res.fixed if fa.status != 'ok']
    for fa in pend:
        for c, v in enumerate([fmt_dia(fa.date, True), fa.shift, fa.role, fa.name or f"({fa.note})", fa.reason,
                               ', '.join(fa.suggestions) or 'cap candidat disponible'], start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1
    if not pend:
        ws.cell(row=r, column=1, value='Cap.')
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='Guàrdies sense cobrir').font = Font(bold=True, size=11)
    r += 1
    if res.uncovered:
        for s in res.uncovered:
            ws.cell(row=r, column=1, value=s.label())
            ws.cell(row=r, column=2, value='Ningú disponible que compleixi les regles: cal cobrir-la a mà.')
            r += 1
    else:
        ws.cell(row=r, column=1, value='Cap.')
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='Avisos').font = Font(bold=True, size=11)
    r += 1
    for w in (e.warnings + res.warnings) or ['Cap.']:
        ws.cell(row=r, column=1, value=w)
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='Observacions individuals (premissa 8)').font = Font(bold=True, size=11)
    r += 1
    _header(ws, r, ['Nom', 'Dies puntuals (matí)', 'Observacions'])
    r += 1
    for rad in sorted(e.radiologists.values(), key=lambda x: fold(x.name)):
        if rad.dies_puntuals or rad.observacions:
            ws.cell(row=r, column=1, value=rad.name)
            ws.cell(row=r, column=2, value=rad.dies_puntuals or None)
            ws.cell(row=r, column=3, value=rad.observacions or None)
            r += 1
    for c, w in zip('ABCDEF', (30, 30, 70, 30, 34, 70)):
        ws.column_dimensions[c].width = max(ws.column_dimensions[c].width or 0, w)


def _historic(ws, res):
    e, per = res.entrada, res.entrada.period
    ws['A1'] = "HISTORIC — l'app el llegeix per calcular el mes següent. No l'editeu ni l'esborreu."
    ws['A1'].font = Font(bold=True, color='C00000')
    for c, v in enumerate(['Periode', per.code, 'Inici', per.start, 'Fi', per.end, 'Versio', __version__], start=1):
        ws.cell(row=2, column=c, value=v)
    cols = ['Nom', 'Classe', 'Deute L (inici)', 'Deute C (inici)', 'Quota L', 'Quota C', 'Ultima guardia']
    for c, h in enumerate(cols, start=1):
        ws.cell(row=4, column=c, value=h).font = Font(bold=True)
    r = 5
    hist_last = getattr(res, 'last_guard_before', {}) or {}
    for rad in sorted(e.radiologists.values(), key=lambda x: fold(x.name)):
        n = rad.name
        sh, db = res.shares.get(n), res.debts.get(n)
        vals = [n, rad.klass,
                round(db[POOL_L], 4) if db else None, round(db[POOL_C], 4) if db else None,
                round(sh[POOL_L], 4) if sh else None, round(sh[POOL_C], 4) if sh else None,
                hist_last.get(n)]
        for c, v in enumerate(vals, start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1
    ws.column_dimensions['A'].width = 34
    for c in 'BCDEFG':
        ws.column_dimensions[c].width = 14

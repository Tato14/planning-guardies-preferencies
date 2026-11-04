"""Plantilles d'entrada: llibre complet (Instruccions, Configuració, Radiòlegs, Vacances).

Tot es genera per codi (sense fitxers binaris al repositori):
  build_input(radiologists, config, any, mes)       -> entrada d'un mes
  build_input_from(entrada_anterior, any, mes)       -> entrada del mes amb la mateixa llista
  blank_input(any, mes)                              -> plantilla buida
"""
from __future__ import annotations

import datetime as dt
import io

import openpyxl
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from .io_entrada import load_entrada, read_config, read_radiologists
from .model import (DIES_REINCORPORACIO, PREF_HEADERS, PREF_KEYS, PREF_NO, PREF_SI, ROL_CS_DIA, ROL_DG_NIT,
                    ROL_GUARDIES, ROLS, Config, Radiologist)
from .periods import DIES_CA, ORDINALS, festius_entre, fmt_dia, month_label, period_for_month

BLAU = '4472C4'
F_TITOL = Font(bold=True, size=11)
F_VALOR = Font(bold=True, size=11, color='1F3864')
F_PETIT = Font(italic=True, size=9, color='808080')
F_HDR = Font(bold=True, size=9, color='FFFFFF')
FILL_HDR = PatternFill('solid', fgColor=BLAU)
FILL_CTX_HDR = PatternFill('solid', fgColor='A6A6A6')
FILL_CTX = PatternFill('solid', fgColor='EDEDED')
FILL_CAPSET = PatternFill('solid', fgColor='EEF3FA')
FILL_FESTIU = PatternFill('solid', fgColor='FFE699')
FILL_META = PatternFill('solid', fgColor='F2F2F2')
FILL_SI = PatternFill('solid', fgColor='C6EFCE')
FILL_NO = PatternFill('solid', fgColor='F8CBAD')
CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
WRAP = Alignment(horizontal='left', vertical='top', wrap_text=True)
THIN = Side(style='thin', color='BFBFBF')
THICK = Side(style='medium', color='1F3864')
CODI_COLORS = {'V': 'C6EFCE', 'B': 'F8CBAD', 'C': 'BDD7EE', 'G': 'FFD966', 'X': 'D9D9D9'}

RAD_HEADERS = (['Nom (Cognoms, Nom)', 'Rol', 'Correu', 'Actiu', 'Mín./mes', 'Màx./mes']
               + [PREF_HEADERS[k] for k in PREF_KEYS] + ['Matí', 'Dies puntuals (matí)', 'Observacions'])


def _dv(ws, formula, rng, title, msg, error=True):
    dv = DataValidation(type='list', formula1=formula, allow_blank=True, showDropDown=False,
                        showErrorMessage=error, showInputMessage=True, promptTitle=title, prompt=msg,
                        errorTitle=title, error=msg)
    ws.add_data_validation(dv)
    dv.add(rng)


def _si_no(v):
    return 'Sí' if v == PREF_SI else ('No' if v == PREF_NO else None)


def write_radiologists(ws, rads: list, spare=25):
    ws['A1'] = 'RADIÒLEGS — llista mestra, capacitat i preferències de guàrdia'
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = ("Rol: guàrdies / fix / fix-i-guàrdies / nou-incorporat / només-dg-nit / només-cap-setmana-dia. "
                "Actiu = No: aquest mes no fa res. Mín./Màx. per mes (buit = no ha respost: 1 guàrdia cada 2 mesos). "
                "Torns: Sí = preferit · No = no pot · buit = indiferent. Matí: per a substitucions de fixos.")
    ws['A2'].font = F_PETIT
    for c, h in enumerate(RAD_HEADERS, start=1):
        cell = ws.cell(row=3, column=c, value=h)
        cell.font = Font(bold=True, color='FFFFFF', size=10)
        cell.fill = FILL_HDR
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[3].height = 32
    for i, r in enumerate(rads, start=4):
        vals = [r.name, r.role, r.email or None, 'Sí' if r.active else 'No',
                r.min_m if r.min_m is not None else None, r.max_m if r.max_m is not None else None]
        vals += [_si_no(r.prefs.get(k)) for k in PREF_KEYS]
        vals += [_si_no(r.mati), r.dies_puntuals or None, r.observacions or None]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=i, column=c, value=(int(v) if isinstance(v, float) and v.is_integer() else v))
            if 5 <= c <= 6 + len(PREF_KEYS) + 1:
                cell.alignment = CENTER
    last = 3 + len(rads) + spare
    first_pref, last_pref = 7, 6 + len(PREF_KEYS)
    c_mati = last_pref + 1
    _dv(ws, '"' + ','.join(ROLS) + '"', f"B4:B{last}", 'Rol', 'Tria un rol de la llista.')
    _dv(ws, '"Sí,No"', f"D4:D{last}", 'Actiu', '"No" = aquest mes no fa ni fixos ni guàrdies. Buit = Sí.')
    _dv(ws, '"Sí,No"', f"{L(first_pref)}4:{L(c_mati)}{last}", 'Preferència',
        'Sí = preferit · No = no pot · buit = indiferent.')
    rng = f"{L(first_pref)}4:{L(c_mati)}{last}"
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"Sí"'], fill=FILL_SI))
    ws.conditional_formatting.add(rng, CellIsRule(operator='equal', formula=['"No"'], fill=FILL_NO))
    ws.conditional_formatting.add(f"A4:{L(len(RAD_HEADERS))}{last}",
                                  FormulaRule(formula=['$D4="No"'], font=Font(italic=True, color='A6A6A6')))
    widths = {1: 34, 2: 20, 3: 30, 4: 7, 5: 7, 6: 7}
    for c in range(1, len(RAD_HEADERS) + 1):
        ws.column_dimensions[L(c)].width = widths.get(c, 7 if c <= c_mati else (24 if c == c_mati + 1 else 50))
    ws.freeze_panes = 'B4'


def write_config(ws, cfg: Config):
    ws['A1'] = 'CONFIGURACIÓ'
    ws['A1'].font = Font(bold=True, size=12)
    ws['A3'] = 'TORNS FIXOS'
    ws['A3'].font = F_TITOL
    hdr = ['Dia setmana', 'Franja', 'Rol', 'Radiòleg (nom exacte)', 'Notes']
    for c, h in enumerate(hdr, start=1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = FILL_HDR
    dies = ['Dilluns', 'Dimarts', 'Dimecres', 'Dijous', 'Divendres']
    rows = []
    for dow in range(5):
        for shift in ('8-16', '16-20'):
            if dow == 4 and shift == '8-16':
                continue
            for role in ('N', 'B', 'N i B'):
                if (dow, shift, role) in cfg.fix_slots:
                    rows.append((dies[dow], shift, role, cfg.fix_slots[(dow, shift, role)] or None, ''))
    if cfg.first_wed is not None:
        rows.append(('Dimecres-1r-mes', '16-20', 'N i B', cfg.first_wed or None,
                     '1r dimecres del mes natural, tarda (la resta de dimecres a la tarda són guàrdia)'))
    for k in range(1, 6):
        rows.append((f'Divendres-{ORDINALS[k]}', '8-16', 'N i B', cfg.friday.get(k) or None,
                     f'{ORDINALS[k]} divendres del mes natural, matí' + (' (si el mes en té cinc)' if k == 5 else '')))
    r = 5
    for row in rows:
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1
    for k in range(6):   # files lliures per afegir fixos
        r += 1
    r += 1
    ws.cell(row=r, column=1, value='PARÀMETRES').font = F_TITOL
    ws.cell(row=r + 1, column=1, value='Descans mínim entre guàrdies (dies)')
    ws.cell(row=r + 1, column=2, value=cfg.rest_days)
    ws.cell(row=r + 2, column=1, value='Mínim per a tothom qui fa guàrdies: 1 guàrdia cada N mesos')
    ws.cell(row=r + 2, column=2, value=cfg.min_every_months)
    ws.cell(row=r + 4, column=1, value=("Els noms han de coincidir exactament amb la pestanya Radiòlegs. Per afegir un torn "
                                        "fix, escriu el dia (Dilluns…Divendres), la franja (8-16 o 16-20), el rol "
                                        "(N, B o N i B) i el nom. Els divendres al matí van per ordinal del mes.")).font = F_PETIT
    for c, w in zip('ABCDE', (22, 9, 8, 34, 70)):
        ws.column_dimensions[c].width = w


def write_vacances(ws, year, month, start, end, festius, names, inactive=(), spare_rows=15, context_marks=None):
    n_ctx = max(0, DIES_REINCORPORACIO - 1)
    ctx = [start - dt.timedelta(days=k) for k in range(n_ctx, 0, -1)]
    dates = [start + dt.timedelta(days=i) for i in range((end - start).days + 1)]
    all_dates = ctx + dates
    c0, HDR, FEST, FIRST = 2, 4, 5, 6
    last_col = c0 + len(all_dates) - 1
    last_row = FIRST + len(names) + spare_rows - 1
    n_weeks = len({d - dt.timedelta(days=d.weekday()) for d in dates})
    right = max(32, last_col)
    ws['A1'] = 'Mes del planning:'
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=9)
    ws['B1'] = month_label(year, month)
    ws.merge_cells(start_row=1, start_column=10, end_row=1, end_column=right)
    ws['J1'] = 'Marca les absències de cada radiòleg: V, B, C, G o X.'
    ws['A2'] = 'Període planificat:'
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=9)
    ws['B2'] = start
    ws.merge_cells(start_row=2, start_column=10, end_row=2, end_column=17)
    ws['J2'] = 'fins al'
    ws.merge_cells(start_row=2, start_column=18, end_row=2, end_column=25)
    ws['R2'] = end
    ws.merge_cells(start_row=2, start_column=26, end_row=2, end_column=right)
    ws['Z2'] = f"{n_weeks} setmanes (dl→dg)" + (' · comença amb dia pont' if start.weekday() != 0 else '')
    for ref in ('A1', 'A2', 'J2'):
        ws[ref].font = F_TITOL
    ws['J1'].font = F_PETIT
    for ref in ('B1', 'B2', 'R2'):
        ws[ref].font, ws[ref].fill, ws[ref].alignment = F_VALOR, FILL_META, LEFT
    ws['B2'].number_format = ws['R2'].number_format = 'dd/mm/yyyy'
    ws['Z2'].font = F_PETIT
    ws['A3'] = f'(columnes grises = últims {n_ctx} dies del període anterior)' if n_ctx else None
    ws['A3'].font = F_PETIT
    ws['A4'] = 'Radiòleg'
    ws['A4'].font, ws['A4'].fill = Font(bold=True, color='FFFFFF'), FILL_HDR
    ws['A5'] = 'Festiu (marca F) →'
    ws['A5'].font = Font(bold=True, color='9C5700')
    fset = set(festius)
    for i, d in enumerate(all_dates):
        col = c0 + i
        is_ctx = d < start
        c3 = ws.cell(row=3, column=col, value=DIES_CA[d.weekday()])
        c3.alignment = CENTER
        c3.font = Font(bold=True, size=9, color='808080' if is_ctx else '1F3864')
        if not is_ctx and d.weekday() >= 5:
            c3.fill = FILL_CAPSET
        c4 = ws.cell(row=HDR, column=col, value=d)
        c4.number_format, c4.alignment, c4.font = 'dd/mm', CENTER, F_HDR
        c4.fill = FILL_CTX_HDR if is_ctx else FILL_HDR
        c5 = ws.cell(row=FEST, column=col)
        c5.alignment, c5.font = CENTER, Font(bold=True, color='9C5700')
        if is_ctx:
            c5.fill = FILL_CTX
        elif d in fset:
            c5.value = 'F'
        for r in range(FIRST, last_row + 1):
            cell = ws.cell(row=r, column=col)
            cell.alignment = CENTER
            if is_ctx:
                cell.fill = FILL_CTX
            elif d.weekday() >= 5:
                cell.fill = FILL_CAPSET
        if (d.weekday() == 0 or d == start) and i > 0:
            for r in range(3, last_row + 1):
                ws.cell(row=r, column=col).border = Border(left=THICK)
    col_of = {d: c0 + i for i, d in enumerate(all_dates)}
    for k, name in enumerate(names):
        c = ws.cell(row=FIRST + k, column=1, value=name)
        if name in inactive:
            c.font = Font(italic=True, color='A6A6A6')
        for d, code in (context_marks or {}).get(name, {}).items():
            if d in col_of and d < start:
                ws.cell(row=FIRST + k, column=col_of[d], value=code)
    for r in range(FIRST, last_row + 1):
        ws.cell(row=r, column=1).border = Border(bottom=THIN)
    body = f"B{FIRST}:{L(last_col)}{last_row}"
    for code, color in CODI_COLORS.items():
        ws.conditional_formatting.add(body, CellIsRule(operator='equal', formula=[f'"{code}"'],
                                                       fill=PatternFill('solid', fgColor=color), font=Font(bold=True)))
    ws.conditional_formatting.add(f"B3:{L(last_col)}{last_row}",
                                  FormulaRule(formula=[f'B${FEST}<>""'], fill=FILL_FESTIU,
                                              font=Font(bold=True, color='9C5700')))
    _dv(ws, '"V,B,C,G,X"', body, 'Codi',
        'V = vacances · B = baixa · C = congrés · G = guàrdia en una altra institució · X = ja assignada.')
    _dv(ws, '"F"', f"{L(c0 + len(ctx))}{FEST}:{L(last_col)}{FEST}", 'Festiu', 'F = el dia es tracta com a festiu.')
    ws.column_dimensions['A'].width = 38
    for col in range(c0, last_col + 1):
        ws.column_dimensions[L(col)].width = 5.7
    ws.freeze_panes = f"B{FIRST}"
    ws.sheet_view.zoomScale = 90


INSTRUCCIONS = [
    ("ENTRADA DEL MES — Planning de guàrdies de Telediagnòstic (v4: preferències i capacitat)", True),
    ("{label} · període {ini} → {fi} ({setmanes} setmanes){pont}", True),
    ("", False),
    ("QUÈ CAL FER CADA MES", True),
    ("1. Radiòlegs: revisa la columna Actiu (No = aquest mes no fa res) i, si algú ha canviat, la seva capacitat i preferències.", False),
    ("2. Vacances: marca les absències. V = vacances · B = baixa · C = congrés · G = guàrdia en una altra institució · X = ja assignada.", False),
    ("   V i B deixen 2 dies de marge de reincorporació. Les columnes grises del principi són del període anterior:", False),
    ("   marca-hi V o B només si l'absència ja venia d'abans.", False),
    ("3. Festius: ja hi són els oficials de Catalunya (fila Festiu). Afegeix-hi o treu-ne una F si cal.", False),
    ("4. Puja a l'app aquest fitxer i el planning del mes anterior (l'últim, amb els canvis que s'hi hagin fet).", False),
    ("", False),
    ("COM ES REPARTEIXEN LES GUÀRDIES", True),
    ("• Ja no hi ha roda alfabètica. Cada guàrdia va a qui li toca més segons capacitat i preferències.", False),
    ("• Hi ha dues bosses independents: laborables (dl-dv) i caps de setmana + festius. Cadascuna es reparteix per separat.", False),
    ("• Cadascú rep la mateixa fracció del seu màxim declarat (ajustada per les absències del mes).", False),
    ("  Ningú que faci guàrdies en té menys d'1 cada 2 mesos.", False),
    ("• Qui no ha declarat capacitat (Mín./Màx. buits) i no és fix: exactament 1 guàrdia cada 2 mesos.", False),
    ("• Preferències: «No» no s'assigna mai; «Sí» té prioritat; buit = indiferent.", False),
    ("• El que algú ha fet de més o de menys es compensa els mesos següents (l'app ho llegeix del planning anterior).", False),
    ("• Descans: s'eviten guàrdies en dies consecutius i una nit just abans d'un matí fix.", False),
    ("• Torns fixos (Configuració): matins, tardes de dilluns i dimarts, 1r dimecres, divendres al matí per ordinal", False),
    ("  (1r-4t i 5è). Si el titular no hi és, el planning ho marca i proposa substituts (pestanya Pendents).", False),
]


def write_instructions(ws, year, month, start, end, verificat=True):
    for row in ws.iter_rows():
        for c in row:
            c.value = None
    n_weeks = len({start + dt.timedelta(days=i) - dt.timedelta(days=(start + dt.timedelta(days=i)).weekday())
                   for i in range((end - start).days + 1)})
    lines = list(INSTRUCCIONS)
    if not verificat:
        lines += [("", False), ("ATENCIÓ: algun any del període encara no té el calendari oficial de festius publicat; "
                                "els festius marcats són provisionals.", True)]
    for i, (t, b) in enumerate(lines, start=1):
        c = ws.cell(row=i, column=1, value=t.format(
            label=month_label(year, month), ini=fmt_dia(start, True), fi=fmt_dia(end, True), setmanes=n_weeks,
            pont=' · comença amb dia pont' if start.weekday() != 0 else ''))
        c.font = Font(bold=b, size=12 if i == 1 else 11)
    ws.column_dimensions['A'].width = 125


def build_input(rads: list, cfg: Config, year: int, month: int, start=None, end=None, festius=None,
                context_marks=None):
    p_start, p_end = period_for_month(year, month)
    start, end = start or p_start, end or p_end
    verificat = True
    if festius is None:
        festius, verificat = festius_entre(start, end)
    wb = openpyxl.Workbook()
    write_instructions(wb.active, year, month, start, end, verificat)
    wb.active.title = 'Instruccions'
    write_config(wb.create_sheet('Configuració'), cfg)
    write_radiologists(wb.create_sheet('Radiòlegs'), rads)
    write_vacances(wb.create_sheet('Vacances'), year, month, start, end, festius,
                   [r.name for r in rads], inactive={r.name for r in rads if not r.active},
                   context_marks=context_marks)
    wb.active = 3
    for s in wb.worksheets:
        s.sheet_view.tabSelected = (s.title == 'Vacances')
    return wb


def build_input_from(src, year: int, month: int, start=None):
    """Entrada d'un mes nou amb la mateixa llista, capacitats, preferències i configuració."""
    wb = src if isinstance(src, openpyxl.Workbook) else openpyxl.load_workbook(
        io.BytesIO(src) if isinstance(src, (bytes, bytearray)) else src, data_only=True)
    warn = []
    rads_d = read_radiologists(wb['Radiòlegs'], warn)
    cfg = read_config(wb['Configuració'], warn)
    for n in cfg.sun_nit_only:                       # perfils especials antics -> rol
        if n in rads_d and rads_d[n].role == ROL_GUARDIES:
            rads_d[n].role = ROL_DG_NIT
    for n in cfg.weekend_day_only:
        if n in rads_d and rads_d[n].role == ROL_GUARDIES:
            rads_d[n].role = ROL_CS_DIA
    # Absències que acaben els últims dies del període d'origen -> columnes grises del nou
    marks = {}
    try:
        src_e = load_entrada(wb)
        new_start = start or period_for_month(year, month)[0]
        if src_e.period.end + dt.timedelta(days=1) == new_start:
            lo = new_start - dt.timedelta(days=max(0, DIES_REINCORPORACIO - 1))
            for name, cons in src_e.absences.items():
                for d, code in cons.items():
                    if lo <= d < new_start and code in ('V', 'B', 'G'):
                        marks.setdefault(name, {})[d] = code
    except Exception:
        pass
    return build_input(list(rads_d.values()), cfg, year, month, start=start, context_marks=marks)


def blank_input(year: int, month: int):
    cfg = Config(first_wed='')
    return build_input([], cfg, year, month)


def to_bytes(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

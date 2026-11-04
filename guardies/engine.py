"""Motor d'assignació: preferències, capacitat i equitat (CP-SAT).

Prioritats, de més a menys important:
  1. Cobrir tots els torns.
  2. Regles dures: absències, «No» de preferència, rols restringits, 1 torn al dia,
     màxim declarat, no fer guàrdia el dia d'un torn fix.
  3. Qui no ha respost: exactament 1 guàrdia cada N mesos (per defecte 2).
  4. Equitat per bosses independents (laborables / cap de setmana i festius):
     cadascú, la seva part proporcional al màxim declarat, ajustada per disponibilitat,
     amb un terra d'1 guàrdia cada N mesos i el deute acumulat dels mesos anteriors.
  5. Descans: evitar guàrdies en dies consecutius i una nit just abans d'un matí fix.
  6. Preferències «Sí».
"""
from __future__ import annotations

import datetime as dt
import math
from collections import defaultdict
from dataclasses import dataclass, field

from ortools.sat.python import cp_model

from .history import History
from .io_entrada import can_work, motiu
from .model import (K_FIX, K_NOU, K_PROP, K_SENSE, POOL_C, POOL_L, PREF_NO, PREF_SI, ROL_CS_DIA,
                    ROL_DG_NIT, Entrada, FixedAssign, Period, Slot)
from .periods import ORDINALS, fold, period_month_of

W_UNCOVERED = 1_000_000
W_SENSE = 20_000
W_FAIR = 1            # × taula 100·(a − objectiu)²
W_REST = 300
W_NIGHT_BEFORE_FIX = 300
W_PREF = 60
SEED = 2026
DIES_MES = 30.44

D = ['dl', 'dt', 'dc', 'dj', 'dv', 'ds', 'dg']


def build_slots(period: Period, cfg) -> list:
    slots = []
    for d in period.dates:
        dow, fest = d.weekday(), d in period.festius
        if fest and dow <= 2:
            slots += [Slot(d, 'dia', 'NB'), Slot(d, 'nit', 'NB')]
        elif fest and dow in (3, 4):
            slots += [Slot(d, 'dia', 'G1'), Slot(d, 'dia', 'G2'), Slot(d, 'nit', 'G1'), Slot(d, 'nit', 'G2')]
        elif dow in (0, 1):
            slots.append(Slot(d, 'nit', 'NB'))
        elif dow == 2:
            fixed_pm = any(k[0] == 2 and k[1] == '16-20' for k in cfg.fix_slots)
            if not ((Period.is_first_wed(d) and cfg.first_wed is not None) or fixed_pm):
                slots.append(Slot(d, '16-20', 'NB'))
            slots.append(Slot(d, 'nit', 'NB'))
        elif dow in (3, 4):
            if not any(k[0] == dow and k[1] == '16-20' for k in cfg.fix_slots):
                slots += [Slot(d, '16-20', 'G1'), Slot(d, '16-20', 'G2')]
            slots += [Slot(d, 'nit', 'G1'), Slot(d, 'nit', 'G2')]
        elif dow == 5:
            slots += [Slot(d, 'dia', 'G1'), Slot(d, 'dia', 'G2'), Slot(d, 'nit', 'G1'), Slot(d, 'nit', 'G2')]
        else:
            slots += [Slot(d, 'dia', 'NB'), Slot(d, 'nit', 'NB')]
    return slots


def pref_keys(s: Slot, period: Period) -> list:
    dow, fest, keys = s.dow, s.date in period.festius, []
    if fest:
        keys.append('fest_dia' if s.shift == 'dia' else 'fest_nit')
    if dow == 5:
        keys.append('ds_dia' if s.shift == 'dia' else 'ds_nit')
    elif dow == 6:
        keys.append('dg_dia' if s.shift == 'dia' else 'dg_nit')
    elif not fest:
        keys.append(f"{D[dow]}_{'tarda' if s.shift == '16-20' else 'nit'}")
    return keys


def _weekday_no_keys(s: Slot, period: Period) -> list:
    """En un festiu entre setmana, el «No» del dia de la setmana també compta
    (qui no pot fer dimarts a la nit tampoc no fa la nit d'un dimarts festiu)."""
    if s.date in period.festius and s.dow <= 4:
        return [f"{D[s.dow]}_nit"] if s.shift == 'nit' else [f"{D[s.dow]}_tarda"]
    return []


def pref_of(r, s, period):
    vals = [r.prefs.get(k) for k in pref_keys(s, period)]
    if PREF_NO in vals or any(r.prefs.get(k) == PREF_NO for k in _weekday_no_keys(s, period)):
        return PREF_NO
    if PREF_SI in vals:
        return PREF_SI
    return None


def build_fixed(e: Entrada) -> list:
    cfg, per, ab, rads = e.config, e.period, e.absences, e.radiologists
    out = []
    for d in per.dates:
        dow = d.weekday()
        if dow >= 5 or d in per.festius:
            continue
        for shift in ('8-16', '16-20'):
            if shift == '16-20' and dow == 2 and Period.is_first_wed(d) and cfg.first_wed is not None:
                entries = [('N i B', cfg.first_wed, 'fix 1r dimecres')]
            elif shift == '8-16' and dow == 4:
                k = Period.friday_ordinal(d)
                entries = [('N i B', cfg.friday.get(k, ''), f'{ORDINALS[k]} divendres')]
            else:
                entries = [(role, cfg.fix_slots[(dow, shift, role)], '') for role in ('N', 'B', 'N i B')
                           if (dow, shift, role) in cfg.fix_slots]
            for role, name, note in entries:
                if not note and not name:
                    note = 'torn fix'
                fa = FixedAssign(d, shift, role, name, note=note)
                if not name:
                    fa.status, fa.reason = 'pendent', 'sense titular a Configuració'
                elif name not in rads:
                    fa.status, fa.reason = 'pendent', f'{name} no és a Radiòlegs'
                elif not can_work(name, d, ab):
                    fa.status, fa.reason = 'substituir', motiu(name, d, ab)
                out.append(fa)
    return out


@dataclass
class Result:
    entrada: Entrada
    slots: list
    assignments: dict                 # Slot -> nom
    fixed: list
    uncovered: list
    shares: dict                      # nom -> {'L','C'} part proporcional del període
    debts: dict                       # nom -> {'L','C'} deute en començar
    targets: dict                     # nom -> {'L','C'} objectiu (part + deute)
    due: set                          # «sense resposta» a qui toca aquest mes
    status: str = ''
    seconds: float = 0.0
    warnings: list = field(default_factory=list)
    last_guard_before: dict = field(default_factory=dict)

    def count(self, name, pool=None) -> int:
        per = self.entrada.period
        return sum(1 for s, n in self.assignments.items()
                   if n == name and (pool is None or per.pool(s.date) == pool))


def _due_sense(e: Entrada, hist: History | None) -> tuple:
    """Qui no ha respost fa 1 guàrdia cada N mesos. A qui li toca aquest mes?

    Amb historial de l'app (pestanya Historic): li toca si fa N mesos o més de la seva última guàrdia.
    Sense historial (primer mes, o planning de l'app anterior): es reparteixen en N grups
    alternats per ordre alfabètic, perquè no els toqui a tots el mateix mes.
    """
    S = sorted((r.name for r in e.active(K_SENSE)), key=fold)
    N = max(1, e.config.min_every_months)
    cur = e.period.year * 12 + e.period.month
    if hist is None or not hist.has_hist_sheet:
        due = {n for i, n in enumerate(S) if i % N == cur % N}
        return due, ("Primer mes amb aquest sistema: els que no han respost es reparteixen en " + str(N) +
                     " grups alternats (aquest mes en fan " + str(len(due)) + " de " + str(len(S)) + ").")
    due = set()
    for n in S:
        last = hist.last_guard(n)
        if last is None:
            due.add(n)
        else:
            y, m = period_month_of(last)
            if cur - (y * 12 + m) >= N:
                due.add(n)
    return due, ''


def solve(e: Entrada, hist: History | None = None, time_limit: float = 20.0) -> Result:
    import time
    t0 = time.time()
    per, cfg, ab, rads = e.period, e.config, e.absences, e.radiologists
    warnings = []
    slots = build_slots(per, cfg)
    fixed = build_fixed(e)
    busy = defaultdict(set)                      # nom -> dies amb torn fix
    busy_morning = defaultdict(set)              # nom -> dies amb torn fix de matí
    for fa in fixed:
        if fa.status == 'ok':
            busy[fa.name].add(fa.date)
            if fa.shift == '8-16':
                busy_morning[fa.name].add(fa.date)
    nxt = per.end + dt.timedelta(days=1)         # dilluns del període següent
    for (dow, shift, role), n in cfg.fix_slots.items():
        if dow == 0 and shift == '8-16' and n:
            busy_morning[n].add(nxt)

    # --- Nou-incorporats: dimecres 16-20 (no el 1r dimecres) ---
    assignments = {}
    nous = sorted((r for r in e.active(K_NOU)), key=lambda r: fold(r.name))
    nou_count = defaultdict(int)
    for s in slots:
        if not (s.dow == 2 and s.shift == '16-20' and s.date not in per.festius):
            continue
        cands = [r for r in nous if can_work(r.name, s.date, ab) and s.date not in busy[r.name]
                 and not Period.is_first_wed(s.date) and pref_of(r, s, per) != PREF_NO
                 and (r.max_m is None or nou_count[r.name] < max(1, int(math.floor(r.max_m + 1e-9))))]
        if cands:
            prev = (lambda n: (hist.guards.get(n, {}).get(POOL_L, 0) if hist else 0))
            r = min(cands, key=lambda r: (nou_count[r.name], prev(r.name), fold(r.name)))
            assignments[s] = r.name
            nou_count[r.name] += 1
            busy[r.name].add(s.date)
    model_slots = [s for s in slots if s not in assignments]

    # --- Elegibilitat ---
    people = [r for r in e.active() if r.klass in (K_PROP, K_SENSE)]
    elig = {}
    for r in people:
        for s in model_slots:
            if r.role == ROL_DG_NIT and not (s.dow == 6 and s.shift == 'nit'):
                continue
            if r.role == ROL_CS_DIA and not (s.dow >= 5 and s.shift == 'dia'):
                continue
            if s.date in busy[r.name] or not can_work(r.name, s.date, ab):
                continue
            p = pref_of(r, s, per)
            if p == PREF_NO:
                continue
            elig[(r.name, s)] = p

    # --- Objectius ---
    due, note = _due_sense(e, hist)
    if note:
        warnings.append(note)
    S_names = {r.name for r in people if r.klass == K_SENSE}
    due &= S_names
    no_slot = sorted(n for n in due if not any(k[0] == n for k in elig))
    P = [r for r in people if r.klass == K_PROP]
    n_tot = max(1, len(model_slots))
    shares = {r.name: {POOL_L: 0.0, POOL_C: 0.0} for r in P}
    avail = {r.name: {POOL_L: 0.0, POOL_C: 0.0} for r in P}
    demand = {}
    for k in (POOL_L, POOL_C):
        sk = [s for s in model_slots if per.pool(s.date) == k]
        nk = len(sk)
        demand[k] = max(0.0, nk - (len(due) - len(no_slot)) * nk / n_tot)
        if not nk:
            continue
        w = {}
        for r in P:
            a = sum(1 for s in sk if (r.name, s) in elig) / nk
            avail[r.name][k] = a
            w[r.name] = (r.max_m or 0) * a
        W = sum(w.values())
        for r in P:
            shares[r.name][k] = demand[k] * w[r.name] / W if W > 0 else 0.0
    # Terra: ningú que faci guàrdies per sota d'1 cada N mesos
    floor = per.n_days / DIES_MES / max(1, cfg.min_every_months)
    floored = set()
    for r in P:
        tot = shares[r.name][POOL_L] + shares[r.name][POOL_C]
        av = avail[r.name][POOL_L] + avail[r.name][POOL_C]
        if (r.max_m or 0) > 0 and av > 0 and tot < floor:
            f = min(floor, max(1.0, r.max_m))
            if tot > 0:
                shares[r.name] = {k: v * f / tot for k, v in shares[r.name].items()}
            else:
                shares[r.name] = {k: f * avail[r.name][k] / av for k in (POOL_L, POOL_C)}
            floored.add(r.name)
    for k in (POOL_L, POOL_C):
        fixed_part = sum(shares[n][k] for n in floored)
        rest = sum(shares[r.name][k] for r in P if r.name not in floored)
        if rest > 0:
            f = max(0.0, demand[k] - fixed_part) / rest
            for r in P:
                if r.name not in floored:
                    shares[r.name][k] *= f
    debts = {r.name: {k: (hist.debt_now(r.name, k) if hist else 0.0) for k in (POOL_L, POOL_C)} for r in P}
    for r in e.radiologists.values():           # inactius de classe P: es conserva el deute per al mes que tornin
        if not r.active and r.klass == K_PROP and hist and r.name in hist.share:
            debts[r.name] = {k: hist.debt_now(r.name, k) for k in (POOL_L, POOL_C)}
            shares[r.name] = {POOL_L: 0.0, POOL_C: 0.0}
    targets = {r.name: {k: max(0.0, shares[r.name][k] + debts[r.name][k]) for k in (POOL_L, POOL_C)} for r in P}

    # --- Model ---
    m = cp_model.CpModel()
    x = {key: m.NewBoolVar(f"x{i}") for i, key in enumerate(elig)}
    u = {s: m.NewBoolVar(f"u{i}") for i, s in enumerate(model_slots)}
    by_slot, by_person = defaultdict(list), defaultdict(list)
    for (n, s), v in x.items():
        by_slot[s].append(v)
        by_person[n].append((s, v))
    for s in model_slots:
        m.Add(sum(by_slot[s]) + u[s] == 1)
    day_expr = {}
    for n, lst in by_person.items():
        per_day = defaultdict(list)
        for s, v in lst:
            per_day[s.date].append(v)
        for d, vs in per_day.items():
            if len(vs) > 1:
                m.Add(sum(vs) <= 1)
            day_expr[(n, d)] = sum(vs)
    obj = [W_UNCOVERED * v for v in u.values()]
    for r in people:
        lst = by_person.get(r.name, [])
        tot = sum(v for _, v in lst) if lst else 0
        if r.klass == K_SENSE:
            if lst:
                m.Add(tot <= 1)
                if r.name in due:
                    obj.append(W_SENSE * (1 - tot))
                else:
                    obj.append(2 * W_SENSE * tot)
            continue
        cap = max(1, int(math.floor(r.max_m + 1e-9)))
        if lst:
            m.Add(tot <= cap)
        for k in (POOL_L, POOL_C):
            vs = [v for s, v in lst if per.pool(s.date) == k]
            hi = min(cap, len(vs))
            T = targets[r.name][k]
            if not vs:
                continue
            a = m.NewIntVar(0, hi, f"A_{k}_{len(obj)}")
            m.Add(a == sum(vs))
            table = [int(round(100 * (i - T) ** 2)) for i in range(hi + 1)]
            c = m.NewIntVar(0, max(table), f"C_{k}_{len(obj)}")
            m.AddElement(a, table, c)
            obj.append(W_FAIR * c)
    # Descans
    R = max(0, cfg.rest_days)
    for r in people:
        last = hist.last_guard(r.name) if hist else None
        for d in per.dates:
            yd = day_expr.get((r.name, d))
            if yd is None:
                continue
            for g in range(1, R + 1):
                y2 = day_expr.get((r.name, d + dt.timedelta(days=g)))
                if y2 is not None:
                    b = m.NewBoolVar('')
                    m.Add(yd + y2 - 1 <= b)
                    obj.append(W_REST * b)
            if last is not None and 0 < (d - last).days <= R:
                obj.append(W_REST * yd)
        # nit just abans d'un torn fix del matí
        for s, v in by_person.get(r.name, []):
            if s.shift == 'nit' and (s.date + dt.timedelta(days=1)) in busy_morning[r.name]:
                obj.append(W_NIGHT_BEFORE_FIX * v)
    # Preferències
    for key, p in elig.items():
        if p == PREF_SI:
            obj.append(-W_PREF * x[key])
    m.Minimize(sum(obj))

    solver = cp_model.CpSolver()
    solver.parameters.num_workers = 8
    solver.parameters.interleave_search = True
    solver.parameters.max_deterministic_time = float(time_limit)
    solver.parameters.max_time_in_seconds = max(60.0, 6 * float(time_limit))
    solver.parameters.random_seed = SEED
    status = solver.Solve(m)
    st = solver.StatusName(status)
    uncovered = []
    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for (n, s), v in x.items():
            if solver.Value(v):
                assignments[s] = n
        uncovered = [s for s in model_slots if solver.Value(u[s])]
    else:
        uncovered = list(model_slots)
        warnings.append(f"El càlcul no ha trobat solució ({st}).")
    if no_slot:
        warnings.append("Sense resposta a qui tocava guàrdia però no pot fer-ne cap aquest període: "
                        + ', '.join(no_slot) + ". Li tocarà el mes que ve.")

    res = Result(entrada=e, slots=slots, assignments=assignments, fixed=fixed, uncovered=uncovered,
                 shares=shares, debts=debts, targets=targets, due=due, status=st,
                 seconds=time.time() - t0, warnings=warnings)
    res.last_guard_before = {r.name: hist.last_guard(r.name) for r in e.radiologists.values()} if hist else {}
    _suggest(res, busy)
    return res


def _suggest(res: Result, busy) -> None:
    """Candidats per als torns fixos que cal substituir o que no tenen titular."""
    e = res.entrada
    per, ab, cfg = e.period, e.absences, e.config
    guard_days = defaultdict(set)
    for s, n in res.assignments.items():
        guard_days[n].add(s.date)
    load = {n: res.count(n) / max(0.5, sum(res.targets.get(n, {}).values()) or 0.5) for n in e.radiologists}
    for fa in res.fixed:
        if fa.status == 'ok':
            continue
        cands = []
        for r in e.active():
            if r.name == fa.name or fa.date in busy[r.name] or fa.date in guard_days[r.name]:
                continue
            if not can_work(r.name, fa.date, ab):
                continue
            if fa.shift == '8-16':
                if r.mati == PREF_NO:
                    continue
                pri = 0 if r.mati == PREF_SI else 1
            else:
                key = f"{D[fa.date.weekday()]}_tarda"
                if r.prefs.get(key) == PREF_NO:
                    continue
                pri = 0 if r.prefs.get(key) == PREF_SI else 1
            if fa.date.weekday() == 4 and fa.shift == '8-16' and r.name == cfg.friday.get(5):
                pri = -1
            dies = fold(r.dies_puntuals)
            dia_nom = ['dilluns', 'dimarts', 'dimecres', 'dijous', 'divendres'][fa.date.weekday()]
            dia_es = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes'][fa.date.weekday()]
            if dia_nom in dies or dia_es in dies:
                pri -= 0.5
            cands.append((pri, round(load.get(r.name, 0), 2), fold(r.name), r.name))
        fa.suggestions = [c[3] for c in sorted(cands)[:3]]

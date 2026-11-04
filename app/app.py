"""Planning de guàrdies de Telediagnòstic per preferències i capacitat — interfície Streamlit."""
from __future__ import annotations

import datetime as dt
import hashlib
import io
import os
import re
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

_here = Path(__file__).resolve().parent
ROOT = _here if (_here / 'guardies').is_dir() else _here.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from guardies import __version__  # noqa: E402
from guardies import demo as gdemo  # noqa: E402
from guardies import engine, survey, templates, writer  # noqa: E402
from guardies.history import parse_planning  # noqa: E402
from guardies.io_entrada import load_entrada  # noqa: E402
from guardies.model import K_FIX, K_NOU, K_PROP, K_SENSE, POOL_C, POOL_L  # noqa: E402
from guardies.periods import fmt_dia, fold, month_label, next_month, period_for_month, prev_month  # noqa: E402

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
st.set_page_config(page_title="Planning de guàrdies TD", page_icon="🩺", layout="wide")


def _secret(name, default=None):
    try:
        v = st.secrets.get(name)
        if v:
            return v
    except Exception:
        pass
    return os.environ.get(name, default)


def _gate():
    pwd = _secret('APP_PASSWORD')
    if not pwd or st.session_state.get('auth_ok'):
        return
    st.title("🩺 Planning de guàrdies TD")
    v = st.text_input("Contrasenya", type="password")
    if v:
        if v == pwd:
            st.session_state['auth_ok'] = True
            st.rerun()
        else:
            st.error("Contrasenya incorrecta.")
    st.stop()


def _bytes(wb) -> bytes:
    return templates.to_bytes(wb)


def _parse_festius(text, period):
    out, avisos = set(), []
    for tok in re.split(r'[,;\s]+', text or ''):
        if not tok:
            continue
        m = re.match(r'^(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?$', tok)
        if not m:
            avisos.append(f"Festiu no reconegut: «{tok}» (format dd/mm)")
            continue
        dd, mm = int(m.group(1)), int(m.group(2))
        c = [d for d in period.dates if d.day == dd and d.month == mm]
        if c:
            out.add(c[0])
        else:
            avisos.append(f"El festiu {tok} no cau dins del període")
    return out, avisos


_gate()
st.title("🩺 Planning de guàrdies TD — preferències i capacitat")
st.caption(f"v{__version__} · Sense roda: cada guàrdia va a qui li toca segons la capacitat declarada, les preferències "
           "i el que ja ha fet els mesos anteriors. Setmanes completes (dl → dg).")

with st.sidebar:
    st.header("Cada mes")
    st.markdown("""
1. **Entrada del mes**: la que et va donar l'app el mes passat (o una plantilla). Revisa *Actiu* i marca les absències a *Vacances*.
2. **Planning del mes anterior**: l'últim fitxer, amb els canvis que s'hi hagin fet a mà.
3. **Genera** i revisa la pestanya *Pendents*.
4. **Publica** el planning el dia 1.
5. Desa'l: serà el *planning anterior* del mes que ve.
""")
    st.divider()
    hoy = dt.date.today()
    ym = next_month(*next_month(hoy.year, hoy.month)) if hoy.day > 1 else next_month(hoy.year, hoy.month)
    st.download_button("📄 Plantilla d'entrada buida", data=_bytes(templates.blank_input(*ym)),
                       file_name=f"Entrada_buida_{ym[0]}-{ym[1]:02d}.xlsx", mime=XLSX, width='stretch')
    if st.button("🧪 Provar amb dades fictícies", width='stretch'):
        st.session_state['demo'] = True

# ------------------------------------------------------------------
# Demo
# ------------------------------------------------------------------
if st.session_state.get('demo'):
    y, m = next_month(dt.date.today().year, dt.date.today().month)
    py, pm = prev_month(y, m)
    prev_in = gdemo.demo_input(py, pm)
    e0 = load_entrada(_bytes(prev_in))
    prev_plan = _bytes(writer.write_planning(engine.solve(e0, None)))
    st.session_state['demo_in'] = _bytes(gdemo.demo_input(y, m))
    st.session_state['demo_prev'] = prev_plan
    st.session_state['demo'] = False
    st.info("S'han carregat dades fictícies (entrada i planning anterior). Pots descarregar-les aquí sota.")

if st.session_state.get('demo_in'):
    c1, c2 = st.columns(2)
    c1.download_button("Demo · entrada", st.session_state['demo_in'], "DEMO_Entrada.xlsx", XLSX)
    c2.download_button("Demo · planning anterior", st.session_state['demo_prev'], "DEMO_Planning_anterior.xlsx", XLSX)

# ------------------------------------------------------------------
# Pas 1 — Entrada
# ------------------------------------------------------------------
st.header("Pas 1 — Entrada del mes")
f_in = st.file_uploader("Fitxer d'entrada (Configuració + Radiòlegs + Vacances)", type=['xlsx'], key='f_in')
in_bytes = f_in.getvalue() if f_in else st.session_state.get('demo_in')
e = None
if in_bytes:
    try:
        e = load_entrada(in_bytes)
        cls = {k: len(e.active(k)) for k in (K_PROP, K_SENSE, K_FIX, K_NOU)}
        st.success(f"**{e.period.label}** · {e.period.text()} ({e.period.n_weeks} setmanes) · "
                   f"{cls[K_PROP]} amb capacitat declarada, {cls[K_SENSE]} sense resposta (1 cada 2 mesos), "
                   f"{cls[K_FIX]} només fixos, {cls[K_NOU]} nou-incorporats · "
                   f"{sum(1 for r in e.radiologists.values() if not r.active)} no actius.")
        for w in e.warnings:
            (st.info if 'dia pont' in w else st.warning)(w)
    except Exception as ex:
        st.error(f"No s'ha pogut llegir l'entrada: {ex}")
        e = None

# ------------------------------------------------------------------
# Pas 2 — Planning anterior
# ------------------------------------------------------------------
h = None
prev_bytes = None
if e:
    st.header("Pas 2 — Planning del mes anterior")
    f_prev = st.file_uploader("Planning del període anterior (el publicat, amb els canvis manuals)", type=['xlsx'], key='f_prev')
    prev_bytes = f_prev.getvalue() if f_prev else (st.session_state.get('demo_prev') if not f_in else None)
    if prev_bytes:
        try:
            h = parse_planning(prev_bytes, list(e.radiologists), e.config)
            n_g = sum(v[POOL_L] + v[POOL_C] for v in h.guards.values())
            st.success(f"Planning anterior: {fmt_dia(h.start, True)} → {fmt_dia(h.end, True)} · {n_g} guàrdies llegides "
                       f"de {len(h.guards)} persones · " + ("amb historial de l'app (equitat encadenada)." if h.has_hist_sheet
                                                             else "sense historial (planning de l'app anterior): l'equitat comença ara."))
            if h.end and h.end + dt.timedelta(days=1) != e.period.start:
                st.warning(f"El planning anterior acaba el {fmt_dia(h.end, True)} i l'entrada comença el "
                           f"{fmt_dia(e.period.start, True)}: no són períodes seguits.")
            if h.unmatched:
                st.warning("Noms del planning anterior que no es poden casar amb Radiòlegs (no compten): "
                           + ', '.join(h.unmatched))
            for w in h.warnings:
                st.warning(w)
        except Exception as ex:
            st.error(f"No s'ha pogut llegir el planning anterior: {ex}")
            h = None
    else:
        st.info("Sense planning anterior: tothom comença de zero i els que no han respost es reparteixen en dos grups alternats.")

# ------------------------------------------------------------------
# Pas 3 — Revisió i generació
# ------------------------------------------------------------------
if e:
    st.header("Pas 3 — Revisió i generació")
    fest_txt = st.text_input("Festius dins del període (dd/mm, separats per comes)", value=e.period.festius_text())
    fest, avisos = _parse_festius(fest_txt, e.period)
    for a in avisos:
        st.warning(a)
    e.period.festius = fest
    due, note = engine._due_sense(e, h)
    with st.expander(f"Qui no ha respost i li toca guàrdia aquest mes ({len(due)})"):
        st.write(', '.join(sorted(due, key=fold)) or 'Ningú.')
        if note:
            st.caption(note)
    key = hashlib.sha1((in_bytes or b'') + (prev_bytes or b'') + fest_txt.encode()).hexdigest()
    ok_sense = True
    if h is None:
        ok_sense = st.checkbox("No tinc el planning del mes anterior: genera'l igualment (el deute de cadascú torna a zero)")
    if st.button("🚀 Generar planning", type='primary', width='stretch', disabled=not ok_sense):
        with st.spinner("Calculant (uns segons)..."):
            res = engine.solve(e, h)
            plan = _bytes(writer.write_planning(res))
            ny, nm = next_month(e.period.year, e.period.month)
            nxt = _bytes(templates.build_input_from(in_bytes, ny, nm))
            rows = []
            for r in sorted(e.radiologists.values(), key=lambda r: fold(r.name)):
                if not r.active or r.klass == K_FIX:
                    continue
                t = res.targets.get(r.name)
                rows.append({'Radiòleg': r.name, 'Classe': {K_PROP: 'Proporcional', K_SENSE: 'Sense resposta',
                                                             K_NOU: 'Nou-incorporat'}[r.klass],
                             'Màx./mes': r.max_m, 'Objectiu': round(t[POOL_L] + t[POOL_C], 1) if t else (1 if r.name in res.due else 0),
                             'Laborables': res.count(r.name, POOL_L), 'Cap de setmana i festius': res.count(r.name, POOL_C)})
            st.session_state['res'] = dict(
                key=key, label=e.period.label, code=e.period.code, n=len(res.slots), unc=len(res.uncovered),
                pend=sum(1 for f in res.fixed if f.status != 'ok'), status=res.status, secs=res.seconds,
                warnings=res.warnings, plan=plan, nxt=nxt, nxt_label=month_label(ny, nm), nxt_code=f"{ny}-{nm:02d}",
                rows=rows)
    R = st.session_state.get('res')
    if R and R['key'] == key:
        st.divider()
        st.subheader(f"Resultat — {R['label']}")
        c1, c2, c3 = st.columns(3)
        c1.metric("Guàrdies cobertes", f"{R['n'] - R['unc']} / {R['n']}")
        c2.metric("Torns fixos per resoldre", R['pend'])
        c3.metric("Càlcul", f"{R['status'].lower()} · {R['secs']:.0f} s")
        if R['unc']:
            st.error(f"{R['unc']} guàrdies sense cobrir: vegeu la pestanya Pendents.")
        for w in R['warnings']:
            st.info(w)
        d1, d2 = st.columns(2)
        d1.download_button(f"📥 Planning {R['label']}", R['plan'], f"Planning_{R['code']}.xlsx", XLSX,
                           width='stretch', key='dl_plan')
        d2.download_button(f"📥 Entrada de {R['nxt_label']} ja preparada", R['nxt'], f"Entrada_{R['nxt_code']}.xlsx",
                           XLSX, width='stretch', key='dl_next',
                           help="Mateixos radiòlegs, capacitats, preferències i configuració. Només cal marcar-hi les absències.")
        with st.expander("Guàrdies per persona"):
            df = pd.DataFrame(R['rows']).fillna('')
            st.dataframe(df.astype(str), hide_index=True, width='stretch')

# ------------------------------------------------------------------
# Eines
# ------------------------------------------------------------------
if e and in_bytes:
    st.divider()
    with st.expander("📅 Preparar l'entrada d'un altre mes"):
        opts, ym = [], (e.period.year, e.period.month)
        for _ in range(18):
            ym = next_month(*ym)
            opts.append(ym)
        tria = st.selectbox("Mes", opts, format_func=lambda t: month_label(*t))
        s_, e_ = period_for_month(*tria)
        st.caption(f"Període: {fmt_dia(s_, True)} → {fmt_dia(e_, True)}")
        st.download_button("📥 Descarregar", _bytes(templates.build_input_from(in_bytes, *tria)),
                           f"Entrada_{tria[0]}-{tria[1]:02d}.xlsx", XLSX, key='dl_altre')
    with st.expander("📋 Actualitzar capacitats i preferències des d'una enquesta"):
        st.caption("Taula amb les columnes Radiólogo, Mín./mes, Máx./mes, Lun tarde … Fest noche, Mañana SÍ/NO i "
                   "Días puntuales. Es casa amb la llista per cognoms.")
        f_s = st.file_uploader("Taula de l'enquesta", type=['xlsx'], key='f_survey')
        if f_s:
            import openpyxl
            from guardies.io_entrada import read_config, read_radiologists
            wb = openpyxl.load_workbook(io.BytesIO(in_bytes), data_only=True)
            w = []
            rads = read_radiologists(wb['Radiòlegs'], w)
            cfg = read_config(wb['Configuració'], w)
            merged, rep = survey.merge(rads, survey.read_survey(f_s.getvalue()))
            canviats = [f"{a} → {b}" for a, b in rep['casats'] if a != b]
            st.write(f"Casats: {len(rep['casats'])} · nous: {len(rep['nous'])} · sense resposta: {len(rep['sense_resposta'])}")
            if rep['ambigus']:
                st.error("Respostes que poden ser de més d'una persona (no s'han aplicat): " + ', '.join(rep['ambigus']))
            if rep['duplicats']:
                st.warning("Respostes repetides (s'ha fet servir la primera): " + ', '.join(a for a, _ in rep['duplicats']))
            if rep['nous']:
                st.warning("No són a la llista (s'hi afegeixen; completeu nom i correu): " + ', '.join(rep['nous']))
            if canviats:
                with st.expander("Com s'han casat els noms"):
                    st.text('\n'.join(canviats))
            wb_new = templates.build_input(sorted(merged.values(), key=lambda r: fold(r.name)), cfg,
                                           e.period.year, e.period.month, start=e.period.start, end=e.period.end)
            st.download_button("📥 Entrada amb les preferències actualitzades", _bytes(wb_new),
                               f"Entrada_{e.period.code}_preferencies.xlsx", XLSX, key='dl_survey')
            st.caption("Atenció: aquest fitxer té la pestanya Vacances buida; marca-hi les absències.")

st.divider()
st.caption("Les dades es processen només durant la sessió i no es desen enlloc.")

"""Línia d'ordres.

  python -m guardies generar ENTRADA.xlsx [--anterior PLANNING.xlsx] [--sortida Planning.xlsx]
  python -m guardies plantilla ENTRADA.xlsx --mes 2027-02 [--sortida Entrada_2027-02.xlsx]
  python -m guardies buida --mes 2027-02 [--sortida Entrada_buida.xlsx]
  python -m guardies enquesta ENTRADA.xlsx TAULA.xlsx [--sortida Entrada_preferencies.xlsx]
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

import openpyxl

from . import engine, survey, templates, writer
from .history import parse_planning
from .io_entrada import load_entrada, read_config, read_radiologists
from .model import K_SENSE
from .periods import fold


def _ym(s):
    y, m = s.split('-')
    return int(y), int(m)


def main(argv=None):
    ap = argparse.ArgumentParser(prog='guardies', description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    g = sub.add_parser('generar')
    g.add_argument('entrada')
    g.add_argument('--anterior')
    g.add_argument('--sortida')
    p = sub.add_parser('plantilla')
    p.add_argument('entrada')
    p.add_argument('--mes', required=True)
    p.add_argument('--sortida')
    b = sub.add_parser('buida')
    b.add_argument('--mes', required=True)
    b.add_argument('--sortida')
    s = sub.add_parser('enquesta')
    s.add_argument('entrada')
    s.add_argument('taula')
    s.add_argument('--sortida')
    a = ap.parse_args(argv)

    if a.cmd == 'generar':
        e = load_entrada(a.entrada)
        h = parse_planning(a.anterior, list(e.radiologists), e.config) if a.anterior else None
        if h is not None:
            if h.end and h.end + dt.timedelta(days=1) != e.period.start:
                print(f"  ! El planning anterior acaba el {h.end:%d/%m/%Y} i aquest període comença el "
                      f"{e.period.start:%d/%m/%Y}: no són seguits.")
            if h.unmatched:
                print('  ! Noms del planning anterior que no es poden casar (no compten):', ', '.join(h.unmatched))
            for w in h.warnings:
                print('  !', w)
        else:
            print("  ! Sense planning anterior: el deute comença a zero i els que no han respost es reparteixen en dos grups.")
        res = engine.solve(e, h)
        out = a.sortida or f"Planning_{e.period.code}.xlsx"
        writer.write_planning(res, out)
        print(f"{e.period.label}: {e.period.text()} · {len(res.slots) - len(res.uncovered)}/{len(res.slots)} guàrdies cobertes "
              f"· {sum(1 for f in res.fixed if f.status != 'ok')} fixos per resoldre · {res.status} {res.seconds:.1f}s → {out}")
        for w in e.warnings + res.warnings:
            print('  ·', w)
        return 0 if not res.uncovered else 1
    if a.cmd == 'plantilla':
        y, m = _ym(a.mes)
        out = a.sortida or f"Entrada_{y}-{m:02d}.xlsx"
        templates.build_input_from(a.entrada, y, m).save(out)
        print('→', out)
    elif a.cmd == 'buida':
        y, m = _ym(a.mes)
        out = a.sortida or f"Entrada_buida_{y}-{m:02d}.xlsx"
        templates.blank_input(y, m).save(out)
        print('→', out)
    elif a.cmd == 'enquesta':
        wb = openpyxl.load_workbook(a.entrada, data_only=True)
        w = []
        rads, cfg = read_radiologists(wb['Radiòlegs'], w), read_config(wb['Configuració'], w)
        e = load_entrada(a.entrada)
        merged, rep = survey.merge(rads, survey.read_survey(a.taula))
        src = Path(a.entrada)
        out = a.sortida or str(src.with_name(src.stem + '_preferencies.xlsx'))
        if Path(out).resolve() == src.resolve():
            raise SystemExit("La sortida no pot ser el mateix fitxer d'entrada.")
        for k in ('ambigus', 'duplicats'):
            if rep[k]:
                print(f"  ! {k}: {rep[k]} (no s'han aplicat; reviseu-los a mà)")
        templates.build_input(sorted(merged.values(), key=lambda r: fold(r.name)), cfg, e.period.year, e.period.month,
                              start=e.period.start, end=e.period.end).save(out)
        print(f"casats {len(rep['casats'])}, nous {rep['nous']}, sense resposta {rep['sense_resposta']} → {out}")
    return 0


if __name__ == '__main__':
    sys.exit(main())

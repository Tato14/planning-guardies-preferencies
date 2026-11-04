"""Dades fictícies per provar l'app (cap nom real)."""
from __future__ import annotations

import random

from .model import (ROL_CS_DIA, ROL_DG_NIT, ROL_FIX, ROL_FIX_GUARDIES, ROL_GUARDIES, ROL_NOU, Config,
                    PREF_KEYS, PREF_NO, PREF_SI, Radiologist)
from .templates import build_input

COGNOMS = ['Alfa', 'Beta', 'Gamma', 'Delta', 'Èpsilon', 'Zeta', 'Eta', 'Teta', 'Iota', 'Kappa', 'Lambda',
           'Mi', 'Ni', 'Xi', 'Òmicron', 'Pi', 'Ro', 'Sigma', 'Tau', 'Ípsilon', 'Fi', 'Khi', 'Psi', 'Omega',
           'Ponent', 'Llevant', 'Migjorn', 'Tramuntana', 'Garbí', 'Mestral', 'Xaloc', 'Gregal', 'Cerç',
           'Marinada', 'Terral', 'Vendaval', 'Morró', 'Brisa', 'Calma', 'Ruixat']
NOMS = ['Anna', 'Bernat', 'Carla', 'David', 'Elena', 'Ferran', 'Gemma', 'Hug', 'Irene', 'Jordi', 'Laia',
        'Marc', 'Núria', 'Oriol', 'Paula', 'Quim', 'Rosa', 'Sergi', 'Teresa', 'Uri']


def demo_radiologists(seed=7):
    rnd = random.Random(seed)
    names = [f"{c}, {NOMS[i % len(NOMS)]}" for i, c in enumerate(COGNOMS)]
    rads = []
    for i, n in enumerate(names):
        if i < 4:
            r = Radiologist(n, ROL_FIX, mati=PREF_SI)
        elif i < 6:
            r = Radiologist(n, ROL_FIX_GUARDIES, min_m=4, max_m=8, mati=PREF_SI)
        elif i == 6:
            r = Radiologist(n, ROL_NOU, min_m=2, max_m=3)
        elif i == 7:
            r = Radiologist(n, ROL_DG_NIT, min_m=1, max_m=1)
        elif i == 8:
            r = Radiologist(n, ROL_CS_DIA, min_m=1, max_m=1)
        elif i < 32:
            mx = rnd.choice([1, 2, 3, 4, 4, 5, 6, 6, 8, 12])
            r = Radiologist(n, ROL_GUARDIES, min_m=max(0, mx - 2), max_m=mx, mati=rnd.choice([PREF_SI, PREF_NO, None]))
            for k in rnd.sample(PREF_KEYS, rnd.choice([0, 0, 2, 4, 6])):
                r.prefs[k] = PREF_SI if rnd.random() < 0.75 else PREF_NO
            if rnd.random() < 0.15:
                r.observacions = rnd.choice(['Prefereix no fer nits seguides.', 'Només amb antelació.',
                                             'Disponible per a substitucions de matí els dijous.'])
        else:
            r = Radiologist(n, ROL_GUARDIES)          # no ha respost
        rads.append(r)
    return rads


def demo_config(rads):
    n = [r.name for r in rads]
    cfg = Config()
    cfg.fix_slots = {(0, '8-16', 'N'): n[0], (0, '8-16', 'B'): n[1], (0, '16-20', 'N'): n[2], (0, '16-20', 'B'): n[1],
                     (1, '8-16', 'N'): n[3], (1, '8-16', 'B'): n[4], (1, '16-20', 'N'): n[3], (1, '16-20', 'B'): n[5],
                     (2, '8-16', 'N'): n[0], (2, '8-16', 'B'): n[1], (3, '8-16', 'N'): n[2], (3, '8-16', 'B'): n[4]}
    cfg.first_wed = n[5]
    cfg.friday = {1: n[9], 2: n[10], 3: n[11], 4: n[12], 5: n[5]}
    return cfg


def demo_input(year, month, seed=7):
    rads = demo_radiologists(seed)
    return build_input(rads, demo_config(rads), year, month)

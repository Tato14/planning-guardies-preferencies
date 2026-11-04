"""Estructures de dades del planning de guàrdies v4."""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

from .periods import DIES_CA, fmt_dia, month_label, weekday_ordinal_in_month

# --- Rols ---------------------------------------------------------------
ROL_GUARDIES = 'guàrdies'
ROL_FIX = 'fix'
ROL_FIX_GUARDIES = 'fix-i-guàrdies'
ROL_NOU = 'nou-incorporat'
ROL_DG_NIT = 'només-dg-nit'
ROL_CS_DIA = 'només-cap-setmana-dia'
ROLS = (ROL_GUARDIES, ROL_FIX, ROL_FIX_GUARDIES, ROL_NOU, ROL_DG_NIT, ROL_CS_DIA)

ROL_ALIES = {   # noms antics (v1-v3) i variants
    'rotador': ROL_GUARDIES, 'guardies': ROL_GUARDIES, 'guàrdies': ROL_GUARDIES,
    'fix-nomes': ROL_FIX, 'fix-només': ROL_FIX, 'fix': ROL_FIX,
    'fix-i-rota': ROL_FIX_GUARDIES, 'fix-i-guardies': ROL_FIX_GUARDIES, 'fix-i-guàrdies': ROL_FIX_GUARDIES,
    'nou-incorporat': ROL_NOU, 'nou incorporat': ROL_NOU,
    'sun-nit-only': ROL_DG_NIT, 'nomes-dg-nit': ROL_DG_NIT, 'només-dg-nit': ROL_DG_NIT,
    'weekend-day-only': ROL_CS_DIA, 'nomes-cap-setmana-dia': ROL_CS_DIA, 'només-cap-setmana-dia': ROL_CS_DIA,
}

# Classes d'assignació
K_FIX = 'FIX'        # només torns fixos
K_PROP = 'P'         # guàrdies proporcionals al màxim declarat
K_SENSE = 'S'        # no ha respost: 1 guàrdia cada N mesos
K_NOU = 'NOU'        # nou-incorporat: dimecres 16-20

# --- Preferències per tipus de torn -------------------------------------
PREF_KEYS = ['dl_tarda', 'dl_nit', 'dt_tarda', 'dt_nit', 'dc_tarda', 'dc_nit', 'dj_tarda', 'dj_nit',
             'dv_tarda', 'dv_nit', 'ds_dia', 'ds_nit', 'dg_dia', 'dg_nit', 'fest_dia', 'fest_nit']
PREF_HEADERS = {   # capçalera de la pestanya Radiòlegs
    'dl_tarda': 'Dl tarda', 'dl_nit': 'Dl nit', 'dt_tarda': 'Dt tarda', 'dt_nit': 'Dt nit',
    'dc_tarda': 'Dc tarda', 'dc_nit': 'Dc nit', 'dj_tarda': 'Dj tarda', 'dj_nit': 'Dj nit',
    'dv_tarda': 'Dv tarda', 'dv_nit': 'Dv nit', 'ds_dia': 'Ds dia', 'ds_nit': 'Ds nit',
    'dg_dia': 'Dg dia', 'dg_nit': 'Dg nit', 'fest_dia': 'Festiu dia', 'fest_nit': 'Festiu nit',
}
PREF_SI, PREF_NO = 'S', 'N'

# --- Absències -----------------------------------------------------------
CODIS_FULL = ('V', 'B', 'C', 'G', 'X')
CODI_INACTIU = 'I'
CODI_REINCORPORACIO = 'R'
CODIS_AMB_REINCORPORACIO = ('V', 'B')
DIES_REINCORPORACIO = 3
CODIS_BLOQUEJANTS = ('V', 'B', 'C', 'G', CODI_INACTIU, CODI_REINCORPORACIO)
MOTIU_TEXT = {'V': 'vacances', 'B': 'baixa', 'C': 'congrés', 'G': 'guàrdia externa',
              CODI_INACTIU: 'NO actiu aquest mes', CODI_REINCORPORACIO: 'marge de reincorporació'}

POOL_L, POOL_C = 'L', 'C'          # laborables (dl-dv) · cap de setmana i festius
POOL_NOM = {POOL_L: 'laborables', POOL_C: 'cap de setmana i festius'}


@dataclass
class Radiologist:
    name: str
    role: str = ROL_GUARDIES
    email: str = ''
    active: bool = True
    min_m: float | None = None
    max_m: float | None = None
    prefs: dict = field(default_factory=dict)      # clau -> 'S' / 'N'
    mati: str | None = None                        # 'S' / 'N' / None
    dies_puntuals: str = ''
    observacions: str = ''
    row: int = 0

    @property
    def klass(self) -> str:
        if self.role == ROL_FIX:
            return K_FIX
        if self.role == ROL_NOU:
            return K_NOU
        # Màx. buit o 0 = no vol fer guàrdies: 1 guàrdia cada N mesos (decisió de la Direcció)
        return K_PROP if (self.max_m is not None and self.max_m > 0) else K_SENSE

    @property
    def te_preferencies(self) -> bool:
        return any(v == PREF_SI for v in self.prefs.values())


@dataclass
class Config:
    fix_slots: dict = field(default_factory=dict)       # (dow, '8-16'|'16-20', 'N'|'B'|'N i B') -> nom
    first_wed: str | None = None                        # 1r dimecres (mes natural) 16-20; None = no hi ha fila (és guàrdia)
    friday: dict = field(default_factory=dict)          # ordinal 1..5 -> nom (divendres 8-16, N i B)
    rest_days: int = 1                                  # dies de descans mínims entre guàrdies
    min_every_months: int = 2                           # 1 guàrdia cada N mesos (sense resposta i terra)
    sun_nit_only: list = field(default_factory=list)    # compatibilitat amb Perfils especials
    weekend_day_only: list = field(default_factory=list)


@dataclass
class Period:
    year: int
    month: int
    dates: list
    festius: set = field(default_factory=set)

    @property
    def start(self) -> dt.date:
        return self.dates[0]

    @property
    def end(self) -> dt.date:
        return self.dates[-1]

    @property
    def n_days(self) -> int:
        return len(self.dates)

    @property
    def label(self) -> str:
        return month_label(self.year, self.month)

    @property
    def code(self) -> str:
        return f"{self.year}-{self.month:02d}"

    def contains(self, d: dt.date) -> bool:
        return self.start <= d <= self.end

    def is_festiu(self, d: dt.date) -> bool:
        return d in self.festius

    def pool(self, d: dt.date) -> str:
        return POOL_C if (d.weekday() >= 5 or d in self.festius) else POOL_L

    @staticmethod
    def is_first_wed(d: dt.date) -> bool:
        return d.weekday() == 2 and d.day <= 7

    @staticmethod
    def friday_ordinal(d: dt.date) -> int:
        return weekday_ordinal_in_month(d)

    def week_mondays(self) -> list:
        m = self.start - dt.timedelta(days=self.start.weekday())
        out = []
        while m <= self.end:
            out.append(m)
            m += dt.timedelta(days=7)
        return out

    @property
    def n_weeks(self) -> int:
        return len(self.week_mondays())

    def text(self) -> str:
        return f"{fmt_dia(self.start, True)} → {fmt_dia(self.end, True)}"

    def festius_text(self) -> str:
        return ', '.join(f"{d:%d/%m}" for d in sorted(self.festius) if self.contains(d))


@dataclass(frozen=True)
class Slot:
    date: dt.date
    shift: str      # 'dia' (8-20) | '16-20' | 'nit' (20-08)
    group: str      # 'NB' | 'G1' | 'G2'

    @property
    def dow(self) -> int:
        return self.date.weekday()

    def label(self) -> str:
        sh = {'dia': 'dia 8-20', '16-20': 'tarda 16-20', 'nit': 'nit 20-08'}[self.shift]
        g = {'NB': 'N i B', 'G1': 'Grup 1', 'G2': 'Grup 2'}[self.group]
        return f"{fmt_dia(self.date)} {sh} ({g})"


@dataclass
class FixedAssign:
    date: dt.date
    shift: str          # '8-16' | '16-20'
    role: str           # 'N' | 'B' | 'N i B'
    name: str | None    # titular configurat
    status: str = 'ok'  # 'ok' | 'substituir' | 'pendent'
    reason: str = ''
    note: str = ''      # p. ex. 'fix 1r dimecres', 'divendres 2n'
    suggestions: list = field(default_factory=list)


@dataclass
class Entrada:
    period: Period
    radiologists: dict                     # nom -> Radiologist
    config: Config
    absences: dict                         # nom -> {date: codi}
    warnings: list = field(default_factory=list)

    def active(self, klass=None) -> list:
        out = [r for r in self.radiologists.values() if r.active]
        return [r for r in out if klass is None or r.klass == klass]

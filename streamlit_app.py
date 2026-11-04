"""Wrapper a l'arrel perquè Streamlit Cloud trobi l'app.

Streamlit Cloud pot conservar a sys.modules la versió ANTIGA dels mòduls propis després
d'un desplegament; es purguen abans d'executar l'app perquè es carreguin sempre del disc.
"""
import sys
from pathlib import Path

ROOT = Path(__file__).parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
for _m in [m for m in list(sys.modules) if m == 'guardies' or m.startswith('guardies.')]:
    sys.modules.pop(_m, None)
exec(open(ROOT / 'app' / 'app.py', encoding='utf-8').read())

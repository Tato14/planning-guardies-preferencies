# Planning de guàrdies TD — preferències i capacitat (v4)

App per generar el planning mensual de guàrdies de Telediagnòstic de l'IDI. Ja no hi ha roda alfabètica: cada guàrdia va a qui li toca segons la seva capacitat declarada, les seves preferències i el que ja ha fet els mesos anteriors. Les regles són a [`docs/Regles.md`](docs/Regles.md).

## Workflow mensual

```
Entrada del mes ──┐
                  ├──► App ──► Planning del mes ──► (canvis manuals) ──► Publicació el dia 1
Planning anterior ┘        └──► Entrada del mes següent, ja preparada
```

1. **Entrada del mes.**
   - Fes servir la que l'app va donar en generar el mes anterior.
   - Revisa la columna *Actiu* de Radiòlegs i marca les absències a *Vacances*. Els festius oficials de Catalunya ja hi són.
2. **Planning del mes anterior.** Puja l'últim fitxer, amb els canvis que s'hi hagin fet a mà. D'aquí l'app treu:
   - qui ha fet guàrdia, per alternar els que no han respost;
   - el deute de cadascú;
   - l'última guàrdia, per al descans a l'inici del període.
3. **Generar.**
   - Revisa la pestanya *Pendents*: substitucions de fixos amb propostes, guàrdies sense cobrir i observacions.
   - Fes els canvis manuals que calguin directament a les pestanyes setmanals.
4. **Publicar** el dia 1 (Teams o Outlook, fluxos de Microsoft 365).
5. **Desar** el planning publicat: el mes següent serà el *planning anterior*.

El planning porta una pestanya **Historic** amb l'estat per al mes següent. No s'ha d'editar.

## Fitxers

- **Entrada:** `Configuració` (torns fixos, divendres per ordinal, paràmetres), `Radiòlegs` (rol, correu, actiu, mín./màx., preferències per torn, matí, dies puntuals i observacions) i `Vacances` (una columna per dia del període).
- **Planning:**
  - una pestanya `Setmana NN (dd mmm-dd mmm)` per setmana, amb el mateix format que llegeix l'Office Script de Microsoft 365;
  - `Resum`, `Pendents` i `Historic`.
- **Plantilles:** les genera l'app; al repositori no hi ha cap Excel.
  - Plantilla buida: a la barra lateral.
  - Entrada d'un altre mes: a *Eines*.
  - Actualització de preferències des d'una enquesta: a *Eines*.

## Desplegament

**Streamlit Community Cloud.**
- Repositori: aquest. Fitxer principal: `streamlit_app.py`. Python 3.11.
- Contrasenya: secret `APP_PASSWORD` a *Settings → Secrets*. Sense secret, l'app queda oberta.
- Cada push a `main` redesplega l'app.

**Infraestructura interna:** amb el `Dockerfile` (port 8501). La contrasenya es passa com a variable d'entorn `APP_PASSWORD`.

**Privacitat.**
- Al repositori no hi ha **cap dada personal**: el `.gitignore` bloqueja tots els Excel i la demo és fictícia i es genera per codi.
- Les dades reals (noms, correus, absències, observacions) només passen per la sessió de l'app i no es desen enlloc.
- Streamlit Community Cloud és un servei extern a l'IDI. Abans d'usar-la amb dades reals, valideu-ho amb el DPD o desplegueu-la internament amb Docker.

## Línia d'ordres i proves

```bash
pip install -r requirements.txt pytest
python -m guardies generar Entrada_2027-01.xlsx --anterior Planning_2026-12.xlsx
python -m guardies plantilla Entrada_2027-01.xlsx --mes 2027-02
python -m guardies enquesta Entrada_2027-01.xlsx Taula_preferencies.xlsx
python -m pytest -q
```

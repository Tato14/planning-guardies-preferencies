# Regles d'assignació — Planning de guàrdies TD v4

Aquest document descriu què fa l'app i per què. Recull les vuit premisses de la Direcció Clínica de Telediagnòstic (correu del 26/10/2026) i les decisions preses a partir d'elles.

## 1. Període

- Cada planning va per **setmanes completes**, de dilluns a diumenge.
- Una setmana pertany al mes en què cau el seu **dilluns**. Per exemple, el planning de desembre de 2026 va del dl 07/12/2026 al dg 03/01/2027.
- Cap planning inclou dies del mes anterior.
- **Publicació (premissa 5):** el planning es publica el dia 1 del mes anterior al període. El de gener de 2027 (dl 04/01) es publica l'1 de desembre. Cal adaptar-hi el disparador del Flow 1 de Microsoft 365, que ara envia el planning 15 dies abans.

## 2. Torns

Els torns fixos es defineixen a la pestanya **Configuració** i no es reparteixen:

- **Matins (8-16) de dilluns a dijous:** Neuro i Body, cada un amb el seu titular.
- **Tardes (16-20) de dilluns i dimarts:** Neuro i Body.
- **1r dimecres del mes natural, a la tarda:** un titular que cobreix Neuro i Body.
- **Divendres al matí (premissa 6):**
  - Roda fixa de quatre radiòlegs, un per al 1r, el 2n, el 3r i el 4t divendres del mes natural.
  - El **5è divendres** el cobreix una persona designada a Configuració.
  - Si el titular d'un divendres falta, la primera proposta de substitut és la persona del 5è divendres.
- **Body dels dimecres al matí (premissa 7):** una titular habitual, configurada a Configuració.

Si un titular té una absència, el planning ho marca com a `[SUBSTITUIR — Nom: motiu]`. Si una fila de Configuració no té titular, surt com a `[PENDENT …]`; els divendres al matí sempre es tracten així quan falta el titular de l'ordinal. La pestanya **Pendents** proposa fins a tres substituts.

Si Configuració defineix una tarda fixa de dimecres, dijous o divendres, aquella tarda deixa de ser guàrdia.

- **Criteri de les propostes:** primer qui ha dit Sí al torn (Matí, o la tarda corresponent), després qui ho té indiferent. Puja qui ho esmenta a «Dies puntuals» i, a igualtat, va primer qui porta menys càrrega.
- **Especialitat:** no se'n té en compte; cal comprovar Neuro o Body.

Les **guàrdies** són la resta de torns i es reparteixen:

- **Dimecres:** tarda (llevat del 1r dimecres) i nit.
- **Dilluns i dimarts:** nit.
- **Dijous i divendres:** tarda i nit, en dos grups.
- **Dissabte:** dia i nit, en dos grups.
- **Diumenge:** dia i nit.
- **Festius:** dia i nit (dos grups si cauen en dijous o divendres).

## 3. Qui fa què

| Rol (pestanya Radiòlegs) | Fa |
|---|---|
| `fix` | Només els seus torns fixos |
| `guàrdies` | Guàrdies segons capacitat i preferències |
| `fix-i-guàrdies` | Els seus fixos i guàrdies (mai el mateix dia que un fix) |
| `nou-incorporat` | Les tardes dels dimecres (menys el 1r), fins al seu màxim |
| `només-dg-nit` | Només nits de diumenge |
| `només-cap-setmana-dia` | Només el torn de dia de dissabte o diumenge |

**Actiu = No:** aquest mes la persona no fa ni fixos ni guàrdies, i no cal esborrar-la de la llista.

## 4. Repartiment (premisses 1, 3 i 4)

**Dues bosses independents (premissa 1):**

- **Laborables:** de dilluns a divendres.
- **Caps de setmana i festius.**

Cada bossa es reparteix per separat, de manera que la càrrega de cap de setmana també és proporcional.

**Capacitat declarada (premissa 4).** Qui ha omplert Mín./Màx. rep la mateixa fracció del seu **màxim** que la resta:

- Amb les dades d'octubre de 2026 la fracció és ~36%. Màx. 12 → 4-5 guàrdies al mes; màx. 6 → unes 2.
- Ningú arriba al seu mínim, perquè la suma dels mínims (105) supera les guàrdies que hi ha (~61 per a aquest grup).
- La part de cadascú s'ajusta als torns que pot fer aquell mes. Les absències i els «No» la redueixen sense generar deute.

**Terra.** Ningú que faci guàrdies en té menys d'**una cada dos mesos**. Afecta qui ha declarat màxim 1.

**Qui no ha respost** (Mín. i Màx. buits i no és fix) o ha declarat **màxim 0** fa **exactament una guàrdia cada dos mesos**:

- **Primer mes:** es reparteixen en dos grups alternats per ordre alfabètic.
- **Després:** li toca si fa dos mesos o més de la seva última guàrdia.

**Deute acumulat.** El que algú fa de més o de menys respecte a la seva part es compensa els mesos següents. Un mes amb Actiu = No conserva el deute:

- L'app ho llegeix del planning anterior, fins i tot dels canvis que s'hi hagin fet a mà.
- El deute es limita a ±2 guàrdies per bossa.

## 5. Preferències (premisses 2, 3 i 8)

A la pestanya Radiòlegs, cada tipus de torn (Dl tarda … Festiu nit) es marca així:

- **No:** no pot. No s'assigna mai.
- **Sí:** preferit. Té prioritat.
- **Buit:** indiferent. Es fa servir quan cal.

En un **festiu entre setmana**, el «No» del dia també compta. Qui ha posat No a «Dt nit» tampoc no fa la nit d'un dimarts festiu.

La columna **Matí** (Sí/No) i **Dies puntuals** serveixen per a les substitucions de fixos del matí.

**Observacions (premissa 8):** condicions particulars de cada radiòleg. Surten a Resum i a Pendents perquè la tècnica les tingui presents.

**Absències (pestanya Vacances):**

| Codi | Significat | Efecte |
|---|---|---|
| V | Vacances | Bloqueja el dia i dos dies més de marge de reincorporació |
| B | Baixa | Igual que V |
| C | Congrés | Bloqueja el dia |
| G | Guàrdia en una altra institució | Bloqueja el dia, l'anterior i el posterior |
| X | Guàrdia ja assignada | Informatiu |

**Lectura del planning anterior.**
- Les guàrdies es reconeixen per la franja i la data, no pel text. Els matins entre setmana, les tardes de dilluns i dimarts i les tardes fixes de Configuració no compten mai com a guàrdia.
- Un nom escrit a mà només es compta si correspon a una sola persona de la llista. Si no, surt com a avís.
- Si dues pestanyes repeteixen una data, només es llegeix la primera.

**Entrada del mes següent.** Les absències V, B o G dels dos últims dies del període ja apareixen a les columnes grises.

## 6. Descans

- **Mai** dues guàrdies el mateix dia, ni una guàrdia el dia d'un torn fix propi.
- Festius: s'usen els calendaris oficials de Catalunya (2026 i 2027 verificats). Per als anys sense calendari publicat es marquen els 14 habituals i la plantilla avisa que són provisionals.
- **S'eviten**, tret que sigui imprescindible per cobrir:
  - guàrdies en dies consecutius, també entre l'últim dia del període anterior i el primer d'aquest;
  - una nit just abans d'un matí fix propi.
- El nombre de dies de descans és un paràmetre de Configuració (per defecte 1).

## 7. Prioritats del càlcul

L'assignació es resol amb un optimitzador (OR-Tools CP-SAT, determinista: les mateixes dades donen el mateix planning). En ordre d'importància:

1. Cobrir totes les guàrdies.
2. Regles dures: absències, «No», rols, un torn al dia, màxim declarat i res el dia d'un fix.
3. Una guàrdia cada dos mesos per a qui no ha respost.
4. Equitat per bosses respecte a la part proporcional, més el deute.
5. Descans.
6. Preferències «Sí».

Una guàrdia que no es pot cobrir complint les regles queda com a `[SENSE COBRIR]` i surt a Pendents.

## 8. Revisió i proves

Una revisió independent del codi va trobar 22 punts (3 d'alta prioritat). Tots els d'alta i mitjana prioritat estan corregits i tenen prova automàtica (`tests/test_revisio.py`).

## 9. Resultats de prova

Simulació de 13 mesos encadenats (des. 2026 - des. 2027) amb la llista real, vacances, baixes, congressos i guàrdies externes aleatòries:

- 0 guàrdies sense cobrir.
- 0 infraccions de regles dures.
- 0 guàrdies en dies consecutius.
- Els que no han respost han fet exactament una guàrdia cada dos mesos.
- Guàrdies en 13 mesos segons el màxim declarat: màx. 1 → 6, màx. 4 → 18, màx. 6 → 29, màx. 12 → 56.
- Temps de càlcul: ~5 s per mes.

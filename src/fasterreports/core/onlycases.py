"""Gli scaffali della sezione Only Cases, e quanto ne serve ogni settimana.

I due fogli sono arrivati nel template il 2026-08-21, disegnati e formattati a
mano:

    OC Eventi              helper: un evento `Available Cases` per riga, puntato
                           alla riga di AT_DATASET da cui viene
    Only Cases Dashboard   la vista: ore per agente, quante dentro lo slot di
                           back office e quante fuori (la malpractice)

Entrambi vivono di `AT_DATASET`, e nessuno dei due viene scritto dalla pipeline:
si ricalcolano da soli. Quello che il programma deve sapere e' una cosa sola —
**se ci sta**.

LA FORMA DEL DIFETTO, che qui e' scritta due volte. `OC Eventi`!A5 e' un array
dinamico che spilla i numeri di riga degli eventi:

    LET(stato, 'Only Cases Dashboard'!$B$5,
        r, SEQUENCE(129999,1,2),
        s, AT_DATASET!$F$2:$F$130000,
        FILTER(r, s=stato, ""))

Cresce quanto serve. Le colonne `B:T` accanto, invece, sono formule riga per riga
tirate fino a 3004 — cioe' 3000 eventi. L'evento 3001 comparirebbe in colonna A e
basta: nessun errore, solo diciannove celle vuote dove ci sono i suoi minuti. Lo
stesso vale per l'elenco degli agenti della dashboard, che ha posto per 60.

QUESTI DUE SCAFFALI HANNO UNA GUARDIA DENTRO IL FOGLIO — le celle `C8` e `C9`
della dashboard dicono «ATTENZIONE: capienza … esaurita». E' un buon lavoro, e
non basta: quelle celle si leggono solo aprendo il file e guardando li'. Il
preflight le anticipa, e lo fa **all'80%**, cosi' le formule si tirano quando c'e'
tempo e non nella settimana in cui i numeri sono gia' incompleti. Che i due
numeri coincidano lo verifica un test sul template vero
(`test_la_capienza_misurata_e_quella_che_dichiara_il_foglio`): se un domani le
formule venissero tirate senza aggiornare la guardia — o viceversa — si vedrebbe
li'.

MISURATO IL 2026-08-24 sul template e sull'export AT del W30 (26 541 righe):

    eventi Only Cases      1 386 su 3 000   46%
    agenti Only Cases         33 su    60   55%

Il blocco giornaliero in fondo alla dashboard (righe 79-93, il dettaglio del
singolo agente) tiene 15 giorni per una settimana da 7: non e' sorvegliato,
perche' non c'e' nessuna crescita che possa portarci vicino.
"""

from __future__ import annotations

from .capienze import Scaffale
from .capienze import punti_da_misurare as _punti

FOGLIO_HELPER = "OC Eventi"
FOGLIO_DASHBOARD = "Only Cases Dashboard"

# La cella della dashboard che dice QUALE stato agente e' "only cases". La si
# legge dal template invece di scriverla qui: e' un parametro del foglio, e chi
# lo cambia deve ottenere che il controllo cambi con lui. Vedi
# `orchestrate._oc_stato`.
CELLA_STATO = ("B", 5)

# Il valore su cui ripiegare se il template non si lascia leggere. E' quello che
# c'e' nel foglio dal primo giorno, ed e' anche uno dei quattro valori dichiarati
# per `AT_DATASET!F` nel contratto.
STATO_DI_DEFAULT = "Available Cases"

SCAFFALI: tuple[Scaffale, ...] = (
    Scaffale(
        etichetta="eventi Only Cases",
        # Colonna B e non A: A e' l'array dinamico, che ha UNA formula (in A5) e
        # spilla. Misurarla darebbe capienza 1. B e' la prima delle diciannove
        # colonne tirate riga per riga, ed e' quella che finisce per prima.
        sheet=FOGLIO_HELPER, col="B", prima_riga=5,
        campo=None,  # non e' un conteggio di distinti: vedi `conteggi`
        nota="colonne B:T di 'OC Eventi' da tirare insieme",
    ),
    Scaffale(
        etichetta="agenti Only Cases",
        # Colonna G e non una delle altre: l'elenco degli agenti sta in `A12:J71`,
        # ma sotto — a riga 79 — comincia il dettaglio giornaliero del singolo
        # agente, e B, C, D, E, F hanno formule anche li'. Misurare una di quelle
        # direbbe 83 posti invece di 60, cioe' il controllo tacerebbe fino a
        # ventitre' agenti oltre il vero. G, H, I, J si fermano alla riga 72 del
        # totale di squadra, che si scarta con `coda=1`.
        sheet=FOGLIO_DASHBOARD, col="G", prima_riga=12, coda=1,
        campo="Agent Email",
        nota="colonne A:J di 'Only Cases Dashboard' da tirare insieme, sopra la riga 72 del totale",
    ),
)


def punti_da_misurare() -> tuple[tuple[str, str, int], ...]:
    """Gli argomenti per `templatescan.scan_formula_extent`."""
    return _punti(SCAFFALI)


def conteggi(righe: list[list], offset: dict[str, int], stato: str) -> dict[str, int]:
    """Quanto chiede la settimana a ciascuno scaffale.

    `righe` sono le righe del blocco `AT_DATASET`; `offset` mappa il nome canonico
    del campo alla sua posizione nella riga; `stato` e' il valore di
    `AT_DATASET!F` che i due fogli considerano "only cases".

    Si contano solo le righe in quello stato, perche' e' quello che fa `FILTER`
    nel foglio: contare tutte le righe di AT_DATASET direbbe 26 541 dove il foglio
    ne mostra 1 386, e il controllo bloccherebbe ogni settimana per un problema
    che non esiste.
    """
    i_stato = offset.get("Agent State")
    i_email = offset.get("Agent Email")
    if i_stato is None:
        return {}

    def _cella(r: list, i: int) -> str:
        return str(r[i]).strip() if i < len(r) and r[i] is not None else ""

    eventi = [r for r in righe if _cella(r, i_stato) == stato]
    out: dict[str, int] = {"eventi Only Cases": len(eventi)}
    if i_email is not None:
        out["agenti Only Cases"] = len(
            {e for r in eventi if (e := _cella(r, i_email))}
        )
    return out

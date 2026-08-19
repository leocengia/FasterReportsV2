"""Gli "scaffali" dei fogli Duplicate Cases, e quanto ne serve ogni settimana.

I tre fogli DC presentano elenchi in una forma che non ha nessun altro foglio del
workbook: **l'elenco dei nomi e' un array dinamico** (`SORTBY(UNIQUE(FILTER(...)))`)
che cresce da se', mentre **le colonne accanto** — il conteggio, la media, la
percentuale — hanno una formula scritta riga per riga e si fermano dove le ha
tirate chi ha fatto il foglio.

La conseguenza e' quella di sempre, e arriva da sola: il giorno in cui gli agenti
coinvolti nei duplicati passano da 34 a 35, il trentacinquesimo compare
nell'elenco **senza nessun numero accanto**. Presente e invisibile insieme.
Nessun `#SPILL!`, nessun `#REF!`: la cella accanto e' semplicemente vuota, come
lo era prima.

E' lo stesso difetto che 'Helper CaseType' aveva ad agosto, e si controlla nello
stesso modo: si misura la capienza **dal template** (non si scrive qui: se domani
tiri le formule piu' in basso, il controllo ti segue) e si confronta con quanti
valori distinti ci sono davvero nei dati.

MISURATO IL 2026-08-19 sull'export della W32 (178 duplicati, 28 agenti):

    agenti                      28 su 34   82%   <-- il piu' stretto
    Type                        19 su 26   73%
    record type                  7 su 12   58%
    parent con >1 duplicato      9 su 20   45%
    Case Origin distinti         8 su 15   53%

Il team HPO e' di circa 36 persone: basta una settimana in cui ne finiscono 35
nei duplicati.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Scaffale:
    """Un elenco a capienza fissa in un foglio DC.

    `sheet`/`col`/`prima_riga` dicono **dove misurare** la capienza: l'ultima riga
    che in quella colonna ha una formula. `campo` dice **cosa contare** nei dati
    per sapere quanto ne serve.
    """

    etichetta: str
    sheet: str
    col: str
    prima_riga: int
    # Il nome canonico del campo di DUP_DATASET di cui contare i valori distinti.
    # `None` per gli scaffali che si contano in un altro modo (vedi `conteggi`).
    campo: str | None
    nota: str = ""

    @property
    def punto(self) -> tuple[str, str, int]:
        return (self.sheet, self.col, self.prima_riga)


# Le colonne elencate sono quelle da TIRARE PIU' IN BASSO quando la capienza
# finisce; la prima di ciascun gruppo e' quella che si misura (le altre dello
# stesso blocco arrivano sempre alla stessa riga, ed e' un test a garantirlo).
SCAFFALI: tuple[Scaffale, ...] = (
    Scaffale(
        etichetta="agenti con casi duplicati",
        sheet="DC Agents & Categories", col="B", prima_riga=7,
        campo="Full Name",
        nota="colonne B, C, D da tirare insieme",
    ),
    Scaffale(
        etichetta="record type",
        sheet="DC Agents & Categories", col="G", prima_riga=7,
        campo="Case Record Type",
        nota="colonne G, H",
    ),
    Scaffale(
        etichetta="Type (case type)",
        sheet="DC Agents & Categories", col="K", prima_riga=7,
        campo="Type",
        nota="colonne K, L",
    ),
    Scaffale(
        etichetta="parent con piu' di un duplicato",
        sheet="DC Timing & Quality", col="G", prima_riga=17,
        campo=None,  # non e' un conteggio di distinti: vedi `conteggi`
        nota="colonne G, H, I",
    ),
    Scaffale(
        etichetta="agenti (aree dei grafici)",
        sheet="DC Dashboard", col="AB", prima_riga=21,
        campo="Full Name",
        nota="colonna AB, nascosta",
    ),
    Scaffale(
        etichetta="record type (aree dei grafici)",
        sheet="DC Dashboard", col="BB", prima_riga=21,
        campo="Case Record Type",
        nota="colonna BB, nascosta",
    ),
    Scaffale(
        etichetta="Type (aree dei grafici)",
        sheet="DC Dashboard", col="BE", prima_riga=21,
        campo="Type",
        nota="colonna BE, nascosta",
    ),
)

# Lo spill della lista `Case Origin` in 'DC Dashboard'!AA5, che alimenta il menu a
# tendina del filtro Origin. Qui il limite NON e' una formula per riga: e' un
# array dinamico che parte da riga 5 e sbatte contro `AA20` ('Agent'), cioe' ha 15
# posti. Se i `Case Origin` distinti li superano, quella cella diventa `#SPILL!`
# e il filtro del dashboard smette di funzionare — un guasto visibile, a
# differenza degli altri, ma comunque da vedere arrivare.
#
# Il rimedio non e' tirare una formula: e' spostare il blocco `AA20:AB...` piu' in
# basso, aggiornando anche le serie del primo grafico. Per questo sta a parte.
SPILL_ORIGIN = Scaffale(
    etichetta="Case Origin distinti (menu del filtro)",
    sheet="DC Dashboard", col="AA", prima_riga=5,
    campo="Case Origin",
    nota="spill di AA5, bloccato da AA20: spostare il blocco AA20:AB… piu' in basso",
)
SPILL_ORIGIN_CAPIENZA = 15


def punti_da_misurare() -> tuple[tuple[str, str, int], ...]:
    """Gli argomenti per `templatescan.scan_formula_extent`."""
    return tuple(s.punto for s in SCAFFALI)


def conteggi(righe: list[list], offset: dict[str, int]) -> dict[str, int]:
    """Quanti valori distinti servono, per etichetta di scaffale.

    `righe` sono le righe del blocco `DUP_DATASET`; `offset` mappa il nome
    canonico del campo alla sua posizione nella riga.

    'parent con piu' di un duplicato' non e' un conteggio di distinti ma di
    **gruppi con almeno due elementi**: e' il numero di righe che la tabella
    "PARENT CASES WITH MORE THAN ONE DUPLICATE" arriva a mostrare.
    """
    out: dict[str, int] = {}
    for s in (*SCAFFALI, SPILL_ORIGIN):
        if s.campo is None:
            continue
        i = offset.get(s.campo)
        if i is None:
            continue
        out[s.etichetta] = len({
            str(r[i]).strip() for r in righe
            if i < len(r) and r[i] is not None and str(r[i]).strip()
        })

    i = offset.get("Parent Case: Case Number")
    if i is not None:
        conta: dict[str, int] = {}
        for r in righe:
            if i < len(r) and r[i] is not None and str(r[i]).strip():
                k = str(r[i]).strip()
                conta[k] = conta.get(k, 0) + 1
        out["parent con piu' di un duplicato"] = sum(
            1 for n in conta.values() if n > 1
        )
    return out

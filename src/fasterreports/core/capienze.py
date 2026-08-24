"""Gli "scaffali": elenchi a capienza fissa nei fogli di presentazione.

Piu' fogli del workbook presentano un elenco nella stessa forma, e con lo stesso
difetto: **l'elenco dei nomi e' un array dinamico** (`SORTBY(UNIQUE(FILTER(...)))`,
`FILTER(SEQUENCE(...))`) che cresce da se', mentre **le colonne accanto** — il
conteggio, la media, la percentuale — hanno una formula scritta riga per riga e
si fermano dove le ha tirate chi ha fatto il foglio.

La conseguenza arriva da sola, senza che nessuno tocchi niente: il giorno in cui
le voci passano da 34 a 35, la trentacinquesima compare nell'elenco **senza
nessun numero accanto**. Presente e invisibile insieme. Nessun `#SPILL!`, nessun
`#REF!`: la cella accanto e' semplicemente vuota, come lo era prima.

Questo modulo tiene la forma comune. Quali scaffali esistano, e quanto ne serva
ogni settimana, lo dicono i moduli delle singole sezioni — `duplicati.py` per i
tre fogli DC, `onlycases.py` per la dashboard Only Cases.

LA CAPIENZA SI MISURA DAL TEMPLATE, non si scrive qui: `templatescan.scan_formula_extent`
trova l'ultima riga che in quella colonna ha una formula. Se domani le formule
vengono tirate piu' in basso, il controllo lo segue da solo. Scriverla come
costante vorrebbe dire avere due verita' che possono divergere, e la seconda
sarebbe muta.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Scaffale:
    """Un elenco a capienza fissa in un foglio di presentazione.

    `sheet`/`col`/`prima_riga` dicono **dove misurare** la capienza: l'ultima riga
    che in quella colonna ha una formula. `campo` dice **cosa contare** nei dati
    per sapere quanto ne serve.

    LA COLONNA DA MISURARE NON E' UNA QUALSIASI del blocco: dev'essere una che
    non prosegue oltre il blocco. `scan_formula_extent` restituisce l'ultima riga
    con una formula, quindi una colonna che continua in un secondo blocco piu' in
    basso — nella dashboard Only Cases il dettaglio giornaliero, quindici righe
    sotto l'elenco degli agenti — darebbe una capienza molto piu' grande di quella
    vera. Cioe' il controllo tacerebbe proprio nel caso in cui deve parlare.

    Una riga di TOTALE subito sotto l'ultima voce e' invece innocua, e non
    costringe a cercare una colonna che non ce l'ha: si dichiara con `coda`.
    """

    etichetta: str
    sheet: str
    col: str
    prima_riga: int
    # Il nome canonico del campo di cui contare i valori distinti. `None` per gli
    # scaffali che si contano in un altro modo (vedi le funzioni `conteggi` dei
    # moduli di sezione).
    campo: str | None
    nota: str = ""
    # Quante righe con formula, nella colonna misurata, stanno SOTTO l'ultima voce
    # dell'elenco e non sono voci: tipicamente la riga del totale.
    coda: int = 0

    @property
    def punto(self) -> tuple[str, str, int]:
        """Gli argomenti per `templatescan.scan_formula_extent`."""
        return (self.sheet, self.col, self.prima_riga)

    def capienza(self, ultima_riga: int) -> int:
        """Quante voci ci stanno, data l'ultima riga con una formula."""
        return ultima_riga - self.prima_riga + 1 - self.coda


def punti_da_misurare(scaffali) -> tuple[tuple[str, str, int], ...]:
    """Gli argomenti per `templatescan.scan_formula_extent`, per piu' scaffali."""
    return tuple(s.punto for s in scaffali)


def capienze(scaffali, estensioni: dict[tuple[str, str], int]) -> dict[str, int]:
    """`{etichetta: capienza}` dalle estensioni misurate sul template.

    Uno scaffale il cui foglio non c'e' (un template piu' vecchio, o una sezione
    non ancora aggiunta) semplicemente non compare: il controllo a valle salta
    quello e fa gli altri, invece di inventare un numero.
    """
    return {
        s.etichetta: s.capienza(estensioni[(s.sheet, s.col)])
        for s in scaffali
        if (s.sheet, s.col) in estensioni
    }

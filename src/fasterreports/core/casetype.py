"""L'elenco (canale, case type) che alimenta 'Helper CaseType'.

Il problema che risolve, misurato sul W33: 'Helper CaseType' nel template porta
una lista scritta a mano di 31 case type per canale (righe 2..32 Phone, 33..63
Non-live). Nei dati di quella settimana i case type distinti erano 32, e
quattro non erano in lista — `Call Assignment` (30 casi), `Specialty Functions`
(8), `Collections` (1), `Live Site Property Settings Issue` (1). Il foglio non
mostrava nessun errore: quei casi semplicemente non esistevano, e `Call
Assignment` superava perfino la soglia di volume, cioe' era un risultato vero
che nessuno avrebbe visto.

Perche' si APPENDE invece di riscrivere tutta la lista ordinata: `CaseType
Deepdive` punta alle righe di 'Helper CaseType' per posizione
(`'Helper CaseType'!$B2`, `$B3`, ...) e porta della formattazione fatta a mano.
Riordinare la lista sposterebbe i case type sotto le righe sbagliate e
rovinerebbe quel lavoro. Lasciando ferme le righe che ci sono e mettendo le
nuove in fondo, il deepdive resta esattamente com'e' — e chi vuole i numeri
completi li trova nell'helper.

Nessun Excel qui: sono liste di stringhe. La parte che parla con xlwings sta in
`omni/writer.py`.
"""

from __future__ import annotations

from .aht_history import mappa_canale


def normalizza(canale: str, case_type: str) -> tuple[str, str]:
    return (str(canale).strip(), str(case_type).strip())


def coppie_dai_dati(righe) -> list[tuple[str, str]]:
    """Le combinazioni (canale, case type) presenti nell'export.

    `righe` e' un iterabile di coppie (canale grezzo, case type) come stanno in
    `SF_DATABASE`. Il canale viene tradotto nelle etichette dell'helper
    (`Phone` / `Non-live`), le stesse che usa lo storico.
    """
    viste: dict[tuple[str, str], None] = {}
    for canale_raw, case_type in righe:
        ct = "" if case_type is None else str(case_type).strip()
        if not ct:
            continue
        viste.setdefault((mappa_canale(canale_raw), ct), None)
    return sorted(viste)


def da_appendere(esistenti, presenti) -> list[tuple[str, str]]:
    """Le coppie da aggiungere in fondo: quelle nei dati e non ancora in lista.

    Ordinate per canale e case type, cosi' che due giri sugli stessi dati
    producano lo stesso risultato — l'ordine di scoperta nell'export non deve
    entrare nel foglio.

    Non restituisce mai qualcosa da RIMUOVERE: un case type che questa settimana
    non ha casi resta in lista con volume zero, ed e' giusto — togliendolo si
    sposterebbero tutte le righe sotto di lui, che e' esattamente cio' che
    'CaseType Deepdive' non puo' sopportare.
    """
    gia = {normalizza(c, t) for c, t in esistenti}
    return sorted({normalizza(c, t) for c, t in presenti} - gia)

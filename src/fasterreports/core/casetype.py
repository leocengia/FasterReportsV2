"""Le coppie (canale, case type) dei dati, contro la lista curata del template.

Il problema che risolve, misurato sul W33: 'Helper CaseType' nel template porta
una lista scritta a mano di 31 case type per canale (righe 2..32 Phone, 33..63
Non-live). Nei dati di quella settimana i case type distinti erano 32, e
quattro non erano in lista — `Call Assignment` (30 casi), `Specialty Functions`
(8), `Collections` (1), `Live Site Property Settings Issue` (1). Il foglio non
mostrava nessun errore: quei casi semplicemente non esistevano, e `Call
Assignment` superava perfino la soglia di volume, cioe' era un risultato vero
che nessuno avrebbe visto.

La lista curata vive in 'Helper CaseType' del template, ed e' l'utente che la
decide: 'CaseType Deepdive' punta a quelle righe per posizione e porta della
formattazione fatta a mano, e le heat map di 'AHT Trend WoW' mostrano solo i case
type che quella lista ammette.

Fino al 2026-08-20 le coppie nuove venivano APPESE in fondo all'helper. E'
stato tolto, perche' allargava la lista curata da se': la conseguenza, misurata
nella W33, era che due combinazioni mai viste prima aggiungevano due righe alle
heat map — con un dato su dodici colonne, e spostando l'ordinamento di tutte le
altre. Ora le coppie fuori lista si ELENCANO nel preflight, con volume e AHT, e
chi cura la lista decide.

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


def fuori_lista(curati, presenti) -> list[tuple[str, str]]:
    """Le coppie nei dati che la lista curata non contiene.

    Ordinate per canale e case type, cosi' che due giri sugli stessi dati
    producano lo stesso elenco — l'ordine di scoperta nell'export non deve
    entrare nel rapporto.

    Non dice mai che qualcosa va RIMOSSO dalla lista curata: un case type che
    questa settimana non ha casi resta in lista, ed e' giusto — togliendolo si
    sposterebbero tutte le righe sotto di lui, che e' esattamente cio' che
    'CaseType Deepdive' non puo' sopportare.
    """
    gia = {normalizza(c, t) for c, t in curati}
    return sorted({normalizza(c, t) for c, t in presenti} - gia)

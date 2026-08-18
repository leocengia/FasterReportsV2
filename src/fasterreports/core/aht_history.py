"""Storico settimanale di volume e AHT per case type.

Perche' esiste: il trend settimana-su-settimana ha bisogno di ricordare le
settimane passate, e l'export di Salesforce contiene solo quella corrente. Il
workbook generato viene sovrascritto ogni settimana, quindi non puo' essere lui
la memoria. La memoria e' un CSV fuori dal workbook (`data/aht_history.csv`), e
dentro il report ne finisce solo una copia che le formule di 'AHT Trend WoW'
leggono.

Qui non c'e' Excel, non c'e' xlwings, non c'e' niente che serva una macchina
Windows: sono liste, tuple e un file di testo. E' voluto — e' la parte in cui
si annidano gli errori di conteggio, e va poter essere verificata per intero
con dei test.

Due cose che il foglio impone e che vale la pena avere in chiaro:

1. **Il canale non e' quello di Salesforce.** L'export dice `Phone` / `Other`;
   lo storico dice `Phone` / `Non-live`. La seconda etichetta e' quella che usa
   il business, ed e' anche quella che 'Helper CaseType' si aspetta in colonna A
   (dove la rimappa a `Other` per interrogare `AHT_Data`). Tradurre qui, una
   volta, evita di avere due vocabolari in giro per il progetto.

2. **`week` e' la settimana ISO, e riparte da 1 ogni gennaio.** E' la
   convenzione dell'azienda, confermata. Il guaio e' che ordinare o filtrare
   sulla sola settimana, a cavallo d'anno, mette la W1 del 2027 *sotto* la W52
   del 2026 (il trend mostrerebbe sempre le vecchie) e fa collidere la W05 del
   2027 con la W05 del 2026 (`SUMIFS` le somma insieme, in silenzio). Percio'
   lo storico porta anche `iso_year` e `week_key` = `iso_year * 100 + week`:
   `week` resta l'etichetta da leggere, `week_key` e' quella su cui si ordina e
   si confronta. E' una ridondanza deliberata — costa una colonna e rende le
   formule del template banali invece di far loro ricostruire la chiave.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from .errors import PipelineError

CANALE_LIVE = "Phone"
CANALE_NON_LIVE = "Non-live"

# L'ordine e' anche quello delle colonne A..G del foglio 'AHT History': tenerli
# allineati fa si' che il CSV si legga senza dover consultare altro.
COLONNE = ("week", "channel", "case_type", "volume", "aht", "iso_year", "week_key")


@dataclass(frozen=True)
class RigaStorico:
    iso_year: int
    week: int
    channel: str
    case_type: str
    volume: int
    aht: float

    @property
    def week_key(self) -> int:
        return self.iso_year * 100 + self.week

    @property
    def chiave(self) -> tuple[int, str, str]:
        """Cosa rende una riga "la stessa riga" di un'altra.

        E' la chiave dell'append idempotente: rigenerare la stessa settimana
        due volte deve sostituire, non raddoppiare.
        """
        return (self.week_key, self.channel, self.case_type)

    def as_row(self) -> list:
        return [
            self.week, self.channel, self.case_type,
            self.volume, self.aht, self.iso_year, self.week_key,
        ]


def mappa_canale(raw) -> str:
    """`Case Origin (group)` -> l'etichetta dello storico.

    Salesforce distingue solo `Phone` da tutto il resto (nel W33: 1976 `Phone`,
    857 `Other`, nient'altro). Tutto cio' che non e' telefono e' "non live":
    email, chat, casi aperti dal partner.
    """
    if raw is None:
        return CANALE_NON_LIVE
    return CANALE_LIVE if str(raw).strip().lower() == "phone" else CANALE_NON_LIVE


def aggrega(righe, *, iso_year: int, week: int) -> list[RigaStorico]:
    """Da (canale, case type, AHT) riga per riga agli aggregati settimanali.

    `righe` e' un iterabile di terne cosi' come stanno nell'export: il canale
    grezzo (`Phone`/`Other`), il case type, e l'AHT del singolo caso, che puo'
    essere vuoto.

    Due regole, e la differenza fra le due conta:

    - **volume** = quanti casi. Conta tutte le righe, anche quelle senza AHT.
    - **aht** = media dei soli AHT numerici. Nel W33 sono 123 righe su 2833 a
      non averlo: contarle come zero abbasserebbe la media di un case type
      piccolo senza motivo. Se nessuna riga della combinazione ha un AHT, il
      valore e' 0.0 — la stessa convenzione che il foglio ha gia' (le formule
      `AVERAGEIFS` del template restituiscono 0 nello stesso caso).

    Le combinazioni a volume zero non esistono per costruzione: si aggregano le
    righe che ci sono, non il prodotto cartesiano di canali per case type.
    """
    volumi: dict[tuple[str, str], int] = {}
    somme: dict[tuple[str, str], float] = {}
    conteggi: dict[tuple[str, str], int] = {}

    for canale_raw, case_type, aht_raw in righe:
        ct = "" if case_type is None else str(case_type).strip()
        if not ct:
            # Un caso senza tipo non appartiene a nessun case type: contarlo
            # sotto l'etichetta vuota creerebbe una riga fantasma nello storico
            # e nel foglio.
            continue
        k = (mappa_canale(canale_raw), ct)
        volumi[k] = volumi.get(k, 0) + 1
        val = _numero(aht_raw)
        if val is not None:
            somme[k] = somme.get(k, 0.0) + val
            conteggi[k] = conteggi.get(k, 0) + 1

    out = [
        RigaStorico(
            iso_year=iso_year,
            week=week,
            channel=canale,
            case_type=ct,
            volume=vol,
            aht=(somme[(canale, ct)] / conteggi[(canale, ct)]
                 if conteggi.get((canale, ct)) else 0.0),
        )
        for (canale, ct), vol in volumi.items()
    ]
    return ordina(out)


def _numero(v) -> float | None:
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def ordina(righe) -> list[RigaStorico]:
    """Settimana, poi canale, poi case type — l'ordine che ha il foglio oggi."""
    return sorted(righe, key=lambda r: (r.week_key, r.channel, r.case_type))


def unisci(storico, nuove) -> list[RigaStorico]:
    """Innesta le righe nuove nello storico, sostituendo quelle omonime.

    Rigenerare il report della stessa settimana due volte — cosa che capita
    ogni volta che si corregge un export e si rilancia — non deve raddoppiare
    le righe di quella settimana ne' lasciare mescolati i numeri vecchi con i
    nuovi. La chiave e' `(week_key, channel, case_type)` e vince l'ultimo
    arrivato: il giro piu' recente e' quello fatto sui dati corretti.

    Attenzione a cosa NON fa: non rimuove le settimane vecchie e non rimuove i
    case type che non compaiono piu'. Lo storico e' cumulativo per definizione,
    e la finestra delle ultime 11 settimane la decide il foglio 'AHT Trend WoW'
    con `TAKE`, non questo modulo.
    """
    per_chiave = {r.chiave: r for r in storico}
    for r in nuove:
        per_chiave[r.chiave] = r
    return ordina(per_chiave.values())


def carica(path: Path) -> list[RigaStorico]:
    """Legge lo storico. Un file che non c'e' e' uno storico vuoto, non un errore.

    La prima settimana in assoluto parte da zero, ed e' normale: non e' un
    guasto da segnalare.
    """
    if not Path(path).is_file():
        return []
    try:
        with open(path, newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            mancanti = [c for c in COLONNE if c not in (reader.fieldnames or [])]
            if mancanti:
                raise PipelineError(
                    f"{path}: lo storico AHT non ha le colonne "
                    f"{', '.join(mancanti)}.\n"
                    f"  Intestazioni trovate: {', '.join(reader.fieldnames or ['(nessuna)'])}\n"
                    f"  Attese: {', '.join(COLONNE)}"
                )
            out = []
            for n, raw in enumerate(reader, start=2):
                try:
                    out.append(RigaStorico(
                        iso_year=int(raw["iso_year"]),
                        week=int(raw["week"]),
                        channel=str(raw["channel"]).strip(),
                        case_type=str(raw["case_type"]).strip(),
                        volume=int(float(raw["volume"])),
                        aht=float(raw["aht"] or 0.0),
                    ))
                except (TypeError, ValueError) as exc:
                    raise PipelineError(
                        f"{path}, riga {n}: valore non leggibile ({exc}).\n"
                        f"  Riga: {raw}\n"
                        f"  Lo storico e' la memoria del trend: meglio fermarsi che\n"
                        f"  proseguire scartando in silenzio una settimana."
                    ) from None
    except OSError as exc:
        raise PipelineError(f"{path}: impossibile leggere lo storico AHT ({exc}).") from None
    return ordina(out)


def scrivi(path: Path, righe) -> Path:
    """Riscrive lo storico, in modo atomico.

    Su un file temporaneo poi `os.replace`, per lo stesso motivo per cui lo fa
    il build: se il processo muore a metà scrittura, quello che resta sul disco
    e' lo storico di prima — intero — non mezzo storico con il nome giusto.
    Perdere la memoria del trend costerebbe una ricostruzione a mano.
    """
    import os

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(f"{path.name}.building")
    try:
        with open(temp, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(COLONNE)
            for r in ordina(righe):
                w.writerow(r.as_row())
        os.replace(temp, path)
    except OSError as exc:
        temp.unlink(missing_ok=True)
        raise PipelineError(f"{path}: impossibile scrivere lo storico AHT ({exc}).") from None
    return path


def settimane(righe) -> list[int]:
    """Le `week_key` presenti, dalla piu' vecchia alla piu' recente."""
    return sorted({r.week_key for r in righe})


def righe_foglio(storico) -> list[list]:
    """Lo storico nella forma che va nelle celle di 'AHT History', A..G.

    Va scritto INTERO, non tagliato alle ultime 11 settimane come chiedeva la
    prima stesura della richiesta: 'AHT Trend WoW'!B4 sceglie da se' le 11 piu'
    recenti (`TAKE(SORT(UNIQUE(...)),11)`), quindi tagliare qui butterebbe via
    dati senza che il foglio mostri niente di piu'. Le sue formule leggono fino
    a riga 100000, cioe' circa 1800 settimane: non e' un limite che si incontra.
    """
    return [list(COLONNE)] + [r.as_row() for r in ordina(storico)]

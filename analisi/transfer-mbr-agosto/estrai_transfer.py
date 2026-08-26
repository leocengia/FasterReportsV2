#!/usr/bin/env python3
"""Estrae i transfer da uno o piu' Omni Report, per il template MBR di agosto.

Sola LETTURA: apre il workbook come zip e legge l'XML del foglio in streaming,
senza openpyxl e senza Excel — stessa scelta di `tools/audit_workbook.py`, che
regge anche i workbook da decine di MB. Nessun file del programma viene toccato.

    python analisi/transfer-mbr-agosto/estrai_transfer.py "samples/omni-report/Omni Report W30.xlsm"
    python analisi/transfer-mbr-agosto/estrai_transfer.py input/*.xlsm --out analisi/transfer-mbr-agosto

Cosa produce:
  - il breakdown della sezione 1 del template (2x2 destinazione x canale) su stdout
    e in `breakdown_<etichetta>.md`;
  - `transfer_casi_<etichetta>.csv`, l'export case-level ordinato per rischio, con
    le due colonne `valid_invalid` e `note_scrub` vuote da compilare a mano.

Le colonne si agganciano per NOME, mai per lettera: l'export SF e' gia' cresciuto
in passato (PSAT da 121 a 134 colonne) e le lettere si spostano.
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import random
import re
import sys
import zipfile
from collections import Counter, defaultdict
from pathlib import Path
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

# I fogli in cui puo' trovarsi l'export SF, in ordine di preferenza: l'Omni
# Report lo chiama SF_DATABASE, il WoW CaseType Deepdive `CSV DATASET`, e i
# workbook di analisi una tantum semplicemente `DATASET`.
FOGLI_CANDIDATI = ("SF_DATABASE", "CSV DATASET", "DATASET")

# Campo -> nomi accettati, in ordine. Il primo che si trova vince.
CAMPI: dict[str, tuple[str, ...]] = {
    "case_number": ("case_number", "Case Number"),
    "case_origin": ("Case Origin",),
    "case_origin_group": ("Case Origin (group)",),
    "channel_group": ("Channel Group",),
    "work_function": ("work_function",),
    "omni_level": ("Omni Level Indicator",),
    "case_status": ("case_status",),
    "res_category": ("Case Resolution Category",),
    "resolution_category": ("resolution_category",),
    "case_type": ("Case Type",),
    "record_type": ("Case Record Type",),
    "primary_category": ("Primary Category",),
    "secondary_category": ("Secondary Category",),
    "employee": ("Employee Name",),
    "manager": ("Manager Name",),
    "date_range": ("Date (Range)",),
    "date_viewpoint": ("Date Viewpoint",),
    "aht": ("Case AHT (mins)",),
    "misrouted": ("Misrouted Cases", "Cases Misrouted"),
    "misrouted_count": ("case_misrouted_count",),
    "has_child": ("Has Child Cases",),
    "parentid": ("parentid",),
    "queue": ("Case Queue Name",),
    "vendor": ("Vendor Name",),
}

# Senza queste non si puo' fare niente: meglio fermarsi con un errore chiaro che
# produrre zeri silenziosi.
OBBLIGATORIE = ("case_number", "case_origin", "work_function", "channel_group")

# I due valori di `Case Origin` che identificano un caso NATO da un transfer.
# 'Transferred from Phone' = ceduto dalla voce; 'Transferred from Live Agent' =
# ceduto da chat/messaging.
ORIGINE_VOICE = "Transferred from Phone"
ORIGINE_BOF = "Transferred from Live Agent"
ORIGINI_TRANSFER = (ORIGINE_VOICE, ORIGINE_BOF)

_WS = re.compile(r"\s+")
_PUNCT = ("(", ")", "%", "–", "-", "/", "|", ".", ",", ":")


def norm(name: str) -> str:
    """Come core/normalize.py: serve solo come secondo tentativo, dopo il nome grezzo."""
    s = str(name).lstrip("﻿").lower().replace("_", " ")
    for p in _PUNCT:
        s = s.replace(p, " ")
    return _WS.sub(" ", s).strip()


def classifica_destinazione(work_function: str | None) -> str:
    """work_function -> Advanced / Basic / Altro.

    `work_function` e' pieno al 100% ma non e' pulito: accanto a Basic/Advanced
    compaiono valori che sono in realta' livelli omni non mappati
    ('Technical Advanced', 'Product Bulk Basic') e valori che non sono un livello
    ('Call Assignment', 'UNMAPPED', 'Project'). I primi si possono ricondurre,
    i secondi NO: finiscono in 'Altro' e si riportano a parte, mai sommati in
    silenzio a una delle due righe che finiscono nel template.
    """
    if not work_function:
        return "Altro"
    v = work_function.strip()
    if v in ("Basic", "Advanced"):
        return v
    low = v.lower()
    if low in ("unmapped", "project", "call assignment"):
        return "Altro"
    if low.endswith(" advanced"):
        return "Advanced"
    if low.endswith(" basic"):
        return "Basic"
    return "Altro"


def canale_origine(case_origin: str | None) -> str:
    if case_origin == ORIGINE_VOICE:
        return "Voice"
    if case_origin == ORIGINE_BOF:
        return "BOF"
    return "?"


# --------------------------------------------------------------------------
# Lettura del workbook
# --------------------------------------------------------------------------

def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(NS + "t")) for si in root]


def _trova_foglio(z: zipfile.ZipFile, nome: str | None) -> tuple[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    fogli = {}
    for s in wb.iter(NS + "sheet"):
        target = rels[s.get(NS_R + "id")]
        fogli[s.get("name")] = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
    candidati = (nome,) if nome else FOGLI_CANDIDATI
    for c in candidati:
        if c in fogli:
            return c, fogli[c]
    raise SystemExit(
        f"nessun foglio {list(candidati)} in questo workbook. Fogli presenti: {sorted(fogli)}"
    )


def _valore(c: ET.Element, ss: list[str]) -> str | None:
    t = c.get("t")
    if t == "inlineStr":
        return "".join(x.text or "" for x in c.iter(NS + "t")) or None
    v = c.find(NS + "v")
    if v is None or v.text is None:
        return None
    if t == "s":
        return ss[int(v.text)]
    return v.text


def _aggancia(intestazioni: dict, contesto: str) -> dict:
    """Campo -> chiave di colonna, agganciando per NOME.

    Prima il nome grezzo, poi quello normalizzato: normalizzare per primo puo'
    far collassare colonne diverse sulla stessa stringa (stessa cautela di
    core/matcher.py). Le chiavi sono lettere per xlsx, indici per csv: qui non
    importa quale delle due, importa che non siano mai posizioni fisse.
    """
    grezzi = {v.strip(): k for k, v in intestazioni.items()}
    normalizzati = {norm(v): k for k, v in intestazioni.items()}
    per_campo: dict = {}
    for campo, alias in CAMPI.items():
        for a in alias:
            if a in grezzi:
                per_campo[campo] = grezzi[a]
                break
            if norm(a) in normalizzati:
                per_campo[campo] = normalizzati[norm(a)]
                break
    mancanti = [c for c in OBBLIGATORIE if c not in per_campo]
    if mancanti:
        raise SystemExit(
            f"{contesto}: mancano le colonne obbligatorie {mancanti}. "
            f"Trovate {len(intestazioni)} intestazioni. "
            f"Se l'export ha cambiato nome, aggiungere l'alias in CAMPI."
        )
    return per_campo


def leggi_workbook(percorso: Path, nome_foglio: str | None = None):
    """Legge il foglio SF di un workbook. Restituisce (sorgente, intestazioni, per_campo, righe).

    ATTENZIONE: l'Omni Report *generato* dalla pipeline non va bene come sorgente.
    Scrive solo gli 11 campi dichiarati in config/columns.yml sotto SF_DATABASE e
    lascia vuote le altre 132 colonne, comprese tutte quelle che servono qui.
    Il posto giusto da cui leggere e' il CSV `SF DATABASE*.csv` di partenza.
    Questa funzione resta per i workbook legacy, riempiti incollando il CSV intero.
    """
    z = zipfile.ZipFile(percorso)
    nome, parte = _trova_foglio(z, nome_foglio)
    ss = _shared_strings(z)

    intestazioni: dict[str, str] = {}   # lettera -> nome
    per_campo: dict[str, str] = {}      # campo -> lettera
    righe: list[dict[str, str | None]] = []

    prima = True
    for _, el in ET.iterparse(z.open(parte), events=("end",)):
        if el.tag != NS + "row":
            continue
        if prima:
            for c in el.iter(NS + "c"):
                nomecol = _valore(c, ss)
                if nomecol:
                    intestazioni[re.match(r"[A-Z]+", c.get("r")).group(0)] = nomecol
            per_campo = _aggancia(intestazioni, f"{percorso.name} / foglio '{nome}'")
            prima = False
            el.clear()
            continue

        cella: dict[str, str | None] = {}
        for c in el.iter(NS + "c"):
            cella[re.match(r"[A-Z]+", c.get("r")).group(0)] = _valore(c, ss)
        riga = {campo: cella.get(lettera) for campo, lettera in per_campo.items()}
        if any(v is not None for v in riga.values()):
            righe.append(riga)
        el.clear()
    return f"foglio '{nome}'", intestazioni, per_campo, righe


def leggi_csv(percorso: Path):
    """Legge l'export `SF DATABASE*.csv`. Stessa firma di leggi_workbook().

    L'export e' UTF-8 con BOM: `utf-8-sig` lo toglie, altrimenti la prima
    intestazione diventa `﻿Date Viewpoint` e non aggancia piu' niente.
    """
    with percorso.open(newline="", encoding="utf-8-sig") as fh:
        lettore = csv.reader(fh)
        try:
            testata = next(lettore)
        except StopIteration:
            raise SystemExit(f"{percorso.name}: file vuoto")
        intestazioni = {i: nome for i, nome in enumerate(testata) if nome.strip()}
        per_campo = _aggancia(intestazioni, percorso.name)
        righe = []
        for campi in lettore:
            riga = {
                campo: (campi[i].strip() or None) if i < len(campi) else None
                for campo, i in per_campo.items()
            }
            if any(v is not None for v in riga.values()):
                righe.append(riga)
    return "csv", intestazioni, per_campo, righe


def leggi(percorso: Path, nome_foglio: str | None = None):
    if percorso.suffix.lower() == ".csv":
        return leggi_csv(percorso)
    return leggi_workbook(percorso, nome_foglio)


# --------------------------------------------------------------------------
# Analisi
# --------------------------------------------------------------------------

def _num(v) -> float | None:
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _data(valore) -> str:
    """La data arriva in due forme, secondo la sorgente.

    Dal workbook e' un seriale Excel; dal CSV `Date (Range)` e' gia' ISO
    (`2026-08-22`). Non si prova a indovinare formati ambigui tipo `8/4/2026`:
    quelli si leggono sia all'americana sia all'europea e qui non servono.
    """
    if valore in (None, ""):
        return ""
    testo = str(valore).strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", testo):
        return testo
    n = _num(testo)
    if n is None:
        return testo
    # Seriale Excel: 1899-12-30 come giorno zero (il bug bisestile del 1900).
    return (_dt.date(1899, 12, 30) + _dt.timedelta(days=int(n))).isoformat()


def analizza(righe: list[dict]) -> dict:
    transfer = [r for r in righe if r.get("case_origin") in ORIGINI_TRANSFER]

    # Gamba ricevente: la 2x2 che finisce nel template.
    matrice = Counter(
        (classifica_destinazione(r.get("work_function")), canale_origine(r.get("case_origin")))
        for r in transfer
    )

    # Bounce: un caso padre che ha generato piu' di un transfer.
    per_padre = Counter(r["parentid"] for r in transfer if r.get("parentid"))
    padri_multipli = {p for p, n in per_padre.items() if n > 1}

    # Soglia AHT: 10o percentile dei transfer, per pescare le lavorazioni-lampo.
    ahts = sorted(a for a in (_num(r.get("aht")) for r in transfer) if a is not None)
    soglia_aht = ahts[max(0, int(len(ahts) * 0.10) - 1)] if ahts else None

    for r in transfer:
        flag = []
        if _num(r.get("misrouted")):
            flag.append("misrouted")
        if r.get("parentid") in padri_multipli:
            flag.append("bounce")
        if r.get("has_child") == "Yes":
            flag.append("genera_figli")
        a = _num(r.get("aht"))
        if soglia_aht is not None and a is not None and a <= soglia_aht:
            flag.append("aht_basso")
        r["_flag"] = "|".join(flag)
        r["_rischio"] = len(flag)

    return {
        "totale_casi": len(righe),
        "transfer": transfer,
        "matrice": matrice,
        "soglia_aht": soglia_aht,
        "padri_multipli": len(padri_multipli),
        # Gamba cedente, per l'annex: non riconcilia con quella ricevente.
        "ceduti_status": [r for r in righe if r.get("case_status") == "Closed - Transferred"],
        "case_transfer": [r for r in righe if r.get("res_category") == "Case Transfer"],
        "call_transfer": [r for r in righe if r.get("res_category") == "Call Transfer"],
    }


def dimensiona_campione(n_popolazione: int, margine: float = 0.05) -> int:
    """Campione per una proporzione, popolazione finita, 95% di confidenza.

    n = N z^2 p q / (d^2 (N-1) + z^2 p q), con p=q=0.5 (il caso peggiore, quello
    che non richiede di sapere in anticipo quanti invalidi ci sono).
    """
    if n_popolazione <= 0:
        return 0
    z2pq = 1.96 ** 2 * 0.25
    n = n_popolazione * z2pq / (margine ** 2 * (n_popolazione - 1) + z2pq)
    return min(n_popolazione, int(n + 0.999))


def marca_campione(transfer: list[dict], seme: str) -> int:
    """Scrive `_campione` su ogni transfer. Restituisce la dimensione del campione.

    Dentro ci vanno il 100% dei casi ad alto rischio piu' un casuale stratificato
    per case type sul resto. L'estrazione e' deterministica a parita' di seme:
    rilanciare lo script due volte deve dare lo stesso campione, altrimenti lo
    scrub gia' fatto non si riaggancia piu'.
    """
    for r in transfer:
        r["_campione"] = r["_rischio"] > 0

    bersaglio = dimensiona_campione(len(transfer))
    resto = [r for r in transfer if not r["_campione"]]
    mancano = bersaglio - (len(transfer) - len(resto))
    if mancano <= 0 or not resto:
        return sum(1 for r in transfer if r["_campione"])

    strati: dict[str, list[dict]] = defaultdict(list)
    for r in resto:
        strati[r.get("case_type") or "(vuoto)"].append(r)

    rng = random.Random(seme)
    quote: list[tuple[str, int]] = []
    for ct, rr in sorted(strati.items()):
        quote.append((ct, min(len(rr), round(mancano * len(rr) / len(resto)))))
    # L'arrotondamento per strato non torna mai esatto: si aggiusta pescando
    # (o restituendo) negli strati piu' capienti, in ordine deterministico.
    scelti: list[dict] = []
    for ct, q in quote:
        rr = sorted(strati[ct], key=lambda r: r.get("case_number") or "")
        rng.shuffle(rr)
        scelti.extend(rr[:q])
    residuo = [r for r in resto if r not in scelti]
    residuo.sort(key=lambda r: r.get("case_number") or "")
    rng.shuffle(residuo)
    while len(scelti) < mancano and residuo:
        scelti.append(residuo.pop())
    for r in scelti[:mancano]:
        r["_campione"] = True
    return sum(1 for r in transfer if r["_campione"])


def rendi_breakdown(a: dict, etichetta: str) -> str:
    m, tr, tot = a["matrice"], a["transfer"], a["totale_casi"]
    out = [f"# Transfer breakdown — {etichetta}", ""]
    out.append(f"Casi totali nell'export: **{tot}**")
    rate = 100 * len(tr) / tot if tot else 0
    out.append(f"Transfer in ingresso: **{len(tr)}** → transfer rate **{rate:.2f}%**")
    out.append("")
    out.append("## Sezione 1 — destinazione x canale di provenienza (gamba ricevente)")
    out.append("")
    out.append("| Queue type | Voice | BOF | TOTALE |")
    out.append("|---|---:|---:|---:|")
    for dest in ("Advanced", "Basic", "Altro"):
        v, b = m[(dest, "Voice")], m[(dest, "BOF")]
        if dest == "Altro" and v + b == 0:
            continue
        out.append(f"| {dest} | {v} | {b} | {v + b} |")
    tv = sum(m[(d, "Voice")] for d in ("Advanced", "Basic", "Altro"))
    tb = sum(m[(d, "BOF")] for d in ("Advanced", "Basic", "Altro"))
    out.append(f"| **TOTALE** | **{tv}** | **{tb}** | **{tv + tb}** |")
    out.append("")
    out.append("Voice = `Case Origin = Transferred from Phone`; "
               "BOF = `Case Origin = Transferred from Live Agent`.")
    out.append("La riga *Altro* raccoglie i `work_function` che non sono un livello "
               "(`Call Assignment`, `UNMAPPED`, `Project`): non va sommata ad Advanced o Basic.")
    out.append("")

    out.append("## Annex — gamba cedente (NON riconcilia con la ricevente)")
    out.append("")
    for nome, righe in (("case_status = Closed - Transferred", a["ceduti_status"]),
                        ("Case Resolution Category = Case Transfer", a["case_transfer"]),
                        ("Case Resolution Category = Call Transfer", a["call_transfer"])):
        if not righe:
            continue
        canali = Counter(r.get("channel_group") for r in righe)
        livelli = Counter(classifica_destinazione(r.get("work_function")) for r in righe)
        out.append(f"- **{nome}**: {len(righe)} casi · canale {dict(canali)} · livello {dict(livelli)}")
    out.append("")
    out.append("Le due gambe sono righe diverse dello stesso export, senza chiave che le "
               "colleghi: i totali non tornano e non devono tornare.")
    out.append("")

    out.append("## Sezione 2 — popolazione da scrubbare")
    out.append("")
    out.append(f"- transfer da scrubbare: **{len(tr)}**")
    hi = [r for r in tr if r["_rischio"] > 0]
    out.append(f"- di cui segnalati ad alto rischio (da scrubbare al 100%): **{len(hi)}**")
    for f, n in Counter(f for r in tr for f in r["_flag"].split("|") if f).most_common():
        out.append(f"    - `{f}`: {n}")
    if a["soglia_aht"] is not None:
        out.append(f"- soglia `aht_basso` (10° percentile dei transfer): {a['soglia_aht']:.2f} min")
    out.append(f"- casi padre con piu' di un transfer figlio (bounce): {a['padri_multipli']}")
    n_camp = sum(1 for r in tr if r.get("_campione"))
    if n_camp:
        out.append(f"- **campione da marcare a mano: {n_camp}** su {len(tr)} "
                   f"({100 * n_camp / len(tr):.0f}%) — alto rischio al 100% piu' un "
                   f"casuale stratificato per case type, per ±5% al 95%")
        restanti = len(tr) - n_camp
        if 0 < restanti <= max(40, len(tr) // 4):
            out.append(f"- NOTA: mancano solo **{restanti}** casi al 100% di copertura. "
                       f"A questi volumi conviene scrubbare tutto: la mail cita Legazpi "
                       f"proprio come esempio di sito al 100%, e un campione va spiegato "
                       f"mentre il 100% no.")
    out.append("")
    out.append("## Controllo per case type")
    out.append("")
    out.append("| Case type | transfer | di cui Advanced | di cui a rischio |")
    out.append("|---|---:|---:|---:|")
    per_tipo = defaultdict(list)
    for r in tr:
        per_tipo[r.get("case_type") or "(vuoto)"].append(r)
    for ct, rr in sorted(per_tipo.items(), key=lambda x: -len(x[1]))[:20]:
        adv = sum(1 for r in rr if classifica_destinazione(r.get("work_function")) == "Advanced")
        ris = sum(1 for r in rr if r["_rischio"] > 0)
        out.append(f"| {ct} | {len(rr)} | {adv} | {ris} |")
    return "\n".join(out) + "\n"


COLONNE_CSV = [
    "case_number", "data", "canale_provenienza", "destinazione", "work_function",
    "omni_level", "case_type", "record_type", "primary_category", "secondary_category",
    "employee", "manager", "aht_min", "misrouted", "has_child", "parentid",
    "queue", "flag_rischio", "punteggio_rischio", "nel_campione", "valid_invalid", "note_scrub",
]


def scrivi_csv(transfer: list[dict], percorso: Path) -> None:
    # Prima chi va marcato, poi per rischio decrescente: chi compila lavora
    # dall'alto e si ferma quando finisce il campione.
    ordinati = sorted(
        transfer,
        key=lambda r: (not r.get("_campione"), -r["_rischio"],
                       r.get("case_type") or "", r.get("case_number") or ""),
    )
    with percorso.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh, delimiter=";")
        w.writerow(COLONNE_CSV)
        for r in ordinati:
            w.writerow([
                r.get("case_number") or "",
                _data(r.get("date_range") or r.get("date_viewpoint")),
                canale_origine(r.get("case_origin")),
                classifica_destinazione(r.get("work_function")),
                r.get("work_function") or "",
                r.get("omni_level") or "",
                r.get("case_type") or "",
                r.get("record_type") or "",
                r.get("primary_category") or "",
                r.get("secondary_category") or "",
                r.get("employee") or "",
                r.get("manager") or "",
                # Virgola decimale: il CSV si apre in Excel in locale italiano.
                (r.get("aht") or "").replace(".", ","),
                r.get("misrouted") or "",
                r.get("has_child") or "",
                r.get("parentid") or "",
                r.get("queue") or "",
                r["_flag"],
                r["_rischio"],
                "SI" if r.get("_campione") else "",
                "", "",
            ])


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", nargs="+", type=Path, help="uno o piu' Omni Report (.xlsm/.xlsx)")
    ap.add_argument("--sheet", default=None, help=f"nome del foglio (default: {FOGLI_CANDIDATI})")
    ap.add_argument("--out", type=Path, default=None, help="cartella dove scrivere gli output")
    ap.add_argument("--etichetta", default=None, help="etichetta per i nomi dei file (default: dai workbook)")
    args = ap.parse_args(argv)

    tutte: list[dict] = []
    for wb in args.workbook:
        if not wb.exists():
            raise SystemExit(f"file non trovato: {wb}")
        sorgente, intestazioni, per_campo, righe = leggi(wb, args.sheet)
        assenti = [c for c in CAMPI if c not in per_campo]
        print(f"[{wb.name}] {sorgente}: {len(intestazioni)} colonne, {len(righe)} righe"
              f"{'  · campi non trovati: ' + ', '.join(assenti) if assenti else ''}", file=sys.stderr)
        # Un workbook generato dalla pipeline ha le intestazioni ma non i dati:
        # meglio dirlo qui che far uscire un report di zeri.
        vuoti = [c for c in ("case_origin", "work_function", "channel_group")
                 if not any(r.get(c) for r in righe)]
        if vuoti:
            raise SystemExit(
                f"{wb.name}: le colonne {vuoti} esistono ma sono TUTTE vuote. "
                f"Se e' un Omni Report generato dalla pipeline e' normale: scrive solo gli "
                f"11 campi del contratto. Usare l'export 'SF DATABASE*.csv' di partenza."
            )
        tutte.extend(righe)

    # Sommando piu' settimane un caso puo' comparire due volte (export che si
    # sovrappongono, o rigenerati). Va tolto qui: piu' avanti si conterebbe due
    # volte sia nel transfer rate sia nel campione di scrub.
    if len(args.workbook) > 1:
        visti: dict[str, dict] = {}
        doppi = 0
        for r in tutte:
            k = r.get("case_number")
            if k and k in visti:
                doppi += 1
                continue
            if k:
                visti[k] = r
        if doppi:
            print(f"[dedup] {doppi} casi presenti in piu' di un export: tenuta la prima "
                  f"occorrenza, {len(visti)} casi distinti", file=sys.stderr)
            tutte = list(visti.values())

    a = analizza(tutte)
    etichetta = args.etichetta or "+".join(w.stem for w in args.workbook)
    # Il seme e' l'etichetta: stesso periodo -> stesso campione a ogni rilancio,
    # cosi' lo scrub gia' fatto si riaggancia.
    marca_campione(a["transfer"], etichetta)
    md = rendi_breakdown(a, etichetta)
    print(md)

    if args.out:
        args.out.mkdir(parents=True, exist_ok=True)
        slug = re.sub(r"[^A-Za-z0-9]+", "_", etichetta).strip("_")
        (args.out / f"breakdown_{slug}.md").write_text(md, encoding="utf-8")
        scrivi_csv(a["transfer"], args.out / f"transfer_casi_{slug}.csv")
        print(f"scritti breakdown_{slug}.md e transfer_casi_{slug}.csv in {args.out}", file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

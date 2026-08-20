"""Controlli di coerenza fra le fonti.

Il preflight verifica ogni dataset per conto suo: le colonne ci sono, i tipi
tornano. Ma i guasti che questo progetto ha trovato nel workbook reale non
stanno *dentro* una fonte: stanno nelle **relazioni** fra fonti.

L'esempio misurato: un agente con skill `HPO␣␣␣*` invece di `HPO` finisce in
`Slot Only Cases` ma non in `Turni`, perché il FILTER di `Helper Turni` usa
l'uguaglianza esatta. Risultato: le regole di malpractice lo valutano per metà.
Nessun errore, nessuna cella rossa — solo una riga che manca.

Ogni controllo dichiara il suo esito: `BLOCCA` ferma la pipeline, `SEGNALA` va a
finire nel preflight e si prosegue. La regola per decidere: blocca se
proseguendo si otterrebbero numeri sbagliati senza accorgersene.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .contract import Dataset
from .names import normalize_name, normalize_skill

BLOCCA = "BLOCCA"
SEGNALA = "SEGNALA"


@dataclass
class Finding:
    check: str
    level: str
    summary: str
    details: list[str] = field(default_factory=list)
    hint: str = ""

    @property
    def blocking(self) -> bool:
        return self.level == BLOCCA


@dataclass
class CoherenceReport:
    findings: list[Finding] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(f.blocking for f in self.findings)

    @property
    def blocking(self) -> list[Finding]:
        return [f for f in self.findings if f.blocking]

    def add(self, f: Finding) -> None:
        self.findings.append(f)

    def render(self, indent: str = "  ") -> list[str]:
        if not self.findings:
            return [f"{indent}nessuna anomalia"]
        out: list[str] = []
        for f in self.findings:
            out.append(f"{indent}[{f.level}] {f.check}: {f.summary}")
            for d in f.details[:12]:
                out.append(f"{indent}    {d}")
            if len(f.details) > 12:
                out.append(f"{indent}    ... e altri {len(f.details) - 12}")
            if f.hint:
                for line in f.hint.splitlines():
                    out.append(f"{indent}    -> {line}")
        return out


# ---------------------------------------------------------------------------

def check_sources(
    *,
    roster_notes=None,
    backoffice_notes=None,
    turni_rows: list[list] | None = None,
    slot_rows: list[list] | None = None,
    week: tuple[date, date] | None = None,
    at_dates: tuple[date, date] | None = None,
    email_agenti: set[str] | None = None,
    contratti: dict[str, str] | None = None,
    wanted_skills: tuple[str, ...] = ("HPO",),
    include_marked: bool = False,
    aliases_available: bool = True,
    at_start_times: list | None = None,
    timezone_offset_hours: float = 0.0,
    week_declared: int | None = None,
    week_inferred: tuple[int, int] | None = None,
    case_owners: dict[str, list] | None = None,
    aliases: dict[str, str] | None = None,
    row_limits: list | None = None,
    last_rows: dict[str, int] | None = None,
    column_stats: dict[str, dict] | None = None,
    date_viewpoint: list | None = None,
    casetype_nuovi: list[tuple[str, str]] | None = None,
    casetype_esclusi: tuple[str, ...] = (),
    casetype_pesi: dict | None = None,
    dup_closed: list | None = None,
    dup_capienze: dict[str, int] | None = None,
    dup_conteggi: dict[str, int] | None = None,
) -> CoherenceReport:
    """Esegue i controlli su ciò che i lettori hanno prodotto.

    Tutti gli argomenti sono opzionali: un controllo che non ha i dati per
    essere fatto viene saltato, non inventato.
    """
    rep = CoherenceReport()

    if not aliases_available and slot_rows is not None:
        rep.add(Finding(
            check="tabella alias nomi non disponibile",
            level=SEGNALA,
            summary="non ho potuto leggere 'Helper Malpractice'!D:E dal template",
            hint=(
                "Il file back-office scrive alcuni nomi diversamente dal roster\n"
                "('alessandro passierello' invece di 'passariello'). Senza la tabella\n"
                "alias quegli agenti risultano orfani, e il controllo qui sotto\n"
                "sugli agenti senza turni segnalera' un problema che non esiste.\n"
                "Prepara il template: la tabella vive li', dov'e' anche il VBA che la usa."
            ),
        ))

    if roster_notes is not None:
        _check_skills(rep, roster_notes, wanted_skills, include_marked)
        _check_blocks(rep, roster_notes)
        _check_non_shift_states(rep, roster_notes)
        _check_stale_cache(rep, roster_notes, "roster turni")
    if backoffice_notes is not None:
        _check_stale_cache(rep, backoffice_notes, "back office")
        _check_requests(rep, backoffice_notes)

    if case_owners and email_agenti is not None:
        for fonte, nomi in case_owners.items():
            _check_case_owners_email(rep, nomi, email_agenti, fonte)

    if case_owners and turni_rows is not None:
        _check_case_owners_senza_turno(rep, case_owners, turni_rows, aliases)

    if row_limits is not None and last_rows:
        _check_row_limits(rep, row_limits, last_rows)

    if column_stats:
        _check_date_ambigue(rep, column_stats)
    if date_viewpoint:
        _check_date_viewpoint(rep, date_viewpoint, week_inferred)
    if casetype_nuovi:
        _check_casetype_nuovi(rep, casetype_nuovi, casetype_esclusi, casetype_pesi)
    if dup_closed:
        _check_settimana_duplicati(rep, dup_closed, week_inferred)
    if dup_capienze and dup_conteggi:
        _check_capienze_duplicati(rep, dup_capienze, dup_conteggi)

    if week_inferred is not None:
        _check_week_declared(rep, week_declared, week_inferred)
    if at_start_times is not None and week is not None:
        _check_at_in_week(rep, at_start_times, week, timezone_offset_hours)

    if turni_rows is not None and slot_rows is not None:
        _check_agent_sets(rep, turni_rows, slot_rows)
    if turni_rows is not None:
        _check_skill_scritta(rep, turni_rows, wanted_skills)
        _check_shift_plausibility(rep, turni_rows)
        _check_days_covered(rep, turni_rows, slot_rows)
        if email_agenti is not None:
            _check_email_agenti(rep, turni_rows, email_agenti)
        if contratti is not None:
            _check_contratti(rep, turni_rows, contratti)
    if at_dates is not None:
        _check_week_alignment(rep, turni_rows, slot_rows, at_dates)

    return rep


# --- 1/2. skill ------------------------------------------------------------

def _check_skills(rep, notes, wanted, include_marked) -> None:
    rep.add(Finding(
        check="skill nel roster",
        level=SEGNALA,
        summary=f"{len(notes.skills_seen)} valori distinti di Skill",
        details=[f"{k!r}: {v} agenti" for k, v in sorted(
            notes.skills_seen.items(), key=lambda kv: -kv[1])],
    ))

    wanted_norm = {normalize_skill(s) for s in wanted}
    quasi = {
        k: v for k, v in notes.skills_seen.items()
        if k not in wanted and normalize_skill(k) in wanted_norm
    }
    if quasi:
        agents = [
            f"{key} (skill {skill!r})"
            for key, skill in sorted(notes.target_agents.items())
            if skill not in wanted
        ]
        rep.add(Finding(
            check="skill che somigliano a quelle richieste",
            level=SEGNALA if include_marked else BLOCCA,
            summary=(
                f"{len(quasi)} valori normalizzano come una skill richiesta ma non "
                f"sono uguali: {', '.join(repr(k) for k in sorted(quasi))}"
            ),
            details=agents,
            hint=(
                "Il FILTER di 'Helper Turni' usa l'uguaglianza esatta "
                "(Turni!$B=\"HPO\"), quindi questi agenti NON entrano in 'Turni' —\n"
                "ma il foglio back-office li include, e le regole di malpractice li\n"
                "valutano a metà: 'Available Cases fuori turno' sì, 'Login in\n"
                "ritardo' contro l'orario di default invece del turno vero.\n"
                "Decidi: se devono entrare, imposta sources.include_marked_skills: true\n"
                "in settings.yml (cambia i numeri); se sono esclusi di proposito,\n"
                "aggiungi la skill esatta a sources.skills."
            )
            if not include_marked else
            "include_marked_skills è attivo: questi agenti sono inclusi in 'Turni'.",
        ))


def _check_blocks(rep, notes) -> None:
    dup = {k: sorted(set(v)) for k, v in notes.agents_in_blocks.items() if len(set(v)) > 1}
    if dup:
        rep.add(Finding(
            check="agenti in piu' blocchi del roster",
            level=BLOCCA,
            summary=f"{len(dup)} agenti hanno una skill richiesta in piu' di un blocco",
            details=[f"{k}: blocchi {', '.join(v)}" for k, v in sorted(dup.items())],
            hint=(
                "Il roster ha blocchi affiancati. Se lo stesso agente e' 'HPO' in due\n"
                "blocchi, produce righe DOPPIE per lo stesso giorno e le ore previste\n"
                "raddoppiano. Nel W30 non succedeva (il blocco 2 non ha HPO)."
            ),
        ))
    if notes.blocks:
        rep.add(Finding(
            check="blocchi del roster",
            level=SEGNALA,
            summary=f"{len(notes.blocks)} blocchi affiancati",
            details=list(notes.blocks),
        ))


def _check_non_shift_states(rep, notes) -> None:
    if not notes.non_shift_states:
        return
    rep.add(Finding(
        check="celle che non sono turni ne' riposi",
        level=SEGNALA,
        summary=", ".join(f"{k}: {len(v)} celle" for k, v in sorted(
            notes.non_shift_states.items())),
        details=[v[0] for v in notes.non_shift_states.values()],
        hint=(
            "'Training' e 'Flessibilita'' non compaiono nel W30, quindi non si sa\n"
            "come il processo manuale li tratti: la pipeline salta quelle righe.\n"
            "Se devono contare come ore previste, va deciso e implementato."
        ),
    ))
    if notes.shift_markers:
        rep.add(Finding(
            check="marcatori sui turni",
            level=SEGNALA,
            summary=", ".join(f"{k}: {v}" for k, v in sorted(notes.shift_markers.items())),
            hint=(
                "Significato non documentato. La pipeline li conserva ma non li usa:\n"
                "se uno di questi vuol dire straordinario, il dato c'e' ancora."
            ),
        ))


def _check_stale_cache(rep, notes, label) -> None:
    if notes.stale_cache:
        rep.add(Finding(
            check=f"valori in cache ({label})",
            level=SEGNALA,
            summary="il file ha formule ma nessuna calcChain",
            hint=(
                "Si leggono i valori memorizzati all'ultimo salvataggio. Se il file\n"
                "e' stato salvato senza ricalcolare, quei valori sono vecchi.\n"
                "Nel dubbio: aprilo in Excel, ricalcola (F9) e risalvalo."
            ),
        ))


def _check_requests(rep, notes) -> None:
    if notes.slot_requests:
        rep.add(Finding(
            check="richieste di cambio turno pendenti",
            level=SEGNALA,
            summary=f"{len(notes.slot_requests)} celle 'REQUEST' nel back office",
            details=notes.slot_requests,
            hint=(
                "Trattate come 'nessuno slot', come fa il processo manuale.\n"
                "Se una REQUEST accettata deve valere come slot, va chiarito."
            ),
        ))


# --- date che si leggono in due modi --------------------------------------

def _check_date_ambigue(rep, column_stats: dict[str, dict]) -> None:
    """Una colonna data che si legge in due modi, senza che nessuno dica quale.

    `8/10/2026` e' il 10 agosto per un export americano e l'8 ottobre per uno
    europeo. Il valore c'e', sembra buono, e la pipeline ne sceglie uno: e' il
    modo peggiore di sbagliare, perche' non lascia traccia. Con `date_format`
    nel contratto il dubbio non esiste e questo controllo non scatta.
    """
    for dataset, stats in sorted(column_stats.items()):
        for st in stats.values():
            if not getattr(st, "ambigue", 0):
                continue
            rep.add(Finding(
                check=f"date ambigue in {dataset}!{st.target_col}",
                level=SEGNALA,
                summary=(
                    f"{st.ambigue} valori di {st.canonical!r} si leggono sia "
                    f"giorno/mese sia mese/giorno"
                ),
                details=[f"esempi: {', '.join(st.esempi_ambigui)}"],
                hint=(
                    "Nessuna delle due letture e' piu' giusta dell'altra guardando la\n"
                    "stringa: al momento vince il primo formato che combacia, cioe' quello\n"
                    "europeo. Se l'export e' americano le date sono sbagliate SENZA errore.\n"
                    "Rimedio: dichiara il formato nel contratto, accanto al campo:\n"
                    '    date_format: "%m/%d/%Y %I:%M:%S %p"'
                ),
            ))


def _check_date_viewpoint(rep, valori: list, week_inferred) -> None:
    """`Date Viewpoint` di SF_DATABASE: un solo lunedi', quello giusto.

    Tableau esporta a granularita' settimanale, quindi la colonna vale lo stesso
    giorno per tutte le righe, ed e' il lunedi' della settimana. Due cose vale
    la pena controllare, perche' entrambe sarebbero silenziose:

    - **Non e' lunedi'** -> quasi sempre significa che il formato della data e'
      stato letto al contrario (10/08 letto come 8 ottobre, che e' un giovedi').
      E' il modo in cui questo controllo intercetta un cambio di formato
      dell'export.
    - **Non e' la settimana del resto del report** -> qualcuno ha scaricato
      l'export SF di una settimana diversa dagli altri. Il report uscirebbe con
      dentro due settimane e nessuna etichetta a dirlo.
    """
    from datetime import date as _date
    from datetime import datetime as _dt

    giorni: dict[_date, int] = {}
    for v in valori:
        if isinstance(v, _dt):
            v = v.date()
        if isinstance(v, _date):
            giorni[v] = giorni.get(v, 0) + 1
    if not giorni:
        return

    if len(giorni) > 1:
        rep.add(Finding(
            check="Date Viewpoint non e' un solo giorno",
            level=SEGNALA,
            summary=f"{len(giorni)} giorni distinti in SF_DATABASE!'Date Viewpoint'",
            details=[f"{d.isoformat()}: {n} righe" for d, n in sorted(giorni.items())],
            hint=(
                "L'export Tableau a granularita' settimanale ne porta uno solo.\n"
                "Piu' giorni = l'export copre piu' di una settimana: il volume e l'AHT\n"
                "finirebbero nello storico sotto un'unica settimana, sommati."
            ),
        ))

    prevalente = max(giorni.items(), key=lambda kv: kv[1])[0]
    if prevalente.weekday() != 0:
        nomi = ("lunedi", "martedi", "mercoledi", "giovedi", "venerdi", "sabato", "domenica")
        rep.add(Finding(
            check="Date Viewpoint non cade di lunedi",
            level=BLOCCA,
            summary=(
                f"{prevalente.isoformat()} e' un {nomi[prevalente.weekday()]}, "
                f"non un lunedi"
            ),
            hint=(
                "La settimana Tableau comincia di lunedi': se qui non lo e', quasi\n"
                "sempre la data e' stata letta al contrario (giorno/mese invece di\n"
                "mese/giorno, o viceversa) e TUTTO lo storico finirebbe sotto la\n"
                "settimana sbagliata.\n"
                "Controlla `date_format` del campo 'Date Viewpoint' in config/columns.yml\n"
                "contro come sono scritte le date nell'export di questa settimana."
            ),
        ))
        return

    if week_inferred:
        iso = prevalente.isocalendar()
        if (iso[0], iso[1]) != tuple(week_inferred):
            rep.add(Finding(
                check="l'export SF e' di un'altra settimana",
                level=BLOCCA,
                summary=(
                    f"SF_DATABASE dice {iso[0]}-W{iso[1]:02d}, il resto del report "
                    f"{week_inferred[0]}-W{week_inferred[1]:02d}"
                ),
                details=[f"Date Viewpoint = {prevalente.isoformat()} (lunedi)"],
                hint=(
                    "Un report con dentro due settimane diverse e nessuna etichetta che\n"
                    "lo dica e' esattamente il genere di file che viene archiviato e poi\n"
                    "riletto come buono.\n"
                    "Riscarica l'export SF della settimana giusta, oppure genera il report\n"
                    "della settimana che l'export copre."
                ),
            ))


def _check_casetype_nuovi(
    rep, nuovi: list[tuple[str, str]], esclusi=(), pesi: dict | None = None
) -> None:
    """Case type nei dati che la lista curata del template non contiene.

    NON entrano da nessuna parte: ne' in 'CaseType Deepdive' (lista curata a
    mano, righe per posizione) ne' nelle heat map di 'AHT Trend WoW' (che leggono
    lo storico, e lo storico ammette solo la lista curata). E' la scelta fatta il
    2026-08-20, dopo aver visto due combinazioni nuove aggiungere due righe alle
    heat map della W33 senza che nessuno le avesse chieste.

    Ma non entrare non deve voler dire non esistere: sono casi veri, lavorati da
    persone vere. Quindi questa segnalazione porta **volume e AHT**, che e' cio'
    che serve per decidere se il case type e' marginale davvero — nella W33
    'Call Assignment' aveva 30 casi, e sarebbe stato sbagliato scartarlo senza
    sapere quanti erano.

    `pesi` e' `{(canale, case type): (volume, aht)}`. Quando non c'e' (un
    chiamante che non li ha calcolati) si elencano le combinazioni senza numeri:
    meno utile, ma non si inventa.
    """
    fuori = {str(t).strip().casefold() for t in esclusi}
    pesi = pesi or {}

    def peso(c, t):
        return pesi.get((c, t)) or pesi.get((c.strip(), t.strip()))

    def riga(c, t):
        p = peso(c, t)
        etichetta = "escluso dal trend per nome" if t.strip().casefold() in fuori else ""
        if p:
            vol, aht = p
            testo = f"{c} | {t}   {vol} casi"
            if aht:
                testo += f", AHT {aht:.1f} min"
        else:
            testo = f"{c} | {t}"
        return testo + (f"   ({etichetta})" if etichetta else "")

    # Dal piu' grosso al piu' piccolo: e' l'ordine in cui si decide.
    ordinati = sorted(
        nuovi, key=lambda ct: -((peso(*ct) or (0, 0))[0]), reverse=False
    )
    dettagli = [riga(c, t) for c, t in ordinati]

    totale = sum((peso(c, t) or (0, 0))[0] for c, t in nuovi)
    summary = (
        f"{len(nuovi)} combinazioni (canale, case type) non sono nella lista "
        f"curata di 'Helper CaseType'"
    )
    if totale:
        summary += f" — {totale} casi in tutto"

    rep.add(Finding(
        check="case type fuori dalla lista curata",
        level=SEGNALA,
        summary=summary,
        details=dettagli,
        hint=(
            "NON entrano nel report: ne' in 'CaseType Deepdive' ne' nelle heat map\n"
            "di 'AHT Trend WoW'. E' voluto — le heat map devono restare confrontabili\n"
            "settimana su settimana, e una riga nuova con un solo dato su dodici\n"
            "colonne sposta anche l'ordinamento di tutte le altre.\n"
            "Se uno di questi ti interessa, aggiungi la coppia (canale, case type) in\n"
            "fondo a 'Helper CaseType' nel template e rilancia: lo storico si riscrive,\n"
            "e il case type compare con TUTTE le settimane che ha nel CSV, non solo\n"
            "da adesso.\n"
            "Guarda il numero di casi prima di decidere: e' il solo modo di sapere se\n"
            "'marginale' e' vero."
        ),
    ))


# --- la settimana e' quella giusta? ---------------------------------------

def _check_week_declared(rep, declared, inferred) -> None:
    """La settimana dichiarata a `--week` coincide con quella dei dati?

    La settimana la decidono i dati: si prende quella in cui `AT_DATASET` sta
    per la gran parte. Il numero digitato serve da controcanto — se non
    coincide, uno dei due e' sbagliato e vale la pena fermarsi: un report
    prodotto sulla settimana giusta ma archiviato col nome di un'altra e' un
    problema che si scopre mesi dopo.
    """
    year, week = inferred
    if declared is None:
        rep.add(Finding(
            check="settimana dedotta dai dati",
            level=SEGNALA,
            summary=f"settimana ISO {week} del {year}",
        ))
        return
    if declared == week:
        rep.add(Finding(
            check="settimana dedotta dai dati",
            level=SEGNALA,
            summary=f"settimana ISO {week} del {year}, coincide con --week {declared}",
        ))
        return
    rep.add(Finding(
        check="settimana dichiarata",
        level=BLOCCA,
        summary=(
            f"hai indicato --week {declared} ma i dati stanno nella settimana "
            f"ISO {week} del {year}"
        ),
        hint=(
            f"Uno dei due e' sbagliato. Se i dati sono giusti, rilancia con\n"
            f"  --week {week}\n"
            f"Se invece volevi davvero la {declared}, l'export non e' quello.\n"
            f"La settimana usata per ritagliare turni e slot e' sempre quella dei\n"
            f"dati, mai quella digitata: cosi' non si producono numeri di una\n"
            f"settimana con il nome di un'altra."
        ),
    ))



def _check_at_in_week(rep, start_times, week, offset_hours) -> None:
    """Quanta attivita' di `AT_DATASET` cade nella settimana scelta.

    E' il controllo che intercetta il `--week` sbagliato. Un po' di sbordo e'
    normale e va solo detto: l'export di AT e' per data Seattle e `Data Milano`
    lo sposta di `offset_hours`, quindi qualche riga finisce oltre il bordo (nel
    W30 sono 3 su 26540). Se invece la maggior parte dei dati e' fuori, la
    settimana e' sbagliata e proseguire produrrebbe un report di un'altra
    settimana.
    """
    from datetime import timedelta

    from .wfmsource import _as_datetime

    lo, hi = week
    dentro = fuori = 0
    giorni_fuori: dict[str, int] = {}
    for v in start_times:
        d = _as_datetime(v)
        if d is None:
            continue
        milano = (d + timedelta(hours=offset_hours)).date()
        if lo <= milano <= hi:
            dentro += 1
        else:
            fuori += 1
            key = milano.isoformat()
            giorni_fuori[key] = giorni_fuori.get(key, 0) + 1

    totale = dentro + fuori
    if not totale:
        return
    quota = fuori / totale

    if dentro == 0:
        rep.add(Finding(
            check="settimana scelta",
            level=BLOCCA,
            summary=(
                f"NESSUNA riga di AT_DATASET cade nella settimana {lo} .. {hi}"
            ),
            details=[f"{k}: {v} righe" for k, v in sorted(giorni_fuori.items())][:10],
            hint=(
                "Il numero passato a --week non corrisponde ai dati.\n"
                "Controlla la settimana, o l'export."
            ),
        ))
    elif quota > 0.20:
        rep.add(Finding(
            check="settimana scelta",
            level=BLOCCA,
            summary=(
                f"il {quota:.0%} delle righe di AT_DATASET cade fuori dalla "
                f"settimana {lo} .. {hi}"
            ),
            details=[f"{k}: {v} righe" for k, v in sorted(giorni_fuori.items())][:10],
            hint=(
                "Troppi dati fuori settimana: probabile --week sbagliato, oppure\n"
                "un export che copre un intervallo diverso da quello atteso."
            ),
        ))
    elif fuori:
        rep.add(Finding(
            check="settimana scelta",
            level=SEGNALA,
            summary=(
                f"{lo} .. {hi} · {dentro} righe dentro, {fuori} fuori "
                f"({quota:.1%})"
            ),
            details=[f"{k}: {v} righe" for k, v in sorted(giorni_fuori.items())][:6],
            hint=(
                f"Normale: l'export di AT e' per data Seattle e Data Milano lo\n"
                f"sposta di {offset_hours:g} ore, quindi qualche riga sborda oltre il\n"
                f"bordo della settimana. Le righe fuori restano in AT_DATASET ma\n"
                f"non hanno turni ne' slot corrispondenti."
            ),
        ))
    else:
        rep.add(Finding(
            check="settimana scelta",
            level=SEGNALA,
            summary=f"{lo} .. {hi} · tutte le {dentro} righe dentro la settimana",
        ))


# --- 2ter. la sezione Duplicate Cases --------------------------------------

def _settimana_prevalente(valori: list):
    """(anno ISO, settimana) della MODA dei giorni, piu' il primo e l'ultimo.

    La moda e non il min/max, per lo stesso motivo per cui `_resolve_week` fa
    cosi' su `AT_DATASET`: un solo caso chiuso a cavallo della mezzanotte del
    lunedi' non deve spostare la settimana di tutto l'export.
    """
    from collections import Counter
    from datetime import date as _date
    from datetime import datetime as _dt

    giorni: list[_date] = []
    for v in valori:
        if isinstance(v, _dt):
            v = v.date()
        if isinstance(v, _date):
            giorni.append(v)
    if not giorni:
        return None, None, None
    settimane = Counter(g.isocalendar()[:2] for g in giorni)
    prevalente = settimane.most_common(1)[0][0]
    return tuple(prevalente), min(giorni), max(giorni)


def _check_settimana_duplicati(rep, closed: list, week_inferred) -> None:
    """L'export duplicati copre la stessa settimana del resto del report?

    Deciso il 2026-08-19: **la stessa**. Lo sfasamento W32/W33 che si vedeva nel
    template era un artefatto del montaggio — quello della W32 era il solo export
    disponibile in quel momento — non un processo che lavora sfasato.

    Quindi BLOCCA, per la stessa ragione di `Date Viewpoint`: un report con
    l'etichetta sbagliata viene archiviato, ed e' peggio di un report che manca.

    Il perimetro si ricava dalle date dei casi (`Date/Time Closed`) e non dalle
    righe del preambolo: il preambolo dice l'intervallo RICHIESTO al report, le
    date dicono quello OTTENUTO, e per l'allineamento conta il secondo — se una
    settimana non ha duplicati chiusi il lunedi', i due non coincidono.
    """
    settimana, primo, ultimo = _settimana_prevalente(closed)
    if settimana is None:
        return
    if not week_inferred:
        rep.add(Finding(
            check="settimana dei duplicati",
            level=SEGNALA,
            summary=(
                f"l'export duplicati copre {settimana[0]}-W{settimana[1]:02d} "
                f"({primo.isoformat()} → {ultimo.isoformat()}), ma non so quella "
                f"del resto del report"
            ),
            hint=(
                "Senza AT_DATASET non c'e' una settimana con cui confrontarla:\n"
                "il controllo si salta invece di inventare un confronto."
            ),
        ))
        return
    if tuple(settimana) == tuple(week_inferred):
        return
    rep.add(Finding(
        check="l'export duplicati e' di un'altra settimana",
        level=BLOCCA,
        summary=(
            f"duplicati {settimana[0]}-W{settimana[1]:02d}, il resto del report "
            f"{week_inferred[0]}-W{week_inferred[1]:02d}"
        ),
        details=[
            f"duplicati (Date/Time Closed): {primo.isoformat()} → {ultimo.isoformat()}",
            f"report (dagli altri export): {week_inferred[0]}-W{week_inferred[1]:02d}",
        ],
        hint=(
            "Riscarica il report duplicati con l'intervallo della settimana giusta,\n"
            "oppure controlla di non aver lasciato in input/ il file della settimana\n"
            "scorsa.\n"
            "Un report che porta la sezione duplicati di un'altra settimana non ha\n"
            "nessuna etichetta che lo dica: i tre fogli DC si leggono come se fossero\n"
            "della settimana in copertina."
        ),
    ))


def _check_capienze_duplicati(rep, capienze: dict[str, int], conteggi: dict[str, int]) -> None:
    """Gli elenchi dei fogli DC hanno posto per quello che c'e' nei dati?

    Gli elenchi sono array dinamici e crescono da se'; le colonne accanto (il
    conteggio, la media, la percentuale) hanno una formula per riga e si fermano
    dove sono state tirate. La voce in eccesso compare **senza nessun numero
    accanto**: presente e invisibile insieme, senza un solo errore. E' lo stesso
    difetto tolto a 'Helper CaseType' ad agosto.

    Come per i limiti di riga, non si aspetta il superamento: si SEGNALA all'80%,
    cosi' le formule si tirano quando c'e' tempo e non nella settimana in cui i
    numeri sono gia' incompleti.
    """
    for etichetta, servono in sorted(conteggi.items()):
        capienza = capienze.get(etichetta)
        if not capienza or not servono:
            continue
        if servono > capienza:
            rep.add(Finding(
                check=f"posto finito per {etichetta}",
                level=BLOCCA,
                summary=f"{servono} da mostrare, il foglio ne tiene {capienza}",
                hint=(
                    f"Le {servono - capienza} voci in eccesso comparirebbero nell'elenco\n"
                    "senza nessun numero accanto: nessun errore, solo celle vuote\n"
                    "dove dovrebbe esserci un conteggio.\n"
                    "Rimedio: tira le formule piu' in basso nel foglio (vedi\n"
                    "docs/piano-duplicates.md §4.3 per le colonne esatte).\n"
                    "La capienza viene letta dal template, quindi appena le tiri\n"
                    "questo controllo se ne accorge da solo."
                ),
            ))
        elif servono > capienza * SOGLIA_ATTENZIONE:
            rep.add(Finding(
                check=f"posto quasi finito per {etichetta}",
                level=SEGNALA,
                summary=f"{servono} su {capienza} ({servono / capienza:.0%})",
                hint=(
                    "Non morde ancora. Ma quando mordera' non lo dira' nessuno: le\n"
                    "voci in eccesso resteranno nell'elenco senza numeri accanto.\n"
                    "Conviene tirare le formule adesso, che c'e' tempo."
                ),
            ))


# --- 3. insiemi di agenti --------------------------------------------------

# Sopra questa quota di riempimento si segnala: il limite non morde ancora, ma la
# settimana in cui mordera' non si vuole scoprirla dai numeri.
SOGLIA_ATTENZIONE = 0.8


def _check_row_limits(rep, row_limits, last_rows: dict[str, int]) -> None:
    """Le righe scritte stanno dentro i limiti scritti nelle formule?

    Molte formule del template leggono intervalli con la riga finale dentro la
    formula (`SUMIFS(AT_DATASET!$P$2:$P$130000, ...)`). Superarla non produce un
    errore: produce **conteggi e medie su un sottoinsieme**, cioe' numeri
    plausibili e piu' bassi del vero.

    E' l'unico guasto di questo progetto che arriva **da solo**, senza che nessuno
    tocchi niente: basta che il volume cresca. Per questo il controllo non aspetta
    il superamento — segnala già all'80%, così il margine si allarga quando c'e'
    tempo e non nella settimana in cui i numeri sono già sbagliati.
    """
    from .templatescan import binding_limits

    vincolanti = binding_limits(list(row_limits))
    for ds, ultima in sorted(last_rows.items()):
        lim = vincolanti.get(ds)
        if lim is None or not ultima:
            continue
        tutti = sorted(
            (l for l in row_limits if l.dataset == ds), key=lambda l: l.max_row
        )
        dettagli = [str(l) for l in tutti if l.max_row < ultima * 2]
        if ultima > lim.max_row:
            rep.add(Finding(
                check=f"formule troppo corte per {ds}",
                level=BLOCCA,
                summary=(
                    f"i dati arrivano a riga {ultima}, ma le formule si fermano a "
                    f"{lim.max_row}"
                ),
                details=dettagli or [str(lim)],
                hint=(
                    "Le righe oltre il limite NON entrano nei calcoli, e non c'e'\n"
                    "nessun errore: escono conteggi e medie su un sottoinsieme.\n"
                    "Allarga l'intervallo nelle formule elencate, oppure — meglio —\n"
                    "riscrivile su colonna intera o sulla tabella, che si\n"
                    "ridimensiona da se' a ogni build. Vedi docs/estendere.md."
                ),
            ))
        elif ultima > lim.max_row * SOGLIA_ATTENZIONE:
            rep.add(Finding(
                check=f"formule vicine al limite per {ds}",
                level=SEGNALA,
                summary=(
                    f"{ultima} righe su un limite di {lim.max_row} "
                    f"({ultima / lim.max_row:.0%})"
                ),
                details=dettagli or [str(lim)],
                hint=(
                    "Non morde ancora. Ma quando mordera' non lo dira' nessuno:\n"
                    "le righe in eccesso resteranno fuori dai calcoli in silenzio.\n"
                    "Conviene allargare adesso, che c'e' tempo."
                ),
            ))


def _check_case_owners_senza_turno(rep, case_owners, turni_rows, aliases) -> None:
    """Chi ha lavorato casi e non ha un turno nel roster.

    La terza direzione dello stesso guasto. Le altre due la guardano dal roster
    ('agenti con slot ma senza turni') e da 'Email Agenti' ('agenti senza
    email'); questa parte dai **casi**, che e' da dove arrivano le persone che
    nel roster HPO non ci sono affatto: un altro team che ha coperto, un
    supervisore, un ingresso nuovo.

    Conseguenza nel report: compaiono in 'Report Agenti' con `Ore previste = 0` e
    la produttivita' non calcolabile, e in `AddLoginRows` il loro orario atteso
    ripiega su `defaultStart` — non sul turno vero, che non esiste.

    Misurato sulla W31: 'lucia serafini' (1 caso, nessun turno HPO).

    Il confronto passa dalla tabella alias, perche' le due fonti scrivono i nomi
    diversamente: il roster ha 'Eleonora Rosa Sissa', Salesforce 'Eleonora
    Sissa'. Senza gli alias questo controllo segnalerebbe quattro persone che
    hanno il loro turno — cioe' sarebbe rumore, e il rumore fa ignorare i
    controlli.
    """
    alias = aliases or {}

    def varianti(nome: str) -> set[str]:
        k = normalize_name(nome)
        return {k, alias.get(k, k)}

    turni: set[str] = set()
    for r in turni_rows:
        if r[0]:
            turni |= varianti(r[0])

    # Il CONTEGGIO dei casi, non solo i nomi: e' quello che distingue un assente
    # con la coda da smaltire da qualcuno che ha lavorato senza turno. Senza il
    # numero, chi legge deve aprire il file per sapere se importa.
    senza: dict[str, dict[str, int]] = {}
    for fonte, nomi in case_owners.items():
        for nome in nomi:
            if not nome:
                continue
            if not (varianti(nome) & turni):
                per_fonte = senza.setdefault(normalize_name(nome), {})
                per_fonte[fonte] = per_fonte.get(fonte, 0) + 1

    if senza:
        rep.add(Finding(
            check="agenti con casi ma senza turno",
            level=SEGNALA,
            summary=f"{len(senza)} nomi hanno lavorato casi ma non sono in 'Turni'",
            details=[
                f"{nome}: "
                + ", ".join(f"{n} casi in {f}" for f, n in sorted(fonti.items()))
                for nome, fonti in sorted(senza.items())
            ],
            hint=(
                "In 'Report Agenti' avranno 'Ore previste' = 0 e nessuna\n"
                "produttivita': non c'e' un turno con cui confrontare le ore.\n"
                "\n"
                "Con POCHI casi (fino a ~5) e' NORMALE e non va corretto: sono\n"
                "agenti assenti che avevano casi in coda, chiusi mentre erano via.\n"
                "Nessuna ora prevista perche' davvero non lavoravano.\n"
                "\n"
                "Da guardare solo se i casi sono molti: allora la persona ha\n"
                "lavorato davvero e la sua riga nel roster manca."
            ),
        ))


def _check_skill_scritta(rep, turni_rows, wanted: tuple[str, ...]) -> None:
    """Ogni riga di `Turni` passa dal FILTER di 'Helper Turni'?

    `Turni!B` ha un solo consumatore, e confronta per uguaglianza esatta:

        Helper Turni!A2 = FILTER(Turni!$A, ..., Turni!$B="HPO")

    Una riga con `HPO                *` sta nel foglio e non entra nel motore:
    nessuna ora prevista, nessun orario di inizio, e "Login in ritardo"
    giudicato contro l'orario di default. E' l'ingresso a meta' da cui e'
    partito tutto il lavoro, e senza questo controllo tornerebbe in silenzio —
    perche' i due insiemi di agenti (Turni e Slot) combaciano, e sono loro che
    l'altro controllo confronta.

    Misurato sulla W31 prima della correzione: `Turni` 37 agenti, 'Helper Turni'
    32.
    """
    fuori: dict[str, set[str]] = {}
    for r in turni_rows:
        skill = r[1]
        if skill not in wanted:
            fuori.setdefault(str(skill), set()).add(normalize_name(r[0]) if r[0] else "?")
    if fuori:
        totale = sum(len(v) for v in fuori.values())
        rep.add(Finding(
            check="skill che il FILTER non riconoscera'",
            level=BLOCCA,
            summary=(
                f"{totale} agenti in 'Turni' hanno un Team/Skill diverso da "
                f"{', '.join(repr(s) for s in wanted)}"
            ),
            details=[
                f"{skill!r}: {', '.join(sorted(agenti))}"
                for skill, agenti in sorted(fuori.items())
            ],
            hint=(
                "'Helper Turni' li scartera' (FILTER con uguaglianza esatta):\n"
                "finirebbero in 'Turni' senza entrare nei calcoli, che e' il\n"
                "peggiore dei due esiti perche' sembrano dentro.\n"
                "Se devono entrare, la skill scritta va portata alla forma\n"
                "canonica; se non devono, non devono nemmeno stare in 'Turni'."
            ),
        ))


def _check_agent_sets(rep, turni_rows, slot_rows) -> None:
    turni = {normalize_name(r[0]) for r in turni_rows if r[0]}
    slot = {r[0] for r in slot_rows if r[0]}
    only_slot = sorted(slot - turni)
    only_turni = sorted(turni - slot)
    if only_slot:
        rep.add(Finding(
            check="agenti con slot ma senza turni",
            level=BLOCCA,
            summary=f"{len(only_slot)} agenti sono in 'Slot Only Cases' ma non in 'Turni'",
            details=only_slot,
            hint=(
                "Senza turni, 'Helper Turni' non li elenca: nessuna ora prevista e\n"
                "nessun orario di inizio. Il VBA giudica il loro 'Login in ritardo'\n"
                "contro l'orario di default, non contro il turno vero — mentre\n"
                "'Available Cases fuori turno' li valuta normalmente.\n"
                "Cause tipiche: skill marcata con asterisco, o nome scritto\n"
                "diversamente fra le due sorgenti (vedi la tabella alias in\n"
                "'Helper Malpractice'!D:E)."
            ),
        ))
    if only_turni:
        rep.add(Finding(
            check="agenti con turni ma senza slot",
            level=SEGNALA,
            summary=f"{len(only_turni)} agenti sono in 'Turni' ma non in 'Slot Only Cases'",
            details=only_turni,
            hint=(
                "Per loro 'Available Cases fuori turno' non scattera' mai: senza\n"
                "finestre di back office, AvailableOutsideSeconds torna 0."
            ),
        ))


# --- 4/10. anagrafiche ----------------------------------------------------

def _check_case_owners_email(rep, nomi: list, email_agenti: set[str], fonte: str) -> None:
    """Chi ha lavorato casi ma non ha una email nel template.

    Girava solo su 'Turni', e cosi' il caso piu' probabile restava fuori: un
    agente che nella settimana **ha chiuso casi** ma non e' nel roster HPO — un
    ingresso nuovo, un supervisore che copre. `Anagrafica` lo elenca da se'
    (SORT(UNIQUE(FILTER(SF_DATABASE!BB)))), ma la sua email e' uno XLOOKUP su
    'Email Agenti' col fallback "": diventa stringa vuota, in silenzio.

    Misurato sulla W31: 'leonardo cengia' e' fra i 37 nomi di SF_DATABASE e non
    fra i 43 di 'Email Agenti'.

    SEGNALA e non BLOCCA: le regole basate sui casi (AHT alto, caso chiuso
    veloce, misrouted) escono comunque, con la email vuota. Sono i confronti col
    turno che si perdono, perche' quelli passano dalla email.
    """
    if not email_agenti:
        return
    missing = sorted({normalize_name(v) for v in nomi if v} - email_agenti)
    if missing:
        rep.add(Finding(
            check=f"agenti di {fonte} senza email",
            level=SEGNALA,
            summary=(
                f"{len(missing)} nomi presenti in {fonte} non sono in 'Email Agenti'"
            ),
            details=missing,
            hint=(
                "In 'Anagrafica' compariranno con la email vuota, perche' lo\n"
                "XLOOKUP su 'Email Agenti' non li trova e ripiega su \"\".\n"
                "Conseguenza: le regole che passano dalla email (login in ritardo,\n"
                "pause, Available fuori turno) non li riguardano, mentre quelle sui\n"
                "casi si. Restano dentro per meta'.\n"
                "Se devono essere nel report, aggiungi la riga in 'Email Agenti';\n"
                "se non devono esserci, allora non dovrebbero avere casi qui."
            ),
        ))


def _check_email_agenti(rep, turni_rows, email_agenti: set[str]) -> None:
    missing = sorted({
        normalize_name(r[0]) for r in turni_rows if r[0]
    } - email_agenti)
    if missing:
        rep.add(Finding(
            check="agenti senza email",
            level=BLOCCA,
            summary=f"{len(missing)} agenti di 'Turni' non sono in 'Email Agenti'",
            details=missing,
            hint=(
                "Il VBA lavora per email: un agente senza email viene scartato da\n"
                "AddATRules, quindi NESSUNA regola di malpractice lo riguarda.\n"
                "La pipeline non puo' inventare le email: aggiungi le righe nel\n"
                "foglio 'Email Agenti' del template."
            ),
        ))


def _check_contratti(rep, turni_rows, contratti: dict[str, str]) -> None:
    missing = sorted({normalize_name(r[0]) for r in turni_rows if r[0]} - set(contratti))
    if missing:
        rep.add(Finding(
            check="agenti senza contratto",
            level=SEGNALA,
            summary=f"{len(missing)} agenti non hanno un contratto in config/contratti.yml",
            details=missing,
            hint=(
                "'Contratto' non e' derivabile dalla sorgente (con Expected hours=0600\n"
                "esistono sia FT sia PT) ed e' informativo: nessun calcolo lo usa.\n"
                "La colonna resta vuota per questi agenti."
            ),
        ))


# --- 5/6. settimana e giorni ---------------------------------------------

def _check_week_alignment(rep, turni_rows, slot_rows, at_dates) -> None:
    lo, hi = at_dates
    for label, rows, idx in (("Turni", turni_rows, 4), ("Slot Only Cases", slot_rows, 1)):
        if not rows:
            continue
        from .coerce import to_datetime

        dates = sorted({to_datetime(r[idx]).date() for r in rows if r[idx] is not None})
        fuori = [d for d in dates if not (lo <= d <= hi)]
        if fuori:
            rep.add(Finding(
                check=f"settimana di {label}",
                level=BLOCCA,
                summary=(
                    f"{len(fuori)} giorni fuori dall'intervallo di AT_DATASET "
                    f"({lo} .. {hi})"
                ),
                details=[d.isoformat() for d in fuori],
                hint=(
                    "Il classico 'ho caricato i turni della settimana scorsa'.\n"
                    "Le regole confronterebbero attivita' e turni di settimane diverse."
                ),
            ))


def _check_days_covered(rep, turni_rows, slot_rows) -> None:
    if not turni_rows or not slot_rows:
        return
    from .coerce import to_datetime

    dt = {to_datetime(r[4]).date() for r in turni_rows if r[4] is not None}
    ds = {to_datetime(r[1]).date() for r in slot_rows if r[1] is not None}
    diff = (dt - ds) | (ds - dt)
    if diff:
        rep.add(Finding(
            check="giorni coperti dalle due fonti",
            level=SEGNALA,
            summary=f"{len(diff)} giorni presenti in una fonte e non nell'altra",
            details=[
                f"{d.isoformat()}: "
                + ("solo Turni" if d in dt else "solo Slot Only Cases")
                for d in sorted(diff)
            ],
        ))


# --- 8. plausibilita' dei turni ------------------------------------------

def _check_shift_plausibility(rep, turni_rows) -> None:
    bad: list[str] = []
    for r in turni_rows:
        nome, ore, data, stato, inizio, fine = r[0], r[3], r[4], r[5], r[6], r[7]
        if stato != "LAVORA":
            continue
        if ore is None or inizio is None or fine is None:
            bad.append(f"{nome} {data}: turno LAVORA con campi vuoti")
            continue
        if ore <= 0:
            bad.append(f"{nome} {data}: {ore} ore")
        elif ore > 16:
            bad.append(f"{nome} {data}: {ore} ore (oltre 16)")
        if fine <= inizio:
            bad.append(
                f"{nome} {data}: fine {fine:.4f} <= inizio {inizio:.4f} "
                f"(turno a cavallo della mezzanotte?)"
            )
    if bad:
        rep.add(Finding(
            check="turni implausibili",
            level=BLOCCA,
            summary=f"{len(bad)} righe con orari o durate fuori scala",
            details=bad,
            hint=(
                "Un turno interpretato male diventa ore previste sbagliate, che e'\n"
                "esattamente l'errore silenzioso che la pipeline deve evitare."
            ),
        ))


def dataset_uses_wfm(dataset: Dataset) -> bool:
    return dataset.reader in ("wfm_roster", "wfm_backoffice")

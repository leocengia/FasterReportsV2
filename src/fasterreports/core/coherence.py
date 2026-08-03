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

    if week_inferred is not None:
        _check_week_declared(rep, week_declared, week_inferred)
    if at_start_times is not None and week is not None:
        _check_at_in_week(rep, at_start_times, week, timezone_offset_hours)

    if turni_rows is not None and slot_rows is not None:
        _check_agent_sets(rep, turni_rows, slot_rows)
    if turni_rows is not None:
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


# --- 3. insiemi di agenti --------------------------------------------------

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

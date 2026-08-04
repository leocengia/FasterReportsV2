"""Errori della pipeline.

Regola: ogni errore deve dire *cosa* manca, *dove* lo si cercava e *cosa* si e'
trovato invece. Un errore che non permette di correggere il CSV senza aprire il
codice e' un errore mal scritto.
"""

from __future__ import annotations

# Quanti header elencare in un errore. SF_DATABASE ne ha 143: elencarli tutti
# sposta il messaggio utile fuori dallo schermo.
HEADER_LIST_LIMIT = 30


class PipelineError(Exception):
    """Base di tutti gli errori attesi: la CLI li stampa senza traceback."""


class ContractError(PipelineError):
    """columns.yml malformato o incoerente."""


class SourceError(PipelineError):
    """Il file sorgente non e' leggibile o non ha una riga di intestazione."""


class ColumnResolutionError(PipelineError):
    """Base degli errori di matching header."""

    def __init__(self, dataset: str, canonical: str, message: str) -> None:
        self.dataset = dataset
        self.canonical = canonical
        super().__init__(message)


class MissingColumnError(ColumnResolutionError):
    def __init__(
        self,
        dataset: str,
        canonical: str,
        aliases: list[str],
        source_headers: list[str],
        exact_only: bool = False,
    ) -> None:
        self.aliases = aliases
        self.source_headers = source_headers
        modo = "solo nome esatto" if exact_only else "nome esatto, normalizzato, poi alias"
        alias_txt = ", ".join(repr(a) for a in aliases) if aliases else "nessuno"

        lines = [
            f"Dataset {dataset}: colonna richiesta {canonical!r} non trovata.",
            f"  Modo di match: {modo}",
            f"  Alias provati: {alias_txt}",
        ]
        near = _suggest(canonical, source_headers)
        if near:
            lines.append(f"  Forse intendevi: {', '.join(repr(n) for n in near)}")

        # Ordinati e troncati: SF ha 143 colonne, e un elenco nell'ordine del
        # file e' un muro di testo in cui non si trova niente.
        shown = sorted(source_headers, key=str.casefold)[:HEADER_LIST_LIMIT]
        rest = len(source_headers) - len(shown)
        lines.append(
            f"  Header presenti nel sorgente ({len(source_headers)}), in ordine alfabetico:"
        )
        suffix = f" ... e altri {rest}" if rest > 0 else ""
        lines.append(f"    {', '.join(repr(h) for h in shown)}{suffix}")

        super().__init__(dataset, canonical, "\n".join(lines))


class AmbiguousColumnError(ColumnResolutionError):
    def __init__(
        self,
        dataset: str,
        canonical: str,
        candidates: list[str],
        via: str,
    ) -> None:
        self.candidates = candidates
        super().__init__(
            dataset,
            canonical,
            f"Dataset {dataset}: colonna {canonical!r} ambigua.\n"
            f"  {len(candidates)} colonne sorgente collassano sullo stesso nome "
            f"(match via {via}): {', '.join(repr(c) for c in candidates)}\n"
            f"  Nessuna corrisponde esattamente al nome canonico, quindi non c'e' modo\n"
            f"  di scegliere senza indovinare.\n"
            f"  Rimedio: in config/columns.yml aggiungi come alias il nome esatto della\n"
            f"  colonna giusta, oppure imposta match: exact se il nome canonico e' corretto.",
        )


class DuplicateTargetError(ContractError):
    """Due campi dello stesso dataset scrivono nella stessa colonna target."""


class CoercionError(PipelineError):
    def __init__(
        self,
        dataset: str,
        canonical: str,
        dtype: str,
        bad_ratio: float,
        limit: float,
        examples: list[str],
    ) -> None:
        self.canonical = canonical
        super().__init__(
            f"Dataset {dataset}: colonna {canonical!r} attesa come {dtype}, ma il "
            f"{bad_ratio:.1%} dei valori non e' convertibile (limite {limit:.1%}).\n"
            f"  Esempi di valori rifiutati: {', '.join(repr(e) for e in examples[:8])}\n"
            f"  Probabile causa: la colonna e' stata mappata su quella sbagliata, "
            f"oppure l'export ha cambiato formato/locale.",
        )


class EmptyDatasetError(PipelineError):
    pass


def _suggest(target: str, headers: list[str], limit: int = 3) -> list[str]:
    """Header piu' simili, per rendere l'errore azionabile."""
    import difflib

    return difflib.get_close_matches(target, headers, n=limit, cutoff=0.6)

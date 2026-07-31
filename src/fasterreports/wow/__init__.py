"""WOW AHT Trend by CT.

Ancora da implementare: vedi src/fasterreports/wow/README.md per l'ordine dei
passi e il perche' non si parte dallo svuotamento del workbook.

Quando si arrivera' a `ingest.py`, il matching degli header NON va riscritto:
si usa `fasterreports.core`. E' lo stesso problema dell'Omni Report — colonne
che cambiano posizione fra un export e l'altro — e i due report leggono export
Salesforce con colonne in comune (`Case Origin (group)`, `Case Type`,
`Case AHT (mins)`).
"""

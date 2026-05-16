#!/usr/bin/env python3
"""Retire JANVIER de Tableau_de_bord_Production.html — ne garde que MAI."""
import re
from pathlib import Path

PATH = Path(__file__).resolve().parent / "Tableau_de_bord_Production.html"
html = PATH.read_text(encoding="utf-8")

# EXCEL_SAISIE_DATA : dates *-05-2026 uniquement
m_ex = re.search(r"window\.EXCEL_SAISIE_DATA = (\{[\s\S]*?\});\s*\n", html)
if not m_ex:
    raise SystemExit("EXCEL_SAISIE_DATA introuvable")
import json

excel = json.loads(m_ex.group(1))
excel_mai = {k: v for k, v in excel.items() if k.endswith("-05-2026")}
excel_json = json.dumps(excel_mai, ensure_ascii=False, separators=(",", ":"))
html = html[: m_ex.start()] + f"window.EXCEL_SAISIE_DATA = {excel_json};\n" + html[m_ex.end() :]

# PRODUCTION : MAI seulement
m_prod = re.search(r"const PRODUCTION = (\{[\s\S]*?\})(;\s*const CHART_FONT)", html)
if not m_prod:
    raise SystemExit("PRODUCTION introuvable")
prod = json.loads(m_prod.group(1))
if "MAI" not in prod:
    raise SystemExit("MAI absent de PRODUCTION")
prod_mai = {"MAI": prod["MAI"]}
prod_json = json.dumps(prod_mai, ensure_ascii=False, separators=(",", ":"))
html = html[: m_prod.start(1)] + prod_json + html[m_prod.end(1) :]

PATH.write_text(html, encoding="utf-8")
print(f"EXCEL_SAISIE_DATA: {len(excel_mai)} jours mai (supprimé {len(excel) - len(excel_mai)} janvier)")
print(f"PRODUCTION: MAI — {len(prod_mai['MAI']['daily_sheets'])} feuilles")

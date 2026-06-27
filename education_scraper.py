"""
Collect Finnish educational institutions from Opintopolku.

Flow:
  1. Search toteutukset → collect toteutusOids
  2. GET /external/toteutus/{oid}?hakukohteet=true → collect hakukohde OIDs
  3. GET /external/hakukohde/{oid} for each → build table
  4. Output Excel

Setup:
    pip install requests pandas openpyxl

Run:
    python education_scraper.py
"""

import re
import time
import requests
import pandas as pd
from pathlib import Path
from playwright.sync_api import sync_playwright, Page

# ── FILTERS ───────────────────────────────────────────────────────────────────
INSTITUTION_TYPES = ["yo"]  # yo = yliopisto, amk = AMK, amm = ammattikoulu

# leave empty for all; pick keys from FIELD_CODES below
FIELDS_OF_STUDY = ["tieto- ja viestintätekniikka"]

# "kandidaatti" = skip entries requiring a prior degree (bachelor's only)
# "maisteri"    = only entries requiring a prior degree (master's only)
# ""            = all
DEGREE_LEVEL = "kandidaatti"

OUTPUT_FILE = "out/finnish_education.xlsx"
PAGE_SIZE = 100
# ─────────────────────────────────────────────────────────────────────────────

SEARCH_URL = "https://opintopolku.fi/konfo-backend/external/search/toteutukset-koulutuksittain"
TOTEUTUS_URL = "https://opintopolku.fi/konfo-backend/external/toteutus"
HAKUKOHDE_URL = "https://opintopolku.fi/konfo-backend/external/hakukohde"

FIELD_CODES = {
    "kasvatusalat":                 "kansallinenkoulutusluokitus2016koulutusalataso1_01",
    "taiteet ja kulttuurialat":     "kansallinenkoulutusluokitus2016koulutusalataso1_02",
    "humanistiset alat":            "kansallinenkoulutusluokitus2016koulutusalataso1_03",
    "liiketalous":                  "kansallinenkoulutusluokitus2016koulutusalataso1_04",
    "luonnontieteet":               "kansallinenkoulutusluokitus2016koulutusalataso1_05",
    "tieto- ja viestintätekniikka": "kansallinenkoulutusluokitus2016koulutusalataso1_06",
    "tekniikka":                    "kansallinenkoulutusluokitus2016koulutusalataso1_07",
    "maa- ja metsätalousalat":      "kansallinenkoulutusluokitus2016koulutusalataso1_08",
    "terveys- ja hyvinvointialat":  "kansallinenkoulutusluokitus2016koulutusalataso1_09",
    "palvelualat":                  "kansallinenkoulutusluokitus2016koulutusalataso1_10",
}

COLUMNS = [
    "Toteutuksen nimi",
    "Opiskelupaikan nimi",
    "Paikkakunta",
    "Haku alkaa",
    "Haku päättyy",
    "Toteutuksen kieli",
    "Paikat kaikille",
    "Paikat ensikertalaisille",
    "Ei-ensikertalaisten %",
    "Linkki",
    "Valintaperusteet",
]


def _fi(obj) -> str:
    if isinstance(obj, str):
        return obj
    if isinstance(obj, dict):
        return obj.get("fi") or obj.get("sv") or obj.get("en") or ""
    return ""


def _date(val: str) -> str:
    return val[:10] if val else ""


def fetch_toteutus_oids() -> list[str]:
    params = {
        "page":           1,
        "size":           PAGE_SIZE,
        "lng":            "fi",
        "koulutustyyppi": ",".join(INSTITUTION_TYPES),
        "opetuskieli":    "oppilaitoksenopetuskieli_1",
        "opetustapa":     "opetuspaikkakk_1",
        "hakutapa":       "hakutapa_01",
    }
    if FIELDS_OF_STUDY:
        params["koulutusala"] = ",".join(
            FIELD_CODES[f] for f in FIELDS_OF_STUDY if f in FIELD_CODES
        )

    oids = []
    while True:
        r = requests.get(SEARCH_URL, params=params, timeout=30)
        r.raise_for_status()
        data = r.json()

        page_hits = data.get("hits", [])
        total = data.get("total", 0)
        print(
            f"  page {params['page']} — {len(page_hits)} hits (total {total})")

        for h in page_hits:
            for t in h.get("toteutukset", []):
                oid = t.get("toteutusOid")
                if oid:
                    oids.append(oid)

        if params["page"] * PAGE_SIZE >= total or not page_hits:
            break
        params["page"] += 1
        # time.sleep(0.3)

    return oids


def fetch_toteutus_detail(oid: str) -> dict:
    r = requests.get(f"{TOTEUTUS_URL}/{oid}",
                     params={"hakukohteet": "true"}, timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_hakukohde(oid: str) -> dict:
    r = requests.get(f"{HAKUKOHDE_URL}/{oid}", timeout=30)
    r.raise_for_status()
    return r.json()


def fetch_valintakoe_spots(page, hk_oid: str) -> tuple:
    url = f"https://opintopolku.fi/konfo/fi/hakukohde/{hk_oid}/valintaperuste"
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        text = page.inner_text("body")
        m = re.search(
            r"Valintakoevalinnassa on (\d+) aloituspaikkaa"
            r".*?ensikertalaisille on varattu (\d+)\)",
            text,
        )
        if m:
            return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None, None


def is_yhteishaku(hk: dict) -> bool:
    return hk.get("hakutapaKoodiUri", "").startswith("hakutapa_01")


def build_row(toteutus: dict, hk: dict, hk_oid: str, page) -> dict:
    # application period
    hakuajat = hk.get("hakuajat") or []
    alkaa = _date(hakuajat[0].get("alkaa", "")) if hakuajat else ""
    paattyy = _date(hakuajat[0].get("paattyy", "")) if hakuajat else ""

    # place
    jarjestyspaikka = hk.get("jarjestyspaikka") or {}
    paikka_nimi = _fi(jarjestyspaikka.get("nimi"))
    paikkakunta = _fi((jarjestyspaikka.get("paikkakunta") or {}).get("nimi"))

    # spots — prefer parsed HTML values, fall back to API metadata
    html_kaikille, html_ensi = fetch_valintakoe_spots(page, hk_oid)
    aloituspaikat = (hk.get("metadata") or {}).get("aloituspaikat") or {}
    paikat_kaikille = html_kaikille or aloituspaikat.get("lukumaara")
    paikat_ei_ensi = html_ensi or aloituspaikat.get("ensikertalaisille")
    if paikat_kaikille and paikat_ei_ensi is not None:
        prosentti = round(100 - paikat_ei_ensi / paikat_kaikille * 100, 1)
    else:
        prosentti = ""

    # language from kielivalinta list e.g. ["fi", "sv"]
    kielet = hk.get("kielivalinta") or toteutus.get("kielivalinta") or []
    kieli = ", ".join(kielet) if kielet else ""

    return {
        "Toteutuksen nimi":        _fi(toteutus.get("nimi")),
        "Opiskelupaikan nimi":     paikka_nimi,
        "Paikkakunta":             paikkakunta,
        "Haku alkaa":              alkaa,
        "Haku päättyy":            paattyy,
        "Toteutuksen kieli":       kieli,
        "Paikat kaikille":         paikat_kaikille,
        "Paikat ensikertalaisille": paikat_ei_ensi,
        "Ei-ensikertalaisten %":    prosentti,
        "Linkki":                   f"https://opintopolku.fi/konfo/fi/toteutus/{toteutus.get('oid', '')}",
        "Valintaperusteet":         f"https://opintopolku.fi/konfo/fi/hakukohde/{hk_oid}/valintaperuste",
    }


def main():
    print("Step 1: fetching toteutus OIDs...")
    toteutus_oids = fetch_toteutus_oids()
    print(f"  → {len(toteutus_oids)} toteutukset\n")

    if not toteutus_oids:
        print("No toteutukset found — check filters.")
        return

    print("Step 2: fetching toteutus details + hakukohde OIDs...")
    pairs: list[tuple[dict, str]] = []   # (toteutus_detail, hakukohde_oid)

    for i, oid in enumerate(toteutus_oids, 1):
        try:
            detail = fetch_toteutus_detail(oid)

            for hk in detail.get("hakukohteet", []):
                hk_oid = hk.get("oid")
                if hk_oid:
                    pairs.append((detail, hk_oid))
            print(
                f"  [{i}/{len(toteutus_oids)}] {oid} → {len(detail.get('hakukohteet', []))} hakukohdetta")
        except Exception as e:
            print(f"  [{i}/{len(toteutus_oids)}] {oid} → error: {e}")
        # time.sleep(0.2)

    print(f"\n  → {len(pairs)} hakukohde OIDs total\n")

    print("Step 3: fetching hakukohde details + building table...")
    rows = []
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page()

        for i, (toteutus, hk_oid) in enumerate(pairs, 1):
            try:
                hk = fetch_hakukohde(hk_oid)
                if not is_yhteishaku(hk):
                    print(
                        f"  [{i}/{len(pairs)}] {hk_oid} → skipped (not yhteishaku)")
                    continue
                requires_degree = any(
                    "pohjakoulutusvaatimuskouta_102" in (
                        v.get("koodiUri") or "")
                    for v in hk.get("pohjakoulutusvaatimus") or []
                )
                if DEGREE_LEVEL == "kandidaatti" and requires_degree:
                    print(f"  [{i}/{len(pairs)}] {hk_oid} → skipped (maisteri)")
                    continue
                if DEGREE_LEVEL == "maisteri" and not requires_degree:
                    print(
                        f"  [{i}/{len(pairs)}] {hk_oid} → skipped (kandidaatti)")
                    continue
                rows.append(build_row(toteutus, hk, hk_oid, page))
                print(f"  [{i}/{len(pairs)}] {hk_oid} → ok")
            except Exception as e:
                print(f"  [{i}/{len(pairs)}] {hk_oid} → error: {e}")

        browser.close()

    if not rows:
        print("No rows to write.")
        return

    Path(OUTPUT_FILE).parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows, columns=COLUMNS)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"\nWrote {len(df)} rows to '{OUTPUT_FILE}'")


if __name__ == "__main__":
    main()

"""
Build one Excel sheet of open Finnish vacancies from multiple sources.

Sources pulled (both open JSON, no login or key):
  1. te-palvelut / Job Market Finland  -> broad: public AND many private postings
  2. Kuntarekry                         -> public sector (municipalities, wellbeing counties)

Purely private boards (Duunitori, LinkedIn, Indeed) have no open API. If you
need those too, an Apify "Duunitori scraper" actor can export a CSV you append
to this sheet. See the note at the bottom.

Setup (run once):
    pip install requests pandas openpyxl

Run:
    python finnish_jobs_to_excel.py
"""

import time
import requests
import pandas as pd

# ----------------------------------------------------------------------
# FILTERS
# ----------------------------------------------------------------------
KEYWORD = ""          # search term, "" = no keyword filter (everything)
# e.g. ["Uusimaa"], [] = all of Finland (te-palvelut only)
REGIONS = []
MAX_PER_SRC = 1000        # cap rows per source so the file stays manageable
PAGE_SIZE = 100
OUTPUT_FILE = "finnish_jobs.xlsx"
# ----------------------------------------------------------------------

# Unified columns every row ends up with
UNIFIED = ["source", "id", "title", "employer",
           "location", "deadline", "url", "description"]


def fetch_te_palvelut():
    """Broad public + private feed behind Job Market Finland."""
    base = "https://paikat.te-palvelut.fi/tpt-api/tyopaikat"
    out, start = [], 0
    while len(out) < MAX_PER_SRC:
        params = {"hakusana": KEYWORD, "start": start, "rows": PAGE_SIZE}
        if REGIONS:
            params["alueet"] = ",".join(REGIONS)
        try:
            r = requests.get(base, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"te-palvelut stopped: {e}")
            break

        docs = data.get("response", {}).get("docs") or data.get("docs") or []
        if not docs:
            break

        for d in docs:
            out.append({
                "source":      "te-palvelut",
                "id":          d.get("ilmoitusnumero"),
                "title":       d.get("tehtavanimi") or d.get("otsikko"),
                "employer":    d.get("tyonantajanNimi"),
                "location":    d.get("kunnat") or d.get("alueet"),
                "deadline":    d.get("hakuPaattyy"),
                "url":         d.get("ilmoituksenURL") or d.get("url"),
                "description": d.get("tehtavat"),
            })
        start += PAGE_SIZE
        time.sleep(0.3)
    return out


def fetch_kuntarekry():
    """Public sector feed (returns JSON directly)."""
    url = "https://www.kuntarekry.fi/fi/api/json-jobs/"
    try:
        r = requests.get(url, timeout=30)
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"kuntarekry stopped: {e}")
        return []

    jobs = data if isinstance(data, list) else data.get(
        "jobs") or data.get("results") or []
    out = []
    for d in jobs[:MAX_PER_SRC]:
        # optional keyword filter, since this feed returns the full list
        if KEYWORD and KEYWORD.lower() not in (d.get("title", "") + d.get("description", "")).lower():
            continue
        path = d.get("url", "")
        out.append({
            "source":      "kuntarekry",
            "id":          d.get("id"),
            "title":       d.get("title"),
            "employer":    d.get("organisation"),
            "location":    "",                      # not always present in this feed
            "deadline":    d.get("publication_end"),
            "url":         f"https://www.kuntarekry.fi{path}" if path.startswith("/") else path,
            "description": d.get("description"),
        })
    return out


def main():
    rows = fetch_te_palvelut() + fetch_kuntarekry()
    if not rows:
        print("Nothing came back. Loosen the filters and retry.")
        return

    df = pd.DataFrame(rows, columns=UNIFIED)
    # rough dedupe: same title at same employer
    df = df.drop_duplicates(
        subset=["title", "employer"]).reset_index(drop=True)
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"Wrote {len(df)} listings to {OUTPUT_FILE}")
    print(df["source"].value_counts().to_string())


if __name__ == "__main__":
    main()

# ----------------------------------------------------------------------
# Adding the private boards later:
#   1. On apify.com search "Duunitori scraper", run it, export results as CSV.
#   2. Load it here and append:
#        extra = pd.read_csv("duunitori.csv")
#        extra = extra.rename(columns={...})  # match the UNIFIED columns
#        combined = pd.concat([pd.read_excel(OUTPUT_FILE), extra])
#        combined.drop_duplicates(subset=["title","employer"]).to_excel(OUTPUT_FILE, index=False)
# ----------------------------------------------------------------------

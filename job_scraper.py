"""
Build one Excel sheet of open Finnish vacancies from multiple sources.

Sources pulled:
  1. te-palvelut / Job Market Finland  -> broad: public AND many private postings
     NOTE: Requires a free KIPA-Subscription-Key from KEHA Centre.
     Register at: https://tyomarkkinatori.fi/en/instructions-and-support/partners/interfaces-for-job-postings
     Contact: tmt-rajapinnat@keha-keskus.fi
  2. Kuntarekry                         -> public sector (municipalities, wellbeing counties)

Purely private boards (Duunitori, LinkedIn, Indeed) have no open API. If you
need those too, an Apify "Duunitori scraper" actor can export a CSV you append
to this sheet. See the note at the bottom.

Setup (run once):
    pip install requests pandas openpyxl

Run:
    python job_scraper.py
"""

import json
import requests
import pandas as pd

# ----------------------------------------------------------------------
# CONFIGURATION
# ----------------------------------------------------------------------
# Get your free key from KEHA Centre: tmt-rajapinnat@keha-keskus.fi
KIPA_SUBSCRIPTION_KEY = ""

KEYWORD = ""          # keyword filter applied locally after fetch, "" = no filter
# Region codes for te-palvelut, e.g. ["01"] for Uusimaa; [] = all of Finland
# Full list: https://tyomarkkinatori.fi/dam/jcr:e9f1db95-5aad-4db4-bd7e-2d010479979a/definitions.zip
REGIONS = []
MAX_PER_SRC = 1000        # cap rows per source so the file stays manageable
OUTPUT_FILE = "finnish_jobs.xlsx"
# ----------------------------------------------------------------------

# Unified columns every row ends up with
UNIFIED = ["source", "id", "title", "employer",
           "location", "deadline", "url", "description"]


def fetch_te_palvelut():
    """Broad public + private feed via Job Market Finland (tyomarkkinatori.fi).

    Requires a KIPA-Subscription-Key — register free at KEHA Centre.
    API: POST https://api.ahtp.fi/kipa/p67/v2/jobpostings
    Response is Newline-delimited JSON (one JobPostingV2 object per line).
    """
    if not KIPA_SUBSCRIPTION_KEY:
        print("te-palvelut skipped: set KIPA_SUBSCRIPTION_KEY (see comment at top of file)")
        return []

    url = "https://api.ahtp.fi/kipa/p67/v2/jobpostings"
    headers = {
        "Content-Type": "application/json",
        "KIPA-Subscription-Key": KIPA_SUBSCRIPTION_KEY,
    }
    body = {"onlyStatus": "PUBLISHED"}
    if REGIONS:
        body["regionIn"] = REGIONS

    try:
        r = requests.post(url, json=body, headers=headers, timeout=60)
        r.raise_for_status()
    except Exception as e:
        print(f"te-palvelut stopped: {e}")
        return []

    def _ml(obj, lang="fi"):
        """Extract value from a multilingual dict, falling back to first available."""
        if isinstance(obj, dict):
            return obj.get(lang) or next(iter(obj.values()), None)
        return obj

    out = []
    for line in r.text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            d = json.loads(line)
        except ValueError:
            continue

        title = _ml(d.get("position", {}).get("title"))
        description = _ml(d.get("position", {}).get("jobDescription")) or ""
        if KEYWORD and KEYWORD.lower() not in f"{title} {description}".lower():
            continue

        out.append({
            "source":      "te-palvelut",
            "id":          d.get("metadata", {}).get("externalId"),
            "title":       title,
            "employer":    _ml(d.get("owner", {}).get("company")),
            "location":    ", ".join(d.get("location", {}).get("municipalities") or []),
            "deadline":    d.get("application", {}).get("expires"),
            "url":         _ml(d.get("application", {}).get("url")),
            "description": description,
        })

        if len(out) >= MAX_PER_SRC:
            break

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

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
annuaire_be_v3.py - Base de prospection MAX des bureaux d'etudes energie
Sortie principale : base_be_tertiaire.xlsx (style, multi-onglets, facon RGE-Closer)

Pipeline :
  1. SCRAPING OPQIBI (qualif 1905 par defaut) : email, tel, site, dirigeant,
     SIREN/SIRET, NAF, CA, effectif, assurances, qualifs + statut probatoire 1905.
  2. ENRICHISSEMENT (gratuit, sans cle API) :
     - recherche-entreprises.api.gouv.fr : etat administratif (active/fermee)
     - BODACC (opendatasoft) : procedures collectives
     - sites web : recuperation email pour les fiches sans email
  3. SCORING : points + tiers (Chaud / Tiede / A qualifier / Hors cible /
     En difficulte / Fermee), signal "1905 probatoire" = nouveau qualifie.
  4. EXPORTS : base_be_tertiaire.xlsx + CSV (national/AURA/69) + run_summary.json

Usage :
  pip install requests beautifulsoup4 pandas openpyxl
  python annuaire_be_v3.py                          # full pipeline
  python annuaire_be_v3.py --from-csv base.csv      # reprend un CSV v2, saute le scraping
  python annuaire_be_v3.py --skip-enrich            # sans appels SIRENE/BODACC/sites
  python annuaire_be_v3.py --quals 1905 1911 1717   # elargir
  python annuaire_be_v3.py --scan-ids 1 9500        # fallback brute force
"""

import argparse
import csv
import json
import re
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://www.opqibi.com"
SLEEP = 0.7
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; base-be-tertiaire/3.0; usage pro)"}
DEPS_AURA = {"01", "03", "07", "15", "26", "38", "42", "43", "63", "69", "73", "74"}
REGION_IDS = list(range(1, 31))

TIER_HOT, TIER_WARM, TIER_QUAL = "CHAUD", "TIEDE", "A QUALIFIER"
TIER_OUT, TIER_DIFF, TIER_DEAD = "HORS CIBLE", "EN DIFFICULTE", "FERMEE"
TIER_ORDER = [TIER_HOT, TIER_WARM, TIER_QUAL, TIER_OUT, TIER_DIFF, TIER_DEAD]

# ================================================================ SCRAPING OPQIBI

def list_fiches_by_search(session, qual):
    urls = set()
    for reg in REGION_IDS:
        try:
            r = session.get(f"{BASE}/recherche-resultat",
                            params={"NewRegion": 1, "Region": reg, "Libelle1": qual},
                            headers=HEADERS, timeout=30)
            if r.status_code != 200:
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            found = 0
            for a in soup.select("a[href*='/fiche/']"):
                href = a["href"]
                urls.add(href if href.startswith("http") else BASE + href)
                found += 1
            if found:
                print(f"  region {reg:>2} : {found} fiches (cumul {len(urls)})")
        except requests.RequestException as e:
            print(f"  region {reg} KO : {e}")
        time.sleep(SLEEP)
    return urls


def list_fiches_by_scan(session, start, end, qual_codes):
    urls = set()
    for fid in range(start, end + 1):
        try:
            r = session.get(f"{BASE}/fiche/{fid}", headers=HEADERS, timeout=20)
            if r.status_code == 200 and any(q in r.text for q in qual_codes):
                urls.add(f"{BASE}/fiche/{fid}")
        except requests.RequestException:
            pass
        if fid % 200 == 0:
            print(f"  scan {fid}/{end} - {len(urls)} retenues")
        time.sleep(SLEEP)
    return urls


LABELS = {
    "email": "E-mail", "site": "Site internet", "forme_juridique": "Forme juridique",
    "capital": "Capital social", "siren": "SIREN", "siret": "SIRET",
    "rcs": "Registre du commerce", "naf": "Code NAF",
    "dirigeant": "Personne(s) ayant le pouvoir", "ca": "Chiffre d'Affaires",
    "effectif": "Effectif total", "apparentement": "Apparentement", "assurances": "Assurance",
}


def parse_fiche(html, url):
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    d = {"url_fiche": url}

    h1 = soup.select_one("h1")
    d["nom"] = h1.get_text(strip=True) if h1 else ""

    for key, label in LABELS.items():
        val = ""
        for i, line in enumerate(lines):
            if line.startswith(label):
                for j in range(i + 1, min(i + 4, len(lines))):
                    nxt = lines[j]
                    if any(nxt.startswith(l2) for l2 in LABELS.values()):
                        break
                    if nxt and not nxt.startswith("("):
                        val = nxt if not val else f"{val} {nxt}"
                        if key != "dirigeant":
                            break
                break
        d[key] = val

    tel_a = soup.select_one("a[href^='tel:']")
    d["telephone"] = tel_a.get_text(strip=True) if tel_a else ""
    if not d["telephone"]:
        m = re.search(r"(?:0|\+33\s?)[1-9](?:[\s.\-]?\d{2}){4}", text)
        d["telephone"] = m.group(0) if m else ""

    if "@" not in d.get("email", ""):
        mail_a = soup.select_one("a[href^='mailto:']")
        if mail_a:
            d["email"] = mail_a.get_text(strip=True)
        else:
            m = re.search(r"[\w.\-]+@[\w.\-]+\.\w{2,}", text)
            d["email"] = m.group(0) if m else ""

    m = re.search(r"Coordonn[ée]es\s*\n(.{5,120}?),?\s*(\d{5})\s+([A-ZÀ-Ü][^,\n]+)", text)
    if m:
        d["adresse"], d["code_postal"], d["ville"] = m.group(1).strip(), m.group(2), m.group(3).strip()
    else:
        m = re.search(r"(\d{5})\s+([A-ZÀ-Ü][A-Za-zÀ-ü\-' ]+)", text)
        d["adresse"] = ""
        d["code_postal"] = m.group(1) if m else ""
        d["ville"] = m.group(2).strip() if m else ""

    quals, statut_1905 = set(), ""
    for a in soup.select("a[href*='nomenclature-fiche/']"):
        code = a["href"].rsplit("/", 1)[-1]
        # ne garder que les liens de tableau de qualifs (texte = le code lui-meme),
        # pas les liens de menu/navigation
        if a.get_text(strip=True) != code:
            continue
        quals.add(code)
        if code == "1905" and not statut_1905:
            # remonter les anciens parents jusqu'a trouver le bloc de section
            # qui dit "probatoire" ou "attribuee" (le site n'utilise pas <table>)
            node = a
            for _ in range(6):
                node = getattr(node, "parent", None)
                if node is None:
                    break
                t = node.get_text(" ", strip=True).lower()
                if len(t) > 2500:
                    break
                if "probatoire" in t:
                    statut_1905 = "probatoire"
                    break
                if "attribu" in t:
                    statut_1905 = "definitive"
                    break
    d["qualifications"] = ";".join(sorted(quals))
    d["statut_1905"] = statut_1905
    d["siret"] = re.sub(r"\D", "", d.get("siret", ""))[:14]
    d["siren"] = re.sub(r"\D", "", d.get("siren", ""))[:9]
    return d


def scrape(session, quals, scan_ids):
    print(f"== Scraping OPQIBI (quals {quals}) ==")
    fiches = set()
    if scan_ids:
        # le scan brute force est PRIORITAIRE : la recherche par region du site
        # ignore les parametres (elle renvoie une page par defaut)
        print(f"  Mode scan IDs {scan_ids[0]}..{scan_ids[1]}")
        fiches = list_fiches_by_scan(session, scan_ids[0], scan_ids[1], quals)
    else:
        for q in quals:
            fiches |= list_fiches_by_search(session, q)
    if not fiches:
        print("!! 0 fiche. Relance avec --scan-ids 1 9500")
        return pd.DataFrame()
    print(f"== Parsing de {len(fiches)} fiches ==")
    rows = []
    for i, url in enumerate(sorted(fiches)):
        try:
            r = session.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200:
                rows.append(parse_fiche(r.text, url))
        except requests.RequestException:
            pass
        if (i + 1) % 25 == 0:
            print(f"  {i + 1}/{len(fiches)} - {sum(1 for x in rows if '@' in x.get('email', ''))} emails")
        time.sleep(SLEEP)
    return pd.DataFrame(rows)

# ================================================================ ENRICHISSEMENT

def enrich_sirene(session, df):
    print("== Enrichissement SIRENE (etat administratif) ==")
    etats = []
    for i, siren in enumerate(df["siren"].fillna("")):
        etat = ""
        if len(str(siren)) == 9:
            try:
                r = session.get("https://recherche-entreprises.api.gouv.fr/search",
                                params={"q": siren, "per_page": 1}, headers=HEADERS, timeout=15)
                if r.status_code == 200:
                    res = r.json().get("results", [])
                    if res:
                        etat = res[0].get("etat_administratif", "") or ""
            except requests.RequestException:
                pass
        etats.append(etat)
        if (i + 1) % 50 == 0:
            print(f"  {i + 1}/{len(df)}")
        time.sleep(0.2)
    df["etat_sirene"] = etats
    return df


def enrich_bodacc(session, df):
    print("== Enrichissement BODACC (procedures collectives) ==")
    url = "https://bodacc-datadila.opendatasoft.com/api/records/1.0/search/"
    procs = []
    for i, siren in enumerate(df["siren"].fillna("")):
        proc = ""
        if len(str(siren)) == 9:
            try:
                r = session.get(url, params={"dataset": "annonces-commerciales",
                                             "q": siren, "rows": 5,
                                             "sort": "dateparution"},
                                headers=HEADERS, timeout=15)
                if r.status_code == 200:
                    for rec in r.json().get("records", []):
                        f = rec.get("fields", {})
                        fam = (f.get("familleavis_lib") or "").lower()
                        if "collective" in fam:
                            proc = f"{f.get('familleavis_lib')} ({str(f.get('dateparution', ''))[:10]})"
                            break
            except requests.RequestException:
                pass
        procs.append(proc)
        if (i + 1) % 50 == 0:
            print(f"  {i + 1}/{len(df)}")
        time.sleep(0.3)
    df["procedure_collective"] = procs
    return df


EMAIL_RE = re.compile(r"[\w.\-]+@[\w.\-]+\.[a-z]{2,}", re.I)
BAD_MAIL = ("example", "wixpress", "sentry", "@2x", ".png", ".jpg")


def enrich_emails_sites(session, df):
    print("== Enrichissement emails via sites web ==")
    if "email_source" not in df.columns:
        df["email_source"] = ""
    todo = df[(~df["email"].astype(str).str.contains("@")) & (df["site"].astype(str).str.len() > 3)]
    print(f"  {len(todo)} fiches sans email mais avec site")
    for idx, row in todo.iterrows():
        site = str(row["site"]).strip()
        if not site.startswith("http"):
            site = "https://" + site
        for path in ("", "/contact", "/mentions-legales"):
            try:
                r = session.get(site.rstrip("/") + path, headers=HEADERS, timeout=10)
                if r.status_code != 200:
                    continue
                for mail in EMAIL_RE.findall(r.text):
                    if not any(b in mail.lower() for b in BAD_MAIL):
                        df.at[idx, "email"] = mail
                        df.at[idx, "email_source"] = "site_web"
                        break
                if "@" in str(df.at[idx, "email"]):
                    break
            except requests.RequestException:
                continue
            finally:
                time.sleep(0.3)
    return df

# ================================================================ SCORING

def score_row(r):
    pts, why = 0, []
    quals = str(r.get("qualifications", ""))
    if "1905" in quals:
        pts += 3; why.append("1905")
    if r.get("statut_1905") == "probatoire":
        pts += 2; why.append("probatoire=nouveau qualifie")
    if "@" in str(r.get("email", "")):
        pts += 2; why.append("email")
    if len(str(r.get("telephone", ""))) > 5:
        pts += 1; why.append("tel")
    try:
        eff = int(re.sub(r"\D", "", str(r.get("effectif", ""))) or 0)
    except ValueError:
        eff = 0
    if 1 <= eff <= 20:
        pts += 2; why.append("effectif cible")
    if len(str(r.get("site", ""))) > 3:
        pts += 1; why.append("site")
    return pts, ", ".join(why), eff


def tier_row(r):
    if r.get("etat_sirene") == "C":
        return TIER_DEAD
    if str(r.get("procedure_collective", "")).strip():
        return TIER_DIFF
    if r["effectif_num"] > 100:
        return TIER_OUT
    if r["score"] >= 7:
        return TIER_HOT
    if r["score"] >= 5:
        return TIER_WARM
    return TIER_QUAL

# ================================================================ EXCEL STYLE

def build_excel(df, path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    HDR_FILL = PatternFill("solid", start_color="0E8E73")
    HDR_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    BODY_FONT = Font(name="Arial", size=10)
    TIER_FILL = {
        TIER_HOT:  PatternFill("solid", start_color="FFE0CC"),
        TIER_WARM: PatternFill("solid", start_color="FFF3D6"),
        TIER_QUAL: PatternFill("solid", start_color="E8F1FA"),
        TIER_OUT:  PatternFill("solid", start_color="F2F2F2"),
        TIER_DIFF: PatternFill("solid", start_color="FADBD8"),
        TIER_DEAD: PatternFill("solid", start_color="E6B8B4"),
    }
    thin = Border(bottom=Side(style="thin", color="E3E9EF"))

    cols = [("tier", "Tier", 15), ("score", "Score", 7), ("nom", "Nom", 32),
            ("email", "Email", 30), ("telephone", "Telephone", 14), ("site", "Site", 26),
            ("dirigeant", "Dirigeant", 30), ("ville", "Ville", 18), ("departement", "Dep", 6),
            ("statut_1905", "1905", 11), ("qualifications", "Qualifs", 16),
            ("effectif", "Effectif", 8), ("ca", "Dernier CA", 20),
            ("procedure_collective", "Procedure collective", 22), ("etat_sirene", "SIRENE", 8),
            ("siret", "SIRET", 16), ("naf", "NAF", 8), ("forme_juridique", "Forme", 10),
            ("adresse", "Adresse", 30), ("code_postal", "CP", 7),
            ("assurances", "Assurances", 14), ("score_detail", "Detail score", 34),
            ("url_fiche", "Fiche OPQIBI", 34)]

    wb = Workbook()

    def write_sheet(ws, data, title):
        ws.title = title
        for j, (_, label, width) in enumerate(cols, 1):
            c = ws.cell(row=1, column=j, value=label)
            c.font, c.fill = HDR_FONT, HDR_FILL
            c.alignment = Alignment(horizontal="center")
            ws.column_dimensions[get_column_letter(j)].width = width
        for i, (_, r) in enumerate(data.iterrows(), 2):
            fill = TIER_FILL.get(r["tier"])
            for j, (key, _, _) in enumerate(cols, 1):
                c = ws.cell(row=i, column=j, value=str(r.get(key, "")))
                c.font, c.border = BODY_FONT, thin
                if fill:
                    c.fill = fill
        ws.freeze_panes = "C2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(2, len(data) + 1)}"

    live = df[~df["tier"].isin([TIER_DIFF, TIER_DEAD])]
    dead = df[df["tier"].isin([TIER_DIFF, TIER_DEAD])]
    write_sheet(wb.active, live, "Prospection")
    ws2 = wb.create_sheet(); write_sheet(ws2, dead, "Ecartes")

    st = wb.create_sheet("Stats")
    st.column_dimensions["A"].width = 36; st.column_dimensions["B"].width = 14
    st["A1"] = "Indicateur"; st["B1"] = "Valeur"
    for c in ("A1", "B1"):
        st[c].font, st[c].fill = HDR_FONT, HDR_FILL
    n = len(live) + 1
    rows = [
        ("Bureaux d'etudes (prospection)", f"=COUNTA(Prospection!C2:C{n})"),
        ("Avec email", f'=COUNTIF(Prospection!D2:D{n},"*@*")'),
        ("Avec telephone", f'=COUNTIF(Prospection!E2:E{n},"?*")'),
        ("CHAUD", f'=COUNTIF(Prospection!A2:A{n},"{TIER_HOT}")'),
        ("TIEDE", f'=COUNTIF(Prospection!A2:A{n},"{TIER_WARM}")'),
        ("1905 probatoires (nouveaux qualifies)", f'=COUNTIF(Prospection!J2:J{n},"probatoire")'),
        ("Ecartes (fermees + difficultes)", f"=COUNTA(Ecartes!C2:C{max(2, len(dead) + 1)})"),
    ]
    for i, (k, v) in enumerate(rows, 2):
        st[f"A{i}"] = k; st[f"B{i}"] = v
        st[f"A{i}"].font = st[f"B{i}"].font = BODY_FONT
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)

# ================================================================ MAIN

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quals", nargs="+", default=["1905"])
    ap.add_argument("--scan-ids", nargs=2, type=int, metavar=("START", "END"))
    ap.add_argument("--from-csv", help="reprendre un CSV existant, sauter le scraping")
    ap.add_argument("--skip-enrich", action="store_true")
    args = ap.parse_args()

    session = requests.Session()

    if args.from_csv:
        df = pd.read_csv(args.from_csv, dtype=str).fillna("")
        print(f"CSV repris : {len(df)} lignes")
    else:
        df = scrape(session, args.quals, args.scan_ids)
        if df.empty:
            return

    for col in ("email", "site", "siren", "statut_1905", "email_source",
                "etat_sirene", "procedure_collective"):
        if col not in df.columns:
            df[col] = ""
    if "departement" not in df.columns or df["departement"].eq("").all():
        df["departement"] = df["code_postal"].astype(str).str[:2]

    # dedoublonnage SANS perdre les fiches au siret vide :
    # url_fiche est toujours unique, puis dedup siret uniquement sur les siret valides
    if "url_fiche" in df.columns:
        df = df.drop_duplicates(subset=["url_fiche"])
    if "siret" in df.columns:
        ok = df["siret"].astype(str).str.len() == 14
        df = pd.concat([df[ok].drop_duplicates(subset=["siret"]), df[~ok]],
                       ignore_index=True)

    # filtre post-parse : ne garder que les structures ayant une qualif visee
    # (les fiches aux qualifs vides = parse incomplet, on les garde par securite)
    pat = "|".join(re.escape(q) for q in args.quals)
    quals_str = df["qualifications"].astype(str)
    avant = len(df)
    df = df[quals_str.str.contains(pat, na=False) | (quals_str.str.strip() == "")]
    print(f"Filtre qualifs {args.quals} : {avant} -> {len(df)} structures")

    if not args.skip_enrich:
        df = enrich_sirene(session, df)
        df = enrich_bodacc(session, df)
        df = enrich_emails_sites(session, df)

    scored = df.apply(score_row, axis=1, result_type="expand")
    df["score"], df["score_detail"], df["effectif_num"] = scored[0], scored[1], scored[2]
    df["tier"] = df.apply(tier_row, axis=1)
    df["tier_rank"] = df["tier"].map({t: i for i, t in enumerate(TIER_ORDER)})
    df = df.sort_values(["tier_rank", "score"], ascending=[True, False]).drop(columns="tier_rank")

    out = Path(".")
    build_excel(df, out / "base_be_tertiaire.xlsx")
    df.to_csv(out / "base_be_max_national.csv", index=False, quoting=csv.QUOTE_MINIMAL)
    df[df["departement"].isin(DEPS_AURA)].to_csv(out / "base_be_max_aura.csv", index=False)
    df[df["departement"] == "69"].to_csv(out / "base_be_max_69.csv", index=False)

    summary = {
        "total": len(df),
        "avec_email": int(df["email"].astype(str).str.contains("@").sum()),
        "avec_tel": int((df["telephone"].astype(str).str.len() > 5).sum()),
        "aura": int(df["departement"].isin(DEPS_AURA).sum()),
        "dep_69": int((df["departement"] == "69").sum()),
        "tiers": df["tier"].value_counts().to_dict(),
        "probatoires_1905": int((df["statut_1905"] == "probatoire").sum()),
        "definitives_1905": int((df["statut_1905"] == "definitive").sum()),
        "qualifs_vides": int((df["qualifications"].astype(str).str.strip() == "").sum()),
    }
    (out / "run_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\n== RESULTATS ==")
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    print("\nFichiers : base_be_tertiaire.xlsx (principal) + 3 CSV + run_summary.json")


if __name__ == "__main__":
    main()

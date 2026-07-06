#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
annuaire_be_v2.py - Base de prospection MAX des bureaux d'etudes energie (national)

Strategie :
  1. Enumerer les fiches OPQIBI qualifiees 1905 (audit energetique tertiaire)
     via les pages de resultats par region, fallback scan d'IDs.
  2. Scraper CHAQUE fiche /fiche/{id} : email, telephone, site, SIREN/SIRET,
     NAF, dirigeant + fonction, CA, effectif, assurances, qualifications.
  3. Croiser avec l'open data RGE ADEME (qualifs, coherence) - optionnel.
  4. Exports : base_be_max_national.csv / _aura.csv / _69.csv

Usage :
  pip install requests beautifulsoup4 pandas
  python annuaire_be_v2.py                 # full national
  python annuaire_be_v2.py --quals 1905 1911 1717   # elargir les qualifs
  python annuaire_be_v2.py --scan-ids 1 9500        # fallback brute force IDs

Politesse scraping : 0.7 s entre requetes, User-Agent identifie.
RGPD : le nom du dirigeant est une donnee personnelle -> registre de traitement,
mention d'origine + lien de desinscription dans chaque email envoye.
"""

import argparse
import csv
import re
import time
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

BASE = "https://www.opqibi.com"
SLEEP = 0.7
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; annuaire-be/2.0; usage pro)"}
DEPS_AURA = {"01", "03", "07", "15", "26", "38", "42", "43", "63", "69", "73", "74"}

# Codes region du site OPQIBI (parametre Region= de /recherche-resultat).
# La numerotation exacte peut differer : le script essaie 1..30 et ignore les vides.
REGION_IDS = list(range(1, 31))


# ------------------------------------------------------------------ etape 1 : lister les fiches

def list_fiches_by_search(session, qual: str) -> set:
    """Recupere les URLs de fiches via les pages de resultats par region."""
    urls = set()
    for reg in REGION_IDS:
        page_url = f"{BASE}/recherche-resultat"
        params = {"NewRegion": 1, "Region": reg, "Libelle1": qual}
        try:
            r = session.get(page_url, params=params, headers=HEADERS, timeout=30)
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


def list_fiches_by_scan(session, start: int, end: int, qual_codes) -> set:
    """Fallback brute force : scanne /fiche/{id} et garde celles avec la qualif visee."""
    urls = set()
    for fid in range(start, end + 1):
        url = f"{BASE}/fiche/{fid}"
        try:
            r = session.get(url, headers=HEADERS, timeout=20)
            if r.status_code == 200 and any(q in r.text for q in qual_codes):
                urls.add(url)
        except requests.RequestException:
            pass
        if fid % 200 == 0:
            print(f"  scan {fid}/{end} - {len(urls)} fiches retenues")
        time.sleep(SLEEP)
    return urls


# ------------------------------------------------------------------ etape 2 : parser une fiche

LABELS = {
    "email": "E-mail",
    "site": "Site internet",
    "forme_juridique": "Forme juridique",
    "capital": "Capital social",
    "siren": "SIREN",
    "siret": "SIRET",
    "rcs": "Registre du commerce",
    "naf": "Code NAF",
    "dirigeant": "Personne(s) ayant le pouvoir",
    "ca": "Chiffre d'Affaires",
    "effectif": "Effectif total",
    "apparentement": "Apparentement",
    "assurances": "Assurance",
}


def parse_fiche(html: str, url: str) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text("\n", strip=True)
    lines = [l.strip() for l in text.split("\n") if l.strip()]

    d = {"url_fiche": url}

    h1 = soup.select_one("h1")
    d["nom"] = h1.get_text(strip=True) if h1 else ""

    # champs "label -> valeur ligne suivante"
    for key, label in LABELS.items():
        val = ""
        for i, line in enumerate(lines):
            if line.startswith(label):
                # la valeur est generalement sur la/les lignes suivantes
                for j in range(i + 1, min(i + 4, len(lines))):
                    nxt = lines[j]
                    if any(nxt.startswith(l2) for l2 in LABELS.values()):
                        break
                    if nxt and not nxt.startswith("("):
                        val = nxt if not val else f"{val} {nxt}"
                        if key not in ("dirigeant",):
                            break
                break
        d[key] = val

    # telephone : lien tel: ou regex
    tel_a = soup.select_one("a[href^='tel:']")
    if tel_a:
        d["telephone"] = tel_a.get_text(strip=True)
    else:
        m = re.search(r"(?:0|\+33\s?)[1-9](?:[\s.\-]?\d{2}){4}", text)
        d["telephone"] = m.group(0) if m else ""

    # email : lien mailto en secours
    if not d.get("email") or "@" not in d["email"]:
        mail_a = soup.select_one("a[href^='mailto:']")
        if mail_a:
            d["email"] = mail_a.get_text(strip=True)
        else:
            m = re.search(r"[\w.\-]+@[\w.\-]+\.\w{2,}", text)
            d["email"] = m.group(0) if m else ""

    # adresse + CP + ville (bloc Coordonnees)
    m = re.search(r"Coordonn[ée]es\s*\n(.{5,120}?),?\s*(\d{5})\s+([A-ZÀ-Ü][^,\n]+)", text)
    if m:
        d["adresse"], d["code_postal"], d["ville"] = m.group(1).strip(), m.group(2), m.group(3).strip()
    else:
        m = re.search(r"(\d{5})\s+([A-ZÀ-Ü][A-Za-zÀ-ü\-' ]+)", text)
        d["code_postal"] = m.group(1) if m else ""
        d["ville"] = m.group(2).strip() if m else ""
        d["adresse"] = ""

    # qualifications (codes 4 chiffres dans les tableaux)
    quals = sorted(set(re.findall(r"\b(1[3-9]\d{2}|20\d{2})\b(?=\s*\n?.{0,120}?qualif|\s*Audit|\s*Etude|\s*\|)",
                                  text, flags=re.IGNORECASE)))
    # plus robuste : codes presents dans les liens nomenclature
    quals_links = {a["href"].rsplit("/", 1)[-1] for a in soup.select("a[href*='nomenclature-fiche/']")}
    d["qualifications"] = ";".join(sorted(set(quals) | quals_links))

    # nettoyage siret
    d["siret"] = re.sub(r"\D", "", d.get("siret", ""))[:14]
    return d


# ------------------------------------------------------------------ main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quals", nargs="+", default=["1905"],
                    help="codes qualification OPQIBI a cibler (ex: 1905 1911 1717 1407)")
    ap.add_argument("--scan-ids", nargs=2, type=int, metavar=("START", "END"),
                    help="fallback : scan brute force des IDs de fiches")
    args = ap.parse_args()

    session = requests.Session()
    out = Path(".")

    print(f"== Etape 1 : lister les fiches OPQIBI (quals {args.quals}) ==")
    fiches = set()
    for q in args.quals:
        fiches |= list_fiches_by_search(session, q)

    if not fiches and args.scan_ids:
        print("  Recherche vide -> fallback scan IDs")
        fiches = list_fiches_by_scan(session, args.scan_ids[0], args.scan_ids[1], args.quals)

    if not fiches:
        print("!! Aucune fiche trouvee. Lance avec --scan-ids 1 9500 (long : ~2h)")
        print("   ou ouvre https://www.opqibi.com/recherche-plus pour verifier les parametres.")
        return

    print(f"\n== Etape 2 : scraping de {len(fiches)} fiches ==")
    rows = []
    for i, url in enumerate(sorted(fiches)):
        try:
            r = session.get(url, headers=HEADERS, timeout=30)
            if r.status_code == 200:
                rows.append(parse_fiche(r.text, url))
        except requests.RequestException:
            pass
        if (i + 1) % 25 == 0:
            print(f"  {i + 1}/{len(fiches)} fiches - {sum(1 for x in rows if x.get('email'))} emails")
        time.sleep(SLEEP)

    df = pd.DataFrame(rows)
    df = df.drop_duplicates(subset=["siret"]) if "siret" in df else df
    df["departement"] = df["code_postal"].astype(str).str[:2]

    cols = ["nom", "email", "telephone", "site", "dirigeant", "adresse", "code_postal",
            "ville", "departement", "siren", "siret", "naf", "forme_juridique",
            "capital", "ca", "effectif", "assurances", "qualifications", "url_fiche"]
    df = df[[c for c in cols if c in df.columns]]

    df.to_csv(out / "base_be_max_national.csv", index=False, quoting=csv.QUOTE_MINIMAL)
    df[df["departement"].isin(DEPS_AURA)].to_csv(out / "base_be_max_aura.csv", index=False)
    df[df["departement"] == "69"].to_csv(out / "base_be_max_69.csv", index=False)

    print("\n== RESULTATS ==")
    print(f"Structures : {len(df)}")
    print(f"Avec email : {(df['email'].str.contains('@', na=False)).sum()}")
    print(f"Avec tel   : {(df['telephone'].astype(str).str.len() > 5).sum()}")
    print(f"AURA : {len(df[df['departement'].isin(DEPS_AURA)])} | Rhone 69 : {len(df[df['departement'] == '69'])}")
    print("\nTop 10 departements :")
    print(df["departement"].value_counts().head(10).to_string())
    print("\nFichiers : base_be_max_national.csv / _aura.csv / _69.csv")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Compare le prix du cadencier au prix réellement payé sur les factures.

Usage : controler_factures.py <dossier du client>
Lit toutes les factures du dossier indiqué dans config.json (les PDF connus, et tout format qu'un lecteur
FACTURES de lecteurs_client.py sait lire), garde pour chaque référence le dernier prix facturé, et écrit
controle_factures.csv. Une même facture copiée deux fois ne compte qu'une fois. Sert à deux choses : confirmer l'unité d'un
prix (à la bouteille, au litre...) et repérer un cadencier périmé ou un prix « brut ».
"""
import csv
import importlib
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.dont_write_bytecode = True

import fitz

from lecteurs import cle_ref, lignes_pdf, nombre


def date_iso(j, m, a):
    return f"{a}-{m}-{j}"


def lire_facture(p):
    """Renvoie (fournisseur, date, [(ref, libellé, prix, promo)]) ou None."""
    try:
        doc = fitz.open(p)
    except Exception:
        return None
    texte = "\n".join(page.get_text() for page in doc)
    haut = texte.upper()
    avoir = False
    if "COMPTOIR DU LITTORAL" in haut and "N° FACTURE" in haut:
        four = "Comptoir du Littoral"
        rangs = " | ".join(" ".join(w[4] for w in mots) for mots in lignes_pdf(doc[0]))
        d = re.search(r"\b\d{8} (\d{2})/(\d{2})/(\d{4}) \d+ ", rangs)
    elif "EPISAVEURS" in haut and "FACTURE N°" in haut:
        four = "Episaveurs"
        d = re.search(r"du (\d{2})\.(\d{2})\.(\d{4})", texte)
    elif "SYSCO" in haut and "CODE PDT" in haut:
        four = "Sysco"
        d = re.search(r"Cde:\d+-(\d{2})/(\d{2})/(\d{2})", texte)
    elif "CODE U.V." in " ".join(" ".join(w[4] for w in mots) for mots in lignes_pdf(doc[0])).upper():
        four = "Pro à Pro"
        rangs = " | ".join(" ".join(w[4] for w in mots) for mots in lignes_pdf(doc[0]))
        d = re.search(r"(\d{2})/(\d{2})/(\d{4}) \d{2}-\d{2}-\d+", rangs)
        avoir = bool(re.search(r"\bAvoir\b", texte))
    else:
        return None
    if not d or avoir:
        return None
    j, m, a = d.groups()
    quand = date_iso(j, m, a if len(a) == 4 else "20" + a)
    out = []
    for page in doc:
        lignes = [" ".join(w[4] for w in mots) for mots in lignes_pdf(page)]
        for i, t in enumerate(lignes):
            suivante = lignes[i + 1] if i + 1 < len(lignes) else ""
            if four == "Comptoir du Littoral":
                x = re.match(r"^(\d{1,6}) (.+?) (\d+) (\d+) (\d+\.\d{4})\b", t)
                if x:
                    out.append((x.group(1), x.group(2), float(x.group(5)), False))
            elif four == "Episaveurs":
                x = re.match(r"^\d+/ (\d+) (.+?) [\d,]+ \S+ [\d,]+ \S+ (\d+,\d{3}) ", t)
                if x:
                    out.append((x.group(1), x.group(2), nombre(x.group(3)), suivante.startswith("Promo")))
            elif four == "Sysco":
                x = re.search(r"\b(\d{5}) (.+?) [FSHA] ([\d,]+) ([A-Z]{1,3}) (\d+,\d{3}) \d ", t)
                if x:
                    out.append((x.group(1), x.group(2), nombre(x.group(5)), "PROMO" in t))
            else:
                x = re.match(r"^(\d{4,6}) (\S+) (.+) ([\d,]+) (\d+,\d{3}) \d (\d+,\d{3}) ([\d,]+)$", t)
                if x:
                    out.append((x.group(1), x.group(3), nombre(x.group(5)), False))
    return four, quand, out


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage : controler_factures.py <dossier du client>")
    ICI = Path(sys.argv[1]).expanduser()
    config = json.load(open(ICI / "config.json", encoding="utf-8"))
    lecteurs_client = []
    if (ICI / "lecteurs_client.py").exists():
        sys.path.insert(0, str(ICI))
        lecteurs_client = getattr(importlib.import_module("lecteurs_client"), "FACTURES", [])
    achats = defaultdict(list)
    compte = defaultdict(int)
    deja_lues = set()
    dossier = Path(config.get("factures") or "factures").expanduser()
    if not dossier.is_absolute():
        dossier = ICI / dossier
    for p in sorted(dossier.glob("**/*")):
        # Les dossiers cachés contiennent souvent des doublons ou des pièces écartées
        if not p.is_file() or any(x.startswith(".") for x in p.relative_to(dossier).parts):
            continue
        # Les lecteurs du client voient tous les formats ; le lecteur fourni ne lit que les PDF
        lecteurs_facture = lecteurs_client + ([lire_facture] if p.suffix.lower() == ".pdf" else [])
        try:
            r = next((x for x in (f(p) for f in lecteurs_facture) if x), None)
        except Exception as e:
            print(f"Facture illisible, ignorée : {p.name} ({type(e).__name__} {e})")
            continue
        if not r:
            continue
        four, quand, lignes = r
        signature = (four, quand, tuple(sorted(lignes)))
        if signature in deja_lues:  # copie d'une facture déjà lue
            continue
        deja_lues.add(signature)
        compte[four] += 1
        for ref, lib, prix, promo in lignes:
            achats[(four, cle_ref(ref))].append((quand, prix, promo, p.name))

    tarifs = list(csv.DictReader(open(ICI / "tarifs.csv", encoding="utf-8-sig")))
    # Pour un même fournisseur et une même référence, on garde le cadencier le plus récent
    recent = {}
    for t in tarifs:
        cle = (t["fournisseur"], cle_ref(t["ref"]))
        if cle not in recent or t["date_tarif"] > recent[cle]["date_tarif"]:
            recent[cle] = t

    sortie = []
    for (four, ref), t in recent.items():
        lignes = achats.get((four, ref))
        if not lignes:
            continue
        # Deux prix le même jour : on garde le plus haut, pour ne pas embellir le résultat
        quand, prix, promo, fichier = max(lignes, key=lambda x: x[:2])
        cad = float(t["prix"])
        normaux = [x for x in lignes if not x[2]]
        hp_date, hp_prix = max(normaux, key=lambda x: x[:2])[:2] if normaux else ("", "")
        sortie.append(dict(fournisseur=four, ref=ref, libelle=t["libelle"], date_tarif=t["date_tarif"],
                           prix_cadencier=cad, date_facture=quand, prix_facture=prix,
                           ecart_pct=round((prix - cad) / cad * 100, 1), promo=int(promo),
                           date_hors_promo=hp_date, prix_hors_promo=hp_prix, nb_achats=len(lignes), facture=fichier))
    sortie.sort(key=lambda s: (s["fournisseur"], -abs(s["ecart_pct"])))
    champs = ["fournisseur", "ref", "libelle", "date_tarif", "prix_cadencier", "date_facture", "prix_facture",
              "ecart_pct", "promo", "date_hors_promo", "prix_hors_promo", "nb_achats", "facture"]
    with open(ICI / "controle_factures.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, champs)
        w.writeheader()
        w.writerows(sortie)

    print("Factures lues :", dict(compte))
    print(f"{'Fournisseur':22}{'réf. cadencier':>15}{'vues sur facture':>18}{'même prix':>11}{'écart > 10 %':>14}")
    for four in sorted({f for f, _ in recent}):
        n = sum(1 for f, _ in recent if f == four)
        s = [x for x in sortie if x["fournisseur"] == four]
        print(f"{four:22}{n:>15}{len(s):>18}{sum(abs(x['ecart_pct']) < 0.5 for x in s):>11}"
              f"{sum(abs(x['ecart_pct']) > 10 for x in s):>14}")


if __name__ == "__main__":
    sys.exit(main())

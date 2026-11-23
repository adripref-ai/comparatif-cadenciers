#!/usr/bin/env python3
"""Reporte dans appariements.csv les clics « OK, validé » et « Ne pas valider » faits sur la page publiée.

Usage : reprendre_clics.py <dossier du client> <clics.json>
clics.json est la liste des documents de la collection « validations » de la page, lue par Claude avec
l'outil ArtifactData : [{"produit": "...", "choix": "ok" ou "refus"}, ...].
« ok » valide la proposition de Claude. « refus » classe le produit en « Ne pas comparer ».
Un produit sans clic ne change pas. Quand plus rien n'est à trancher, config.json n'est plus provisoire.
Ensuite : relancer comparer.py, calculer_classeur.py et generer_page.py, puis republier.
"""
import csv
import json
import sys
from collections import OrderedDict
from pathlib import Path

sys.dont_write_bytecode = True

from validations import CHOIX, SANS_VALIDATION, TYPES


def main():
    if len(sys.argv) != 3:
        sys.exit("Usage : reprendre_clics.py <dossier du client> <clics.json>")
    ici = Path(sys.argv[1]).expanduser()
    brut = json.load(open(Path(sys.argv[2]).expanduser(), encoding="utf-8"))
    clics = {}
    for x in brut if isinstance(brut, list) else brut.get("docs", []):
        x = x.get("data", x)
        if x.get("choix") not in ("ok", "refus") or not x.get("produit"):
            sys.exit(f"Clic illisible : {x}. Attendu : produit, et choix « ok » ou « refus ».")
        clics[x["produit"]] = x["choix"]

    chemin = ici / "appariements.csv"
    with open(chemin, encoding="utf-8-sig") as f:
        lecteur = csv.DictReader(f)
        colonnes = list(lecteur.fieldnames)
        lignes = list(lecteur)
    for c in ("proposition", "valide", "commentaire"):
        if c not in colonnes:
            colonnes.append(c)
    par_produit = OrderedDict()
    for a in lignes:
        par_produit.setdefault(a["produit"], []).append(a)

    valides, refuses = [], []
    for nom, groupe in par_produit.items():
        if nom not in clics:
            continue
        premiere = groupe[0]
        propose = TYPES.get((premiere.get("proposition") or premiere["type"]).strip().lower(), "different")
        choix = propose if clics[nom] == "ok" else "different"
        for a in groupe:
            a["proposition"] = propose
            a["type"] = choix
            a["valide"] = "oui"
            a.setdefault("commentaire", "")
        (valides if clics[nom] == "ok" else refuses).append(nom)

    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, colonnes)
        w.writeheader()
        w.writerows(lignes)

    absents = [n for n in clics if n not in par_produit]
    restants = [n for n, g in par_produit.items()
                if TYPES.get(g[0]["type"].strip().lower()) not in SANS_VALIDATION and g[0].get("valide") != "oui"]
    print(f"{len(valides) + len(refuses)} clic(s) repris sur {len(par_produit)} produits : "
          f"{len(valides)} validé(s), {len(refuses)} refusé(s).")
    if refuses:
        print(f"Passés en « {CHOIX['different']} » : " + ", ".join(refuses))
    if absents:
        print("ATTENTION, clics sur des produits absents d'appariements.csv (renommés ou retirés), non repris : "
              + ", ".join(absents))

    cfg_chemin = ici / "config.json"
    config = json.load(open(cfg_chemin, encoding="utf-8"))
    if bool(restants) != config.get("provisoire", True):
        config["provisoire"] = bool(restants)
        json.dump(config, open(cfg_chemin, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    if not restants:
        print("Toutes les correspondances sont validées : le comparatif n'est plus provisoire.")
    else:
        print(f"Encore {len(restants)} produit(s) à trancher : " + ", ".join(restants[:12])
              + (" ..." if len(restants) > 12 else ""))
    print("Ensuite : relancer comparer.py, calculer_classeur.py, generer_page.py, puis republier la page.")


if __name__ == "__main__":
    main()

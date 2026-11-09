#!/usr/bin/env python3
"""Reporte dans appariements.csv les décisions prises par le client dans l'onglet « À valider ».

Usage : reprendre_validations.py <dossier du client> [classeur]
Le classeur est par défaut Comparatif-fournisseurs.xlsx, dans le dossier du client. Pour une feuille
Google Sheets, la télécharger d'abord au format Excel (Fichier, Télécharger, Microsoft Excel) et donner
son chemin. Une case Décision vide rend le produit à la proposition de Claude.
Quand plus rien n'est à trancher, config.json passe de provisoire à validé.
Ensuite : relancer comparer.py, calculer_classeur.py et generer_page.py, puis republier.
"""
import csv
import json
import sys
from collections import OrderedDict
from pathlib import Path

sys.dont_write_bytecode = True

from validations import CHOIX, ONGLET, TYPES, decision_type, lire_decisions


def main():
    if len(sys.argv) not in (2, 3):
        sys.exit("Usage : reprendre_validations.py <dossier du client> [classeur]")
    ici = Path(sys.argv[1]).expanduser()
    classeur = Path(sys.argv[2]).expanduser() if len(sys.argv) == 3 else ici / "Comparatif-fournisseurs.xlsx"
    if not classeur.exists():
        sys.exit(f"Classeur introuvable : {classeur}")
    decisions = lire_decisions(classeur)
    if decisions is None:
        sys.exit(f"Pas d'onglet « {ONGLET} » dans {classeur.name} : classeur d'une version précédente, "
                 "relancer comparer.py puis faire valider.")

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

    erreurs, changes, repris = [], [], 0
    for nom, groupe in par_produit.items():
        if nom not in decisions:
            continue
        texte, commentaire = decisions[nom]
        try:
            choix = decision_type(texte)
        except ValueError:
            erreurs.append(f"{nom} : décision « {texte} » illisible. Choisir {', '.join(CHOIX.values())}.")
            continue
        premiere = groupe[0]
        propose = TYPES.get((premiere.get("proposition") or premiere["type"]).strip().lower(), "different")
        for a in groupe:
            a["proposition"] = propose
            a["commentaire"] = commentaire
            # Case vide : le produit revient à la proposition de Claude
            a["type"] = choix or propose
            a["valide"] = "oui" if choix else ""
        if choix:
            repris += 1
            if choix != propose:
                changes.append(f"  {nom} : {CHOIX[propose]} -> {CHOIX[choix]}")
    if erreurs:
        print("\n".join(erreurs))
        sys.exit(1)

    with open(chemin, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, colonnes)
        w.writeheader()
        w.writerows(lignes)

    absents = [n for n in decisions if n not in par_produit]
    # Les produits proposés identiques ne demandent pas de décision ; les autres, si
    restants = [n for n, g in par_produit.items()
                if TYPES.get(g[0]["type"].strip().lower()) != "identique" and g[0].get("valide") != "oui"]
    print(f"{repris} décision(s) reprise(s) sur {len(par_produit)} produits.")
    if changes:
        print("Choix différents de la proposition de Claude :")
        print("\n".join(changes))
    if absents:
        print("Produits du classeur absents d'appariements.csv, ignorés : " + ", ".join(absents))

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

#!/usr/bin/env python3
"""Prépare les réglages du cadencier de commande dans le dossier du client.

Usage : preparer_cadencier.py <dossier du client>
À lancer après controler_factures.py, qui écrit achats.csv.
- fournisseurs.csv : une ligne par fournisseur, jours de livraison déduits des factures.
- stocks.csv : une ligne par référence achetée, stock à saisir.
Ce que le client a déjà saisi n'est jamais effacé : seules les cases vides sont complétées.
"""
import sys
from pathlib import Path

sys.dont_write_bytecode = True

from lecteurs import cle_ref
from suivi import (CHAMPS_FOURNISSEURS, CHAMPS_STOCKS, DELAI_DEFAUT, ecrire_csv, lire_csv,
                   livraisons_observees)


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage : preparer_cadencier.py <dossier du client>")
    ici = Path(sys.argv[1]).expanduser()
    if not (ici / "achats.csv").exists():
        sys.exit("achats.csv absent : lancer controler_factures.py d'abord.")
    achats = lire_csv(ici / "achats.csv")
    tarifs = lire_csv(ici / "tarifs.csv")

    # Fournisseurs : ceux des tarifs et ceux des factures
    obs = livraisons_observees(achats)
    noms = list(dict.fromkeys([t["fournisseur"] for t in tarifs] + [a["fournisseur"] for a in achats]))
    deja = {f["fournisseur"]: f for f in lire_csv(ici / "fournisseurs.csv")}
    fournisseurs = []
    for nom in noms:
        f = dict.fromkeys(CHAMPS_FOURNISSEURS, "")
        f.update(deja.get(nom, {}), fournisseur=nom)
        o = obs.get(nom)
        if o and not f["jours_livraison"]:
            un_seul = len(o["dates"]) == 1
            f["jours_livraison"] = o["jours"] + (" ?" if un_seul else "")
            if o["par_semaine"] and not f["livraisons_par_semaine"]:
                f["livraisons_par_semaine"] = o["par_semaine"]
            f["source"] = f["source"] or (f"{len(o['dates'])} livraison(s) sur facture" + (", à confirmer" if un_seul else ""))
        if not f["delai_jours"]:
            f["delai_jours"] = DELAI_DEFAUT
        f["source"] = f["source"] or "à renseigner"
        fournisseurs.append(f)
    fournisseurs += [f for n, f in deja.items() if n not in noms]
    ecrire_csv(ici / "fournisseurs.csv", CHAMPS_FOURNISSEURS, fournisseurs)

    # Stocks : une ligne par référence achetée
    stocks = lire_csv(ici / "stocks.csv")
    vues = {(s["fournisseur"], cle_ref(s["ref"])) for s in stocks}
    ajout = 0
    for a in achats:
        cle = (a["fournisseur"], cle_ref(a["ref"]))
        if cle not in vues:
            vues.add(cle)
            stocks.append(dict(fournisseur=a["fournisseur"], ref=a["ref"], libelle=a["libelle"], stock="",
                               unite=a["unite"], date_stock="", conso_semaine=""))
            ajout += 1
    stocks.sort(key=lambda s: (s["fournisseur"], s["libelle"]))
    ecrire_csv(ici / "stocks.csv", CHAMPS_STOCKS, stocks)

    print(f"{'Fournisseur':22}{'jours de livraison':>20}{'par semaine':>13}  source")
    for f in fournisseurs:
        print(f"{f['fournisseur']:22}{f['jours_livraison'] or '-':>20}{str(f['livraisons_par_semaine'] or '-'):>13}  {f['source']}")
    print(f"stocks.csv : {len(stocks)} références achetées, dont {ajout} ajoutées, "
          f"{sum(1 for s in stocks if not str(s['stock']).strip())} sans stock saisi.")


if __name__ == "__main__":
    main()

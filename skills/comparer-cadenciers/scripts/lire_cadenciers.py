#!/usr/bin/env python3
"""Lit les cadenciers d'un client et écrit tarifs.csv dans son dossier.

Usage : lire_cadenciers.py <dossier du client>
Le dossier contient sources.csv (colonnes lecteur, fichier, date). La date est facultative,
au format 08/12/2025 ou 2025-12-08 : elle remplace la date du fichier quand le document n'en porte pas.
Un lecteur propre au client se met dans lecteurs_client.py, dans le même dossier.
Un fichier illisible est signalé et n'empêche pas de lire les autres.
"""
import csv
import importlib
import sys
from pathlib import Path

sys.dont_write_bytecode = True

from lecteurs import COLONNES, LECTEURS, date_iso


def main():
    if len(sys.argv) != 2:
        sys.exit("Usage : lire_cadenciers.py <dossier du client>")
    ici = Path(sys.argv[1]).expanduser()
    lecteurs = dict(LECTEURS)
    if (ici / "lecteurs_client.py").exists():
        sys.path.insert(0, str(ici))
        lecteurs.update(getattr(importlib.import_module("lecteurs_client"), "LECTEURS", {}))

    tout, perdus, illisibles = [], False, []
    print(f"{'Document':58} {'annoncées':>9} {'lues':>5} {'à vérifier':>10}")
    for s in csv.DictReader(open(ici / "sources.csv", encoding="utf-8-sig")):
        if s["lecteur"] not in lecteurs:
            sys.exit(f"Lecteur inconnu : {s['lecteur']}. Connus : {', '.join(sorted(lecteurs))}")
        chemin = Path(s["fichier"]).expanduser()
        if not chemin.is_absolute():
            chemin = ici / chemin
        try:
            date_indiquee = date_iso(s["date"]) if s.get("date") else ""
            lignes, annonces = lecteurs[s["lecteur"]](chemin)
        except Exception as e:  # un fichier illisible ne doit pas arrêter les autres
            illisibles.append(f"{chemin.name} : {type(e).__name__} {e}")
            print(f"{chemin.name[:58]:58} {'ILLISIBLE':>9}")
            continue
        for x in lignes:
            if date_indiquee:
                x["date_tarif"], x["source_date"] = date_indiquee, "indiquée"
        doute = sum(1 for x in lignes if x["confiance"] != "sûr")
        refs = [x["ref"] for x in lignes]
        doubles = sorted({r for r in refs if refs.count(r) > 1})
        if doubles:
            print(f"ATTENTION : références en double dans {chemin.name} : {', '.join(doubles)}. Seule une ligne sera comparée.")
            perdus = True
        print(f"{chemin.name[:58]:58} {annonces:>9} {len(lignes):>5} {doute:>10}")
        perdus |= annonces != len(lignes)
        tout += lignes
    with open(ici / "tarifs.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, COLONNES)
        w.writeheader()
        w.writerows(tout)
    print(f"Total : {len(tout)} lignes dans tarifs.csv")
    a_confirmer = sorted({x["document"] for x in tout if x["source_date"] == "date du fichier"})
    if a_confirmer:
        print("Date prise sur le fichier, à confirmer dans sources.csv :", ", ".join(a_confirmer))
    for x in illisibles:
        print("ILLISIBLE :", x)
    if perdus:
        print("ATTENTION : des lignes annoncées n'ont pas été lues.")
    if perdus or illisibles:
        sys.exit(1)


if __name__ == "__main__":
    main()

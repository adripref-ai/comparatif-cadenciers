#!/usr/bin/env python3
"""Reprend dans stocks.csv les stocks saisis par le client dans son tableur ou sur la page.

Usage : reprendre_stocks.py <dossier du client> <fichier> [--date JJ/MM/AAAA]
<fichier> peut être :
- le classeur rempli par le client, dans Excel, LibreOffice ou Google Sheets (Fichier, Télécharger, Microsoft Excel) ;
- l'onglet « Cadencier de commande » exporté en CSV, ou relu dans Google Sheets et recopié en CSV ;
- le texte copié par le bouton « Copier les stocks » de la page, enregistré en CSV (fournisseur,ref,stock).
Sont repris le stock, sa date et une consommation corrigée à la main. Un stock sans date prend la date
donnée par --date, sinon celle du jour. Une case vide n'efface jamais ce qui est déjà dans stocks.csv.
À lancer avant de reconstruire le classeur, qui repart de stocks.csv.
"""
import csv
import sys
from datetime import date, datetime
from pathlib import Path

sys.dont_write_bytecode = True

from lecteurs import cle_ref, date_iso
from suivi import CHAMPS_STOCKS, ESTIMATION, ecrire_csv, lire_csv, nombre_ou_vide

ONGLET = "Cadencier de commande"
NOMS = {"fournisseur": "Fournisseur", "ref": "Réf.", "stock": "Stock", "date_stock": "Date du stock",
        "conso": "Consommation par semaine", "libelle": "Libellé", "unite": "Unité", "estimation": ESTIMATION}


def lignes_du_fichier(chemin):
    """Lignes sous forme de dictionnaires, avec les clés de NOMS."""
    if chemin.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(chemin, data_only=True)
        if ONGLET not in wb.sheetnames:
            sys.exit(f"Onglet « {ONGLET} » absent de {chemin.name}.")
        brut = [list(r) for r in wb[ONGLET].iter_rows(values_only=True)]
    else:
        texte = chemin.read_text(encoding="utf-8-sig")
        sep = ";" if texte.split("\n", 1)[0].count(";") > texte.split("\n", 1)[0].count(",") else ","
        brut = list(csv.reader(texte.splitlines(), delimiter=sep))
    if not brut:
        return []
    entete = [str(h or "").strip() for h in brut[0]]
    index = {}
    for cle, nom in NOMS.items():
        for i, h in enumerate(entete):
            if h in (nom, cle):
                index[cle] = i
    if "fournisseur" not in index or "ref" not in index or "stock" not in index:
        sys.exit("Colonnes Fournisseur, Réf. et Stock introuvables : est-ce bien le cadencier de commande ?")
    out = []
    for r in brut[1:]:
        x = {cle: (r[i] if i < len(r) else None) for cle, i in index.items()}
        if x["fournisseur"] and str(x["ref"] or "").strip():
            out.append(x)
    return out


def en_date(v, defaut):
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    v = str(v or "").strip()
    if not v:
        return defaut
    try:
        return date_iso(v)
    except ValueError:
        return defaut


def main():
    args = sys.argv[1:]
    defaut = date.today().isoformat()
    if "--date" in args:
        i = args.index("--date")
        defaut = date_iso(args[i + 1])
        del args[i:i + 2]
    if len(args) != 2:
        sys.exit(__doc__)
    ici, chemin = Path(args[0]).expanduser(), Path(args[1]).expanduser()
    stocks = lire_csv(ici / "stocks.csv")
    par_cle = {(s["fournisseur"], cle_ref(s["ref"])): s for s in stocks}
    repris, conso_reprises, inconnues = 0, [], []
    for x in lignes_du_fichier(chemin):
        cle = (str(x["fournisseur"]).strip(), cle_ref(str(x["ref"]).strip()))
        s = par_cle.get(cle)
        if s is None:  # faute de frappe, ou référence pas encore achetée : on ne l'invente pas
            if nombre_ou_vide(x.get("stock")) is not None:
                inconnues.append(f"{cle[0]} {str(x['ref']).strip()}")
            continue
        stock = nombre_ou_vide(x.get("stock"))
        if stock is not None:
            s["stock"] = f"{stock:g}"
            s["date_stock"] = en_date(x.get("date_stock"), defaut)
            repris += 1
        # Consommation : reprise seulement si elle diffère de celle que le classeur avait écrite
        conso, estim = nombre_ou_vide(x.get("conso")), nombre_ou_vide(x.get("estimation"))
        if conso is not None and "estimation" in x and (estim is None or abs(conso - estim) > 0.005):
            s["conso_semaine"] = f"{conso:g}"
            conso_reprises.append(f"{cle[0]} {s['ref']} : {conso:g} par semaine")
    stocks.sort(key=lambda s: (s["fournisseur"], s["libelle"]))
    ecrire_csv(ici / "stocks.csv", CHAMPS_STOCKS, stocks)
    print(f"{repris} stocks repris de {chemin.name} dans stocks.csv.")
    if conso_reprises:
        print("Consommations corrigées à la main, reprises :\n  " + "\n  ".join(conso_reprises))
    if inconnues:
        print("Références absentes de stocks.csv, non reprises (vérifier le fournisseur et la référence) : "
              + ", ".join(inconnues))
    if not repris:
        print("Aucun stock saisi dans ce fichier.")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Construit la page interactive à partir du classeur.

Usage : generer_page.py <dossier du client> [nom du classeur]
Écrit page-comparatif.html dans le dossier du client, prête à publier comme artefact.
Les résultats sont recalculés ici à partir des prix du classeur, avec les mêmes règles que ses
formules : la page ne dépend donc pas d'Excel, et une correction faite dans le classeur est reprise.
"""
import json
import math
import sys
from datetime import date, datetime
from html import escape
from pathlib import Path

import openpyxl

ICI = Path(__file__).parent
# Clair, sombre. Mêmes teintes que dans le classeur.
COULEURS = {"Sysco": ("#1D5FA6", "#7DB8F0"), "Pro à Pro": ("#1E7A3C", "#6FD18E"), "Episaveurs": ("#B45309", "#F0A95E"),
            "Comptoir du Littoral": ("#6D28D9", "#B79CF5"), "Sodiboissons": ("#BE123C", "#F58BA2")}
PALETTE = [("#0F766E", "#5FD3C4"), ("#A16207", "#E8C24F"), ("#9D174D", "#F08CB8"), ("#1E40AF", "#93B4F7"),
           ("#4D7C0F", "#AEDB5F"), ("#7E22CE", "#CE9BF5"), ("#C2410C", "#F5A26F"), ("#155E75", "#6CCBE3")]


def jour(v):
    if isinstance(v, (datetime, date)):
        return v.strftime("%Y-%m-%d")
    return v or ""


def arrondi(x, n):
    """Arrondi comme ROUND dans Excel (moitié vers le haut), pour que page et classeur concordent."""
    return math.floor(x * 10 ** n + 0.5) / 10 ** n


def table(ws, ligne_titres):
    """Renvoie les lignes d'une feuille sous forme de dictionnaires, clés = titres de colonnes."""
    titres = [c.value for c in ws[ligne_titres]]
    for r in ws.iter_rows(min_row=ligne_titres + 1, values_only=True):
        if r[0] is not None:
            yield dict(zip(titres, r))


def main():
    if len(sys.argv) < 2:
        sys.exit("Usage : generer_page.py <dossier du client> [nom du classeur]")
    dossier = Path(sys.argv[1]).expanduser()
    classeur = dossier / (sys.argv[2] if len(sys.argv) > 2 else "Comparatif-fournisseurs.xlsx")
    config = json.load(open(dossier / "config.json", encoding="utf-8"))
    client = config["client"]
    # Une page refaite à partir d'un classeur périmé afficherait d'anciens chiffres
    entrees = [dossier / n for n in ("tarifs.csv", "appariements.csv", "controle_factures.csv", "config.json")]
    plus_recente = max((p for p in entrees if p.exists()), key=lambda p: p.stat().st_mtime, default=None)
    if not classeur.exists() or (plus_recente and plus_recente.stat().st_mtime > classeur.stat().st_mtime):
        sys.exit(f"Classeur absent ou plus ancien que {plus_recente.name if plus_recente else 'ses données'} : "
                 "relancer comparer.py d'abord.")
    wb = openpyxl.load_workbook(classeur)

    wc = wb["Comparatif"]
    rang = next(r for r in range(1, 40) if wc.cell(r, 1).value == "Produit")
    titres = [c.value for c in wc[rang]]
    fournisseurs = titres[titres.index("Famille") + 1:titres.index("Moins cher")]

    offres = {}
    for d in table(wb["Détail"], 1):
        cad, fac = d["Prix cadencier HT"], d["Dernière facture hors promo"]
        dcad, dfac = d["Date du cadencier"], d["Date de la facture"]
        par_facture = fac is not None and dfac is not None and dfac > dcad
        retenu = fac if par_facture else cad
        offres.setdefault(d["Produit"], []).append(dict(
            f=d["Fournisseur"], ref=d["Réf."], lib=d["Libellé du cadencier"], cad=cad, uc=d["Unité de prix"],
            dcad=jour(dcad), fac=fac, dfac=jour(dfac), pu=arrondi(retenu / d["Quantité par unité de prix"], 4),
            origine="facture" if par_facture else "cadencier", fiable=d["Fiabilité"],
            quand=dfac if par_facture else dcad,
            # Une offre reste valable jusqu'à son dernier jour inclus, comme dans le classeur
            expire=not par_facture and isinstance(d.get("Fin de validité"), datetime)
            and d["Fin de validité"].date() < date.today()))

    produits = []
    for p in table(wc, rang):
        liste = offres[p["Produit"]]
        prix = sorted(o["pu"] for o in liste)
        ecart = pct = None
        if p["Type"] == "non comparable":
            gagnant = "non comparable"
        elif len(prix) < 2:
            gagnant = "seul"
        else:
            ecart, pct = prix[1] - prix[0], (prix[1] - prix[0]) / prix[0]
            gagnant = "égalité" if prix[0] == prix[1] else next(
                f for f in fournisseurs for o in liste if o["f"] == f and o["pu"] == prix[0])
        dates = [o.pop("quand") for o in liste]
        mois = arrondi((max(dates) - min(dates)).days / 30.4, 1)
        doutes = sum(o["fiable"] == "à vérifier" for o in liste)
        produits.append(dict(
            nom=p["Produit"], unite=p["Unité"], famille=p["Famille"], gagnant=gagnant, ecart=ecart, pct=pct,
            mois=mois, doutes=doutes, type=p["Type"], note=p["Remarque"] or "",
            solide=(gagnant not in ("non comparable", "seul") and mois <= 3 and doutes == 0
                    and not any(o["expire"] for o in liste)), offres=liste))

    factures = [dict(f=x["Fournisseur"], lib=x["Libellé"], cad=x["Prix cadencier"], dcad=jour(x["Date du cadencier"]),
                     fac=x["Prix facturé"], dfac=jour(x["Date de la facture"]),
                     ecart=(x["Prix facturé"] - x["Prix cadencier"]) / x["Prix cadencier"],
                     promo=x["Promo"] == "oui", n=x["Nombre d'achats"]) for x in table(wb["Cadencier contre facture"], 1)]
    factures.sort(key=lambda x: -abs(x["ecart"]))
    evolutions = [dict(f=x["Fournisseur"], lib=x["Libellé"], avant=x["Ancien prix"], davant=jour(x["Ancien cadencier"]),
                       apres=x["Nouveau prix"], dapres=jour(x["Nouveau cadencier"]),
                       ecart=(x["Nouveau prix"] - x["Ancien prix"]) / x["Ancien prix"])
                  for x in table(wb["Évolutions"], 1)]

    # Un fournisseur vu sur une facture mais absent du comparatif garde quand même sa couleur
    tous = fournisseurs + sorted({x["f"] for x in factures + evolutions} - set(fournisseurs))
    teintes, libres = [], iter(PALETTE * 3)
    for f in tous:
        teintes.append(COULEURS.get(f) or next(libres))
    clair = " ".join(f"--f{i}:{c[0]};" for i, c in enumerate(teintes))
    sombre = " ".join(f"--f{i}:{c[1]};" for i, c in enumerate(teintes))

    donnees = dict(client=client, provisoire=config.get("provisoire", True), fournisseurs=tous, produits=produits, factures=factures, evolutions=evolutions)
    texte = json.dumps(donnees, ensure_ascii=False).replace("</", "<\\/")
    page = (ICI / "page-modele.html").read_text(encoding="utf-8")
    page = (page.replace("__TITRE__", escape(f"Comparatif {client}"))
                .replace("__COULEURS_CLAIR__", clair).replace("__COULEURS_SOMBRE__", sombre)
                .replace("__DONNEES__", texte))
    sortie = dossier / "page-comparatif.html"
    sortie.write_text(page, encoding="utf-8")
    print(f"PAGE: {sortie}")


if __name__ == "__main__":
    main()

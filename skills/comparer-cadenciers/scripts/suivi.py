"""Mercuriale, tirée des tarifs et des achats relevés sur les factures.

La mercuriale met tous les fournisseurs et tous leurs tarifs dans un seul tableau, avec le dernier
prix payé sur facture et le verdict du comparatif.
familles.csv (facultatif, dans le dossier du client) : fournisseur, ref, famille.
"""
import csv
from collections import defaultdict
from pathlib import Path

from lecteurs import cle_ref

# Fonctions de tableur reconnues à la fois par Excel, LibreOffice et Google Sheets.
# Le classeur ne doit utiliser qu'elles, pour s'ouvrir chez un client qui n'a pas Excel.
COMPATIBLES = {"AND", "CEILING", "COUNT", "COUNTIF", "COUNTIFS", "IF", "IFERROR", "INDEX", "ISNUMBER", "ISTEXT",
               "MATCH", "MAX", "MAXIFS", "MIN", "MINIFS", "N", "OR", "ROUND", "SMALL", "SUMIF", "TODAY"}


def fonctions_incompatibles(wb):
    """Fonctions du classeur absentes de COMPATIBLES : elles risquent de ne pas marcher dans Google Sheets."""
    import re
    vues = set()
    for ws in wb:
        for ligne in ws.iter_rows():
            for c in ligne:
                if isinstance(c.value, str) and c.value.startswith("="):
                    sans_texte = re.sub(r'"[^"]*"', "", c.value)
                    vues |= {f.replace("_xlfn.", "") for f in re.findall(r"([A-Z_][A-Z0-9_.]*)\(", sans_texte)}
    return sorted(vues - COMPATIBLES)


def lire_csv(chemin):
    chemin = Path(chemin)
    return list(csv.DictReader(open(chemin, encoding="utf-8-sig"))) if chemin.exists() else []


def derniers_achats(achats):
    """Par référence achetée : dernier prix payé (hors promotion s'il existe), date du dernier achat, nombre d'achats."""
    groupes = defaultdict(list)
    for a in achats:
        groupes[(a["fournisseur"], cle_ref(a["ref"]))].append(a)
    out = {}
    for cle, lignes in groupes.items():
        lignes.sort(key=lambda a: a["date_livraison"])
        # Le dernier prix hors promotion, s'il existe, est celui qu'on paiera à la prochaine commande
        normal = [a for a in lignes if str(a.get("promo", "0")) in ("0", "")] or lignes
        out[cle] = dict(dernier_prix=float(normal[-1]["prix"]), date_dernier=lignes[-1]["date_livraison"], nb=len(lignes))
    return out


# ---------- Onglets du classeur ----------

def ajouter_onglets(wb, ici, recent, produits, outils, comparatif):
    """Ajoute la Mercuriale au classeur.
    outils : fonctions de mise en forme de comparer.py (entete, habiller, jour, PatternFill, FormulaRule, L).
    comparatif : rangs et colonnes de l'onglet Comparatif, pour reprendre son verdict par formule."""
    entete, habiller, jour = outils["entete"], outils["habiller"], outils["jour"]
    PatternFill, FormulaRule, L = outils["PatternFill"], outils["FormulaRule"], outils["L"]
    familles = {(f["fournisseur"], cle_ref(f["ref"])): f["famille"] for f in lire_csv(Path(ici) / "familles.csv")}
    appariement = {}
    for nom, p in produits.items():
        for a, _t in p["lignes"]:
            appariement[(a["fournisseur"], cle_ref(a["ref"]))] = (nom, p["famille"])
    achats = lire_csv(Path(ici) / "achats.csv")
    payes = derniers_achats(achats)

    def famille_produit(cle):
        nom, fam = appariement.get(cle, ("", ""))
        return familles.get(cle) or fam or "À classer", nom

    def verdict(r):
        if not comparatif:
            return ""
        p, d, cm = comparatif["premiere"], comparatif["dern"], comparatif["moins_cher"]
        return (f'=IF(B{r}="","",IFERROR(INDEX(Comparatif!${cm}${p}:${cm}${d},'
                f'MATCH(B{r},Comparatif!$A${p}:$A${d},0)),""))')

    # ---------- Mercuriale : tous les fournisseurs, tous leurs tarifs ----------
    wm = wb.create_sheet("Mercuriale", 1)
    tit = ["Famille", "Produit comparé", "Fournisseur", "Réf.", "Libellé", "Marque", "Conditionnement", "Prix HT",
           "Unité de prix", "Prix par kg, litre ou unité", "Unité", "Date du tarif", "Fin de validité", "Offre",
           "Dernier prix payé", "Dernier achat", "Achats relevés", "Moins cher (comparatif)"]
    entete(wm, 1, tit)
    lignes = []
    for cle, t in recent.items():
        fam, nom = famille_produit(cle)
        lignes.append((fam == "À classer", fam, nom or "~", t["fournisseur"], t["libelle"], cle, t))
    lignes.sort(key=lambda x: x[:5])
    for r, (_, fam, nom, four, _lib, cle, t) in enumerate(lignes, 2):
        c = payes.get(cle)
        valeurs = [fam, "" if nom == "~" else nom, four, t["ref"], t["libelle"], t.get("marque", ""), t["conditionnement"],
                   float(t["prix"]), t["unite_prix"].replace("€/", ""),
                   float(t["prix_unitaire"]) if t["prix_unitaire"] else None, t["unite"],
                   jour(t["date_tarif"]), jour(t["fin_validite"]),
                   f'=IF(AND(M{r}<>"",M{r}<TODAY()),"expirée","")',
                   c["dernier_prix"] if c else None, jour(c["date_dernier"]) if c else None, c["nb"] if c else None,
                   verdict(r)]
        for k, v in enumerate(valeurs, 1):
            wm.cell(r, k, v)
        for k, fmt in ((8, "0.000"), (10, "0.00"), (15, "0.000")):
            wm.cell(r, k).number_format = fmt
        for k in (12, 13, 16):
            wm.cell(r, k).number_format = "dd/mm/yyyy"
    dern_m = len(lignes) + 1
    habiller(wm, 2, dern_m, len(tit))
    for k, w in enumerate((20, 30, 19, 9, 42, 14, 16, 9, 8, 11, 7, 11, 11, 9, 10, 11, 8, 19), 1):
        wm.column_dimensions[L(k)].width = w
    wm.freeze_panes = "C2"
    if dern_m >= 2:
        wm.auto_filter.ref = f"A1:{L(len(tit))}{dern_m}"
        wm.conditional_formatting.add(f"N2:N{dern_m}", FormulaRule(formula=['$N2="expirée"'],
                                      fill=PatternFill("solid", bgColor="FEF3C7")))
    return len(recent)

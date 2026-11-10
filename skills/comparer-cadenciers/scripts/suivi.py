"""Mercuriale et Historique, tirés des tarifs et des factures (controle_factures.csv, historique_factures.csv).

La mercuriale met tous les fournisseurs et tous leurs tarifs dans un seul tableau, avec le dernier
prix payé sur facture et le verdict du comparatif.
familles.csv (facultatif, dans le dossier du client) : fournisseur, ref, famille.
"""
import csv
from pathlib import Path

from lecteurs import cle_ref

# Fonctions de tableur reconnues à la fois par Excel, LibreOffice et Google Sheets.
# Le classeur ne doit utiliser qu'elles, pour s'ouvrir chez un client qui n'a pas Excel.
COMPATIBLES = {"AND", "COUNT", "COUNTIF", "COUNTIFS", "IF", "IFERROR", "INDEX", "ISNUMBER", "MATCH", "MAXIFS", "MIN",
               "MINIFS", "N", "OR", "ROUND", "SMALL", "TODAY"}


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


def derniers_payes(controle):
    """Par référence : dernier prix payé ramené à l'unité de prix du tarif, date du dernier achat, nombre d'achats.
    Même règle que le Détail (controle_factures.csv) : le dernier prix hors promotion, sinon le dernier prix."""
    out = {}
    for x in controle:
        if x.get("prix_hors_promo"):
            prix, facteur = float(x["prix_hors_promo"]), float(x["facteur_hors_promo"] or 1)
        else:
            prix, facteur = float(x["prix_facture"]), float(x["facteur"] or 1)
        # 15,763 €/kg facturés deviennent 22,068 € le demi-pain de 1,4 kg du tarif
        out[(x["fournisseur"], cle_ref(x["ref"]))] = dict(paye=round(prix * facteur, 4), date=x["date_facture"],
                                                           nb=int(x["nb_achats"]))
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
    payes = derniers_payes(lire_csv(Path(ici) / "controle_factures.csv"))

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
                   c["paye"] if c else None, jour(c["date"]) if c else None, c["nb"] if c else None,
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


def ajouter_historique(wb, ici, tarifs, recent, produits, outils):
    """Ajoute l'onglet Historique : chaque prix connu d'une référence, cadencier ou facture, dans l'ordre des dates.
    Seules les références vues à au moins deux dates y figurent. Renvoie le nombre de références suivies."""
    entete, habiller, jour = outils["entete"], outils["habiller"], outils["jour"]
    Font, PatternFill, FormulaRule, L = outils["Font"], outils["PatternFill"], outils["FormulaRule"], outils["L"]
    nom_produit = {(a["fournisseur"], cle_ref(a["ref"])): nom for nom, p in produits.items() for a, _t in p["lignes"]}
    points = {}

    def ajouter(cle, point):
        # Deux documents du même jour au même prix ne font qu'un relevé
        date_, origine, _doc, prix, up = point[:5]
        points.setdefault(cle, {}).setdefault((date_, origine, prix, up), point)

    for t in tarifs:
        pu = float(t["prix_unitaire"]) if t["prix_unitaire"] else None
        ajouter((t["fournisseur"], cle_ref(t["ref"])),
                (t["date_tarif"], "Cadencier", t["document"], float(t["prix"]), t["unite_prix"].replace("€/", ""), pu,
                 t["unite"], ""))
    for x in lire_csv(Path(ici) / "historique_factures.csv"):
        cle = (x["fournisseur"], cle_ref(x["ref"]))
        t = recent.get(cle)
        if not t:
            continue
        prix = float(x["prix_ramene"])
        pu = round(prix / float(t["qte"]), 4) if t["qte"] else None
        ajouter(cle, (x["date"], "Facture", x["facture"], prix, t["unite_prix"].replace("€/", ""),
                      pu, t["unite"], "oui" if x["promo"] == "1" else ""))
    suivis = sorted(((cle, sorted(pts.values(), key=lambda p: p[:4])) for cle, pts in points.items()
                     if len({p[0] for p in pts.values()}) >= 2 and cle in recent),
                    key=lambda kv: (kv[0][0], recent[kv[0]]["libelle"]))

    wh = wb.create_sheet("Historique")
    tit = ["Fournisseur", "Réf.", "Libellé", "Produit comparé", "Date", "Origine", "Document", "Prix HT", "Unité de prix",
           "Prix par kg, litre ou unité", "Unité", "Promo", "Depuis le premier prix"]
    entete(wh, 1, tit)
    r = 2
    gris = PatternFill("solid", fgColor="F1F3F5")
    for k, (cle, pts) in enumerate(suivis):
        t = recent[cle]
        premier = None
        for date_, origine, doc, prix, up, pu, unite, promo in pts:
            valeurs = [cle[0], t["ref"], t["libelle"], nom_produit.get(cle, ""), jour(date_), origine, doc, prix, up, pu,
                       unite, promo]
            for c, v in enumerate(valeurs, 1):
                wh.cell(r, c, v)
            if premier is None and pu is not None:
                premier = r
            # L'évolution se lit au kg, au litre ou à l'unité : un pack et une bouteille ne se comparent pas au prix brut
            wh.cell(r, 13, f'=IF(OR(J{r}="",J${premier}=""),"",J{r}/J${premier}-1)' if premier else None)
            wh.cell(r, 5).number_format = "dd/mm/yyyy"
            wh.cell(r, 8).number_format = "0.000"
            wh.cell(r, 10).number_format = "0.00"
            wh.cell(r, 13).number_format = "+0.0%;-0.0%;0.0%"
            if k % 2:
                for c in range(1, len(tit) + 1):
                    wh.cell(r, c).fill = gris
            r += 1
    dern = r - 1
    habiller(wh, 2, dern, len(tit))
    for k, w in enumerate((20, 9, 42, 28, 11, 10, 28, 10, 9, 11, 7, 7, 12), 1):
        wh.column_dimensions[L(k)].width = w
    wh.freeze_panes = "D2"
    if dern >= 2:
        wh.auto_filter.ref = f"A1:{L(len(tit))}{dern}"
        wh.conditional_formatting.add(f"M2:M{dern}", FormulaRule(formula=["AND(ISNUMBER($M2),$M2>0)"],
                                      font=Font(bold=True, color="B3261E")))
        wh.conditional_formatting.add(f"M2:M{dern}", FormulaRule(formula=["AND(ISNUMBER($M2),$M2<0)"],
                                      font=Font(bold=True, color="1E7A3C")))
    return len(suivis)

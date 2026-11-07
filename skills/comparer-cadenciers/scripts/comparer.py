#!/usr/bin/env python3
"""Construit le classeur comparatif à partir de tarifs.csv, appariements.csv et controle_factures.csv.

Usage : comparer.py <dossier du client>
Nombre de fournisseurs libre : il suffit d'en ajouter dans appariements.csv.
Tous les calculs restent des formules Excel, pour pouvoir corriger un prix à la main.
"""
import csv
import json
import sys
from collections import OrderedDict
from datetime import date
from pathlib import Path

sys.dont_write_bytecode = True

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

ICI = None
COULEURS = {"Sysco": "1D5FA6", "Pro à Pro": "1E7A3C", "Episaveurs": "B45309",
            "Comptoir du Littoral": "6D28D9", "Sodiboissons": "BE123C"}
PALETTE = ["0F766E", "A16207", "9D174D", "1E40AF", "4D7C0F", "7E22CE", "C2410C", "155E75"]
POLICE = "Arial"
GRIS = "F1F3F5"
fin = Side(style="thin", color="D5D9DE")
BORD = Border(bottom=fin)


def lire(nom):
    return list(csv.DictReader(open(ICI / nom, encoding="utf-8-sig")))


def jour(txt):
    return date.fromisoformat(txt) if txt else None


def entete(ws, rang, titres, couleur="2B3440"):
    for i, t in enumerate(titres, 1):
        c = ws.cell(rang, i, t)
        c.font = Font(name=POLICE, bold=True, color="FFFFFF", size=10)
        c.fill = PatternFill("solid", fgColor=couleur)
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[rang].height = 44


def habiller(ws, premier, dernier, nb_col):
    for r in range(premier, dernier + 1):
        for c in range(1, nb_col + 1):
            cell = ws.cell(r, c)
            cell.font = Font(name=POLICE, size=10, bold=cell.font.bold, color=cell.font.color)
            cell.border = BORD
            cell.alignment = Alignment(vertical="center", wrap_text=cell.alignment.wrap_text)


def main():
    global ICI
    if len(sys.argv) != 2:
        sys.exit("Usage : comparer.py <dossier du client>")
    ICI = Path(sys.argv[1]).expanduser()
    config = json.load(open(ICI / "config.json", encoding="utf-8"))
    client = config["client"]
    SORTIE = ICI / "Comparatif-fournisseurs.xlsx"
    tarifs = lire("tarifs.csv")
    recent = {}
    for t in tarifs:
        cle = (t["fournisseur"], t["ref"].lstrip("0"))
        if cle not in recent or t["date_tarif"] > recent[cle]["date_tarif"]:
            recent[cle] = t
    factures = {}
    if (ICI / "controle_factures.csv").exists():
        factures = {(x["fournisseur"], x["ref"]): x for x in lire("controle_factures.csv")}

    produits = OrderedDict()
    erreurs = []
    for a in lire("appariements.csv"):
        cle = (a["fournisseur"], a["ref"].lstrip("0"))
        t = recent.get(cle)
        if not t:
            erreurs.append(f"Référence absente des cadenciers : {cle}")
            continue
        if not a["qte_forcee"] and "€/" + t["unite"] != a["unite"]:
            erreurs.append(f"Unité différente pour {a['produit']} {cle} : {t['unite']} au lieu de {a['unite']}")
        produits.setdefault(a["produit"], dict(famille=a["famille"], unite=a["unite"], type=a["type"],
                                               note=a["note"], lignes=[]))["lignes"].append((a, t))
    if erreurs:
        print("\n".join(erreurs))
        sys.exit(1)

    fournisseurs = list(OrderedDict.fromkeys(l[0]["fournisseur"] for p in produits.values() for l in p["lignes"]))
    fournisseurs.sort(key=lambda f: list(COULEURS).index(f) if f in COULEURS else len(COULEURS))
    for i, f in enumerate(x for x in fournisseurs if x not in COULEURS):
        COULEURS[f] = PALETTE[i % len(PALETTE)]
    wb = Workbook()

    # ---------- Détail : une ligne par produit et par fournisseur ----------
    wd = wb.active
    wd.title = "Détail"
    tit = ["Produit", "Unité", "Fournisseur", "Réf.", "Libellé du cadencier", "Prix cadencier HT", "Unité de prix",
           "Date du cadencier", "Dernière facture hors promo", "Date de la facture", "Prix retenu HT",
           "Quantité par unité de prix", "Prix par unité de comparaison", "Date du prix retenu", "Origine du prix",
           "Fiabilité", "Note de lecture", "Fin de validité"]
    entete(wd, 1, tit)
    r = 2
    cellule = {}
    for nom, p in produits.items():
        for a, t in p["lignes"]:
            f = factures.get((a["fournisseur"], a["ref"].lstrip("0")))
            qte = float(a["qte_forcee"] or t["qte"])
            wd.cell(r, 1, nom)
            wd.cell(r, 2, p["unite"])
            wd.cell(r, 3, a["fournisseur"])
            wd.cell(r, 4, t["ref"])
            wd.cell(r, 5, t["libelle"])
            wd.cell(r, 6, float(t["prix"])).font = Font(color="0000FF")
            wd.cell(r, 7, t["unite_prix"].replace("€/", ""))
            wd.cell(r, 8, jour(t["date_tarif"])).font = Font(color="0000FF")
            if f and f["prix_hors_promo"]:
                wd.cell(r, 9, float(f["prix_hors_promo"])).font = Font(color="0000FF")
                wd.cell(r, 10, jour(f["date_hors_promo"])).font = Font(color="0000FF")
            wd.cell(r, 11, f'=IF(AND(I{r}<>"",J{r}>H{r}),I{r},F{r})')
            wd.cell(r, 12, qte).font = Font(color="0000FF")
            wd.cell(r, 13, f"=K{r}/L{r}")
            wd.cell(r, 14, f'=IF(AND(I{r}<>"",J{r}>H{r}),J{r},H{r})')
            wd.cell(r, 15, f'=IF(AND(I{r}<>"",J{r}>H{r}),"facture","cadencier")')
            # Une facture confirme un prix, pas une quantité : « à vérifier » ne se lève qu'avec une
            # quantité validée, « si facture » dès qu'une facture existe pour cette référence
            doute = not a["qte_forcee"] and (t["confiance"] == "à vérifier" or (t["confiance"] == "si facture" and not f))
            wd.cell(r, 16, "à vérifier" if doute else "vu sur facture" if f else "lu sur le document")
            wd.cell(r, 17, t["note"] if doute else "")
            wd.cell(r, 18, jour(t["fin_validite"]))
            cellule[(nom, a["fournisseur"])] = r
            r += 1
    dern_d = r - 1
    habiller(wd, 2, dern_d, len(tit))
    for rr in range(2, dern_d + 1):
        for c, fmt in ((6, "0.000"), (9, "0.000"), (11, "0.000"), (13, "0.00"), (12, "0.###")):
            wd.cell(rr, c).number_format = fmt
        for c in (8, 10, 14, 18):
            wd.cell(rr, c).number_format = "dd/mm/yyyy"
    for i, w in enumerate((36, 11, 19, 9, 44, 10, 8, 11, 11, 11, 10, 10, 12, 11, 10, 16, 40, 11), 1):
        wd.column_dimensions[L(i)].width = w
    wd.freeze_panes = "B2"
    wd.auto_filter.ref = f"A1:{L(len(tit))}{dern_d}"
    wd.conditional_formatting.add(f"P2:P{dern_d}", FormulaRule(formula=['$P2="à vérifier"'],
                                  fill=PatternFill("solid", bgColor="FEF3C7")))

    # ---------- Comparatif ----------
    wc = wb.create_sheet("Comparatif", 0)
    n = len(fournisseurs)
    c0 = 4  # première colonne de prix
    wc["A1"] = f"Comparatif fournisseurs à partir des cadenciers · {client}"
    wc["A1"].font = Font(name=POLICE, bold=True, size=14)
    wc["A2"] = ("Prix HT ramenés au kg, au litre ou à l'unité indiquée. Quand une facture est plus récente que le "
                "cadencier, c'est son prix hors promotion qui est retenu.")
    wc["A2"].font = Font(name=POLICE, size=10, color="55606E")

    if config.get("provisoire", True):
        wc["A3"] = "Résultat provisoire : les correspondances entre produits restent à valider."
        wc["A3"].font = Font(name=POLICE, size=10, bold=True, color="B3261E")

    # Bilan
    wc["A4"] = "Bilan (produits comparables)"
    wc["A4"].font = Font(name=POLICE, bold=True, size=11)
    rang_titres = 4 + n + 3
    premiere = rang_titres + 1
    titres = ["Produit", "Unité", ""] + fournisseurs + ["Moins cher", "Écart avec le 2e (par unité)", "Écart %",
                                                       "Mois entre les prix", "Prix à vérifier", "Type", "Remarque"]
    # colonnes : A produit, B unité, C famille, D... prix
    entete(wc, rang_titres, titres)
    wc.cell(rang_titres, 3, "Famille")
    for i, f in enumerate(fournisseurs):
        wc.cell(rang_titres, c0 + i).fill = PatternFill("solid", fgColor=COULEURS[f])

    a, b = L(c0), L(c0 + n - 1)
    cm, ce, cp, cmo, cd, ct, cn = (c0 + n + i for i in range(7))
    r = premiere
    familles = list(OrderedDict.fromkeys(p["famille"] for p in produits.values()))
    ordre = sorted(produits.items(), key=lambda kv: (familles.index(kv[1]["famille"]), kv[1]["type"] == "different"))
    for nom, p in ordre:
        wc.cell(r, 1, nom)
        wc.cell(r, 2, p["unite"])
        wc.cell(r, 3, p["famille"])
        for i, f in enumerate(fournisseurs):
            if (nom, f) in cellule:
                c = wc.cell(r, c0 + i, f"='Détail'!M{cellule[(nom, f)]}")
                c.font = Font(color="008000")
                c.number_format = "0.00"
        wc.cell(r, ct, {"identique": "identique", "equivalent": "équivalent", "different": "non comparable"}[p["type"]])
        rng = f"{a}{r}:{b}{r}"
        typ = f"{L(ct)}{r}"
        wc.cell(r, cm, f'=IF({typ}="non comparable","non comparable",IF(COUNT({rng})<2,"seul",'
                       f'IF(MIN({rng})=SMALL({rng},2),"égalité",INDEX(${a}${rang_titres}:${b}${rang_titres},MATCH(MIN({rng}),{rng},0)))))')
        wc.cell(r, ce, f'=IF(OR({typ}="non comparable",COUNT({rng})<2),"",SMALL({rng},2)-MIN({rng}))').number_format = "0.00"
        wc.cell(r, cp, f'=IF({L(ce)}{r}="","",{L(ce)}{r}/MIN({rng}))').number_format = "0%"
        wc.cell(r, cmo, f"=ROUND((_xlfn.MAXIFS('Détail'!$N$2:$N${dern_d},'Détail'!$A$2:$A${dern_d},$A{r})"
                        f"-_xlfn.MINIFS('Détail'!$N$2:$N${dern_d},'Détail'!$A$2:$A${dern_d},$A{r}))/30.4,0)").number_format = "0"
        wc.cell(r, cd, f"=COUNTIFS('Détail'!$A$2:$A${dern_d},$A{r},'Détail'!$P$2:$P${dern_d},\"à vérifier\")").number_format = "0"
        wc.cell(r, cn, p["note"])
        r += 1
    dern = r - 1
    habiller(wc, premiere, dern, cn)
    for rr in range(premiere, dern + 1):
        wc.cell(rr, cm).font = Font(name=POLICE, size=10, bold=True)
        wc.cell(rr, cn).alignment = Alignment(wrap_text=True, vertical="center")

    # Bilan : une ligne par fournisseur
    entetes_bilan = ["Fournisseur", "Produits où il est le moins cher"]
    for i, t in enumerate(entetes_bilan, 1):
        c = wc.cell(5, i, t)
        c.font = Font(name=POLICE, bold=True, size=10)
    for i, f in enumerate(fournisseurs):
        wc.cell(6 + i, 1, f).font = Font(name=POLICE, size=10, bold=True, color=COULEURS[f])
        c = wc.cell(6 + i, 2, f'=COUNTIF(${L(cm)}${premiere}:${L(cm)}${dern},A{6 + i})')
        c.font = Font(name=POLICE, size=10)

    # Le prix gagnant ressort, les prix écartés restent sobres
    for i, f in enumerate(fournisseurs):
        col = L(c0 + i)
        wc.conditional_formatting.add(
            f"{col}{premiere}:{col}{dern}",
            FormulaRule(formula=[f'AND(ISNUMBER({col}{premiere}),${L(cm)}{premiere}=${col}${rang_titres})'],
                        fill=PatternFill("solid", bgColor="D1FAE5"), font=Font(bold=True, color="065F46")))
    wc.conditional_formatting.add(f"{L(cmo)}{premiere}:{L(cmo)}{dern}", FormulaRule(
        formula=[f"AND(ISNUMBER({L(cmo)}{premiere}),{L(cmo)}{premiere}>3)"], fill=PatternFill("solid", bgColor="FEF3C7")))
    wc.conditional_formatting.add(f"{L(cp)}{premiere}:{L(cp)}{dern}", FormulaRule(
        formula=[f"AND(ISNUMBER({L(cp)}{premiere}),{L(cp)}{premiere}>0.5)"], fill=PatternFill("solid", bgColor="FED7AA")))
    wc.conditional_formatting.add(f"{L(cd)}{premiere}:{L(cd)}{dern}", FormulaRule(
        formula=[f"{L(cd)}{premiere}>0"], fill=PatternFill("solid", bgColor="FEF3C7")))
    wc.conditional_formatting.add(f"A{premiere}:{L(cn)}{dern}", FormulaRule(
        formula=[f'${L(ct)}{premiere}="non comparable"'], font=Font(color="8A94A3")))

    largeurs = {1: 38, 2: 12, 3: 20}
    for i in range(n):
        largeurs[c0 + i] = 12
    largeurs.update({cm: 19, ce: 11, cp: 8, cmo: 8, cd: 8, ct: 14, cn: 46})
    for c, w in largeurs.items():
        wc.column_dimensions[L(c)].width = w
    wc.freeze_panes = wc.cell(premiere, 2)
    wc.auto_filter.ref = f"A{rang_titres}:{L(cn)}{dern}"

    # ---------- Cadencier contre facture ----------
    wf = wb.create_sheet("Cadencier contre facture")
    titres = ["Fournisseur", "Réf.", "Libellé", "Date du cadencier", "Prix cadencier", "Date de la facture",
              "Prix facturé", "Écart %", "Promo", "Nombre d'achats"]
    entete(wf, 1, titres)
    controle = sorted(lire("controle_factures.csv") if factures else [], key=lambda x: (x["fournisseur"], -abs(float(x["ecart_pct"]))))
    for i, x in enumerate(controle, 2):
        wf.cell(i, 1, x["fournisseur"])
        wf.cell(i, 2, x["ref"])
        wf.cell(i, 3, x["libelle"])
        wf.cell(i, 4, jour(x["date_tarif"])).number_format = "dd/mm/yyyy"
        wf.cell(i, 5, float(x["prix_cadencier"])).number_format = "0.000"
        wf.cell(i, 6, jour(x["date_facture"])).number_format = "dd/mm/yyyy"
        wf.cell(i, 7, float(x["prix_facture"])).number_format = "0.000"
        wf.cell(i, 8, f"=(G{i}-E{i})/E{i}").number_format = "+0.0%;-0.0%;0.0%"
        wf.cell(i, 9, "oui" if x["promo"] == "1" else "")
        wf.cell(i, 10, int(x["nb_achats"]))
    habiller(wf, 2, len(controle) + 1, len(titres))
    for i, w in enumerate((21, 9, 44, 11, 10, 11, 10, 9, 7, 9), 1):
        wf.column_dimensions[L(i)].width = w
    wf.freeze_panes = "A2"
    wf.auto_filter.ref = f"A1:J{len(controle) + 1}"

    # ---------- Évolutions entre deux cadenciers d'un même fournisseur ----------
    we = wb.create_sheet("Évolutions")
    titres = ["Fournisseur", "Réf.", "Libellé", "Ancien cadencier", "Ancien prix", "Nouveau cadencier", "Nouveau prix", "Évolution"]
    entete(we, 1, titres)
    groupes = {}
    for t in tarifs:
        groupes.setdefault((t["fournisseur"], t["ref"]), []).append(t)
    i = 2
    for (four, ref), g in sorted(groupes.items()):
        if len(g) < 2:
            continue
        g.sort(key=lambda x: x["date_tarif"])
        if g[0]["prix"] == g[-1]["prix"]:
            continue
        we.cell(i, 1, four)
        we.cell(i, 2, ref)
        we.cell(i, 3, g[-1]["libelle"])
        we.cell(i, 4, jour(g[0]["date_tarif"])).number_format = "dd/mm/yyyy"
        we.cell(i, 5, float(g[0]["prix"])).number_format = "0.000"
        we.cell(i, 6, jour(g[-1]["date_tarif"])).number_format = "dd/mm/yyyy"
        we.cell(i, 7, float(g[-1]["prix"])).number_format = "0.000"
        we.cell(i, 8, f"=(G{i}-E{i})/E{i}").number_format = "+0.0%;-0.0%;0.0%"
        i += 1
    habiller(we, 2, i - 1, len(titres))
    for k, w in enumerate((21, 9, 44, 12, 10, 12, 10, 10), 1):
        we.column_dimensions[L(k)].width = w
    we.freeze_panes = "A2"

    # ---------- Références sans équivalent ----------
    ws = wb.create_sheet("Sans équivalent")
    titres = ["Fournisseur", "Réf.", "Libellé", "Prix HT", "Unité de prix", "Prix par kg ou litre", "Unité", "Date du tarif"]
    entete(ws, 1, titres)
    pris = {(l[0]["fournisseur"], l[0]["ref"].lstrip("0")) for p in produits.values() for l in p["lignes"]}
    seuls = sorted((t for k, t in recent.items() if k not in pris), key=lambda t: (t["fournisseur"], t["libelle"]))
    for i, t in enumerate(seuls, 2):
        ws.cell(i, 1, t["fournisseur"])
        ws.cell(i, 2, t["ref"])
        ws.cell(i, 3, t["libelle"])
        ws.cell(i, 4, float(t["prix"])).number_format = "0.000"
        ws.cell(i, 5, t["unite_prix"].replace("€/", ""))
        if t["prix_unitaire"]:
            ws.cell(i, 6, float(t["prix_unitaire"])).number_format = "0.00"
        ws.cell(i, 7, t["unite"])
        ws.cell(i, 8, jour(t["date_tarif"])).number_format = "dd/mm/yyyy"
    habiller(ws, 2, len(seuls) + 1, len(titres))
    for k, w in enumerate((21, 10, 46, 10, 9, 11, 8, 11), 1):
        ws.column_dimensions[L(k)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(len(seuls) + 1, 2)}"

    # ---------- Lisez-moi ----------
    wl = wb.create_sheet("Lisez-moi", 0)
    wl.column_dimensions["A"].width = 26
    wl.column_dimensions["B"].width = 92
    wl["A1"] = f"Comparatif à partir de cadenciers · {client}"
    wl["A1"].font = Font(name=POLICE, bold=True, size=14)
    lignes = [
        ("But", "Comparer les prix des fournisseurs à partir de leurs cadenciers" + (", contrôlés sur les factures." if factures else ". Aucune facture n'a servi au contrôle.")),
        ("Comparatif", "Un produit par ligne, un fournisseur par colonne. Le prix le plus bas ressort en vert."),
        ("Détail", "D'où vient chaque prix : libellé, date, quantité, facture. Les prix en bleu se corrigent à la main, le reste suit."),
        ("Cadencier contre facture", "Prix du cadencier face au dernier prix payé, pour chaque référence retrouvée sur une facture."),
        ("Évolutions", "Prix qui ont changé entre deux cadenciers d'un même fournisseur."),
        ("Sans équivalent", "Références lues mais comparées à rien, faute de produit équivalent chez un autre fournisseur."),
        ("Fond jaune", "« Mois entre les prix » supérieur à 3 : les prix comparés n'ont pas le même âge, le résultat est indicatif. « Prix à vérifier » : unité non confirmée par une facture."),
        ("Fond orange", "Écart supérieur à 50 % : souvent une erreur d'unité ou deux produits différents. À contrôler avant de s'y fier."),
        ("Prix retenu", "Le prix du cadencier, ou celui de la dernière facture hors promotion si elle est plus récente."),
        ("Non comparable", "Produit affiché pour mémoire, jamais compté dans le bilan."),
        ("Limites", "Les correspondances entre produits ont été proposées par Claude et restent à valider. Quand un document ne porte pas de date, c'est celle du fichier qui est utilisée. Ce sont des prix HT, hors franco, remises de fin d'année et frais de livraison."),
    ]
    for i, (k, v) in enumerate(lignes, 3):
        wl.cell(i, 1, k).font = Font(name=POLICE, bold=True, size=10)
        c = wl.cell(i, 2, v)
        c.font = Font(name=POLICE, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        wl.cell(i, 1).alignment = Alignment(vertical="top")
    wb.active = 1
    wb.save(SORTIE)
    print(f"{len(produits)} produits, {n} fournisseurs, {dern_d - 1} lignes de détail -> {SORTIE.name}")


if __name__ == "__main__":
    main()

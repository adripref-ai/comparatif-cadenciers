#!/usr/bin/env python3
"""Construit le classeur comparatif à partir de tarifs.csv, appariements.csv et controle_factures.csv.

Usage : comparer.py <dossier du client>
Nombre de fournisseurs libre : il suffit d'en ajouter dans appariements.csv.
Tous les calculs restent des formules, pour pouvoir corriger un prix à la main. Elles marchent dans Excel,
LibreOffice et Google Sheets.
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
from openpyxl.worksheet.datavalidation import DataValidation

from lecteurs import SEUIL_UNITE, cle_ref, nombre, unite_confirmee
from suivi import ajouter_onglets, fonctions_incompatibles
from validations import CHOIX, LIBELLE_TYPE, ONGLET, etat, lire_decisions

ICI = None
COULEURS = {"Sysco": "1D5FA6", "Pro à Pro": "1E7A3C", "Episaveurs": "B45309",
            "Comptoir du Littoral": "6D28D9", "Sodiboissons": "BE123C"}
PALETTE = ["0F766E", "A16207", "9D174D", "1E40AF", "4D7C0F", "7E22CE", "C2410C", "155E75"]
TYPES = {"identique": "identique", "equivalent": "equivalent", "équivalent": "equivalent",
         "different": "different", "différent": "different"}
POLICE = "Arial"
GRIS = "F1F3F5"
fin = Side(style="thin", color="D5D9DE")
BORD = Border(bottom=fin)


def lire(nom):
    return list(csv.DictReader(open(ICI / nom, encoding="utf-8-sig")))


# Clés de config.json qui changent le classeur. Les autres (artefact, google_sheets, tableur) n'y sont pas.
CONFIG_UTILE = ("client", "factures", "provisoire")

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
        cle = (t["fournisseur"], cle_ref(t["ref"]))
        if cle not in recent or t["date_tarif"] > recent[cle]["date_tarif"]:
            recent[cle] = t
    factures = {}
    if (ICI / "controle_factures.csv").exists():
        with open(ICI / "controle_factures.csv", encoding="utf-8-sig") as f:
            colonnes = next(csv.reader(f), [])
        # Une version précédente ne ramenait pas la facture à l'unité du cadencier : le classeur se contredirait
        if "facteur_hors_promo" not in colonnes:
            sys.exit("controle_factures.csv vient d'une version précédente du plugin : "
                     "relancer controler_factures.py, puis comparer.py.")
        factures = {(x["fournisseur"], cle_ref(x["ref"])): x for x in lire("controle_factures.csv")}

    produits = OrderedDict()
    erreurs = []
    for a in lire("appariements.csv"):
        cle = (a["fournisseur"], cle_ref(a["ref"]))
        t = recent.get(cle)
        if not t:
            erreurs.append(f"Référence absente des cadenciers : {cle}")
            continue
        typ = TYPES.get(a["type"].strip().lower())
        if not typ:
            erreurs.append(f"Type inconnu pour {a['produit']} : « {a['type']} ». Écrire identique, equivalent ou different.")
            continue
        try:
            qf = nombre(a["qte_forcee"]) if (a["qte_forcee"] or "").strip() else None
        except ValueError:
            erreurs.append(f"qte_forcee illisible pour {a['produit']} {cle} : « {a['qte_forcee']} »")
            continue
        decision, commentaire = etat(a)
        p = produits.setdefault(a["produit"], dict(famille=a["famille"], unite=a["unite"], type=typ,
                                                   note=a["note"], lignes=[], decision=decision,
                                                   commentaire=commentaire,
                                                   proposition=TYPES.get((a.get("proposition") or "").strip().lower(), typ)))
        if a["unite"] != p["unite"]:
            erreurs.append(f"Unités mélangées dans {a['produit']} : {a['unite']} pour {a['fournisseur']}, "
                           f"{p['unite']} sur la première ligne")
        elif not qf and "€/" + t["unite"] != p["unite"]:
            erreurs.append(f"Unité différente pour {a['produit']} {cle} : €/{t['unite']} au lieu de {p['unite']}. "
                           "Mettre une qte_forcee ou choisir une autre référence.")
        if any(x["fournisseur"] == a["fournisseur"] for x, _ in p["lignes"]):
            erreurs.append(f"Deux références de {a['fournisseur']} pour {a['produit']} : n'en garder qu'une.")
        p["lignes"].append((dict(a, qte_forcee=qf), t))
    if erreurs:
        print("\n".join(erreurs))
        sys.exit(1)
    # Des choix faits par le client dans le classeur seraient perdus en le reconstruisant
    if SORTIE.exists():
        dans_classeur = lire_decisions(SORTIE) or {}
        oublies = [nom for nom, (dec, com) in dans_classeur.items()
                   if nom in produits and (dec, com) != (produits[nom]["decision"], produits[nom]["commentaire"])]
        if oublies:
            sys.exit(f"Le classeur contient des choix du client pas encore reportés ({', '.join(oublies[:8])}"
                     f"{' ...' if len(oublies) > 8 else ''}) : lancer d'abord reprendre_validations.py, "
                     "ou supprimer le classeur pour abandonner ces choix.")
    if not produits:
        print("Aucun produit présent chez deux fournisseurs : rien à mettre en face. "
              "Le classeur ne garde que les références lues (onglet Sans équivalent).")

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
           "Fiabilité", "Note de lecture", "Fin de validité", "Offre", "Conditionnement"]
    entete(wd, 1, tit)
    r = 2
    cellule = {}
    rangs = {}  # lignes du Détail de chaque produit, qui se suivent
    prevu = {}  # prix par unité de comparaison attendu, pour ranger l'onglet Meilleurs prix
    for nom, p in produits.items():
        rangs[nom] = (r, r + len(p["lignes"]) - 1)
        for a, t in p["lignes"]:
            f = factures.get((a["fournisseur"], cle_ref(a["ref"])))
            # Une facture confirme l'unité si son prix hors promo reste proche du cadencier. Sinon elle
            # est sans doute dans une autre unité (un carton pour une bouteille) : son prix n'est pas retenu.
            # Le prix facturé est d'abord ramené à l'unité du cadencier (15,763 €/kg pour un demi-pain de 1,4 kg)
            hp = (float(f["prix_hors_promo"]) * float(f["facteur_hors_promo"] or 1)
                  if f and f["prix_hors_promo"] not in ("", None) else None)
            confirme = bool(f) and unite_confirmee(t["prix"], hp)
            facture_ecartee = bool(f) and hp is not None and not confirme
            # Comparer à l'unité vendue (qte_forcee = 1) quand seule la quantité manquait ne déduit rien
            quantite_seule = a["qte_forcee"] == 1 and t["note"].startswith("quantité absente")
            doute = t["confiance"] in ("à vérifier", "si facture") and not confirme and not quantite_seule
            qte = a["qte_forcee"] or float(t["qte"])
            wd.cell(r, 1, nom)
            wd.cell(r, 2, p["unite"])
            wd.cell(r, 3, a["fournisseur"])
            wd.cell(r, 4, t["ref"])
            wd.cell(r, 5, t["libelle"])
            wd.cell(r, 6, float(t["prix"])).font = Font(color="0000FF")
            wd.cell(r, 7, t["unite_prix"].replace("€/", ""))
            wd.cell(r, 8, jour(t["date_tarif"])).font = Font(color="0000FF")
            if confirme:
                wd.cell(r, 9, round(hp, 4)).font = Font(color="0000FF")
                wd.cell(r, 10, jour(f["date_hors_promo"])).font = Font(color="0000FF")
            wd.cell(r, 11, f'=IF(AND(I{r}<>"",J{r}>H{r}),I{r},F{r})')
            wd.cell(r, 12, qte).font = Font(color="0000FF")
            # Arrondi au dix-millième : 9,90 € les 3 kg et 3,30 €/kg sont bien à égalité
            wd.cell(r, 13, f"=ROUND(K{r}/L{r},4)")
            wd.cell(r, 14, f'=IF(AND(I{r}<>"",J{r}>H{r}),J{r},H{r})')
            wd.cell(r, 15, f'=IF(AND(I{r}<>"",J{r}>H{r}),"facture","cadencier")')
            wd.cell(r, 16, "à vérifier" if doute else "vu sur facture" if confirme else "lu sur le document")
            notes = [t["note"]] if doute and t["note"] else []
            if confirme and float(f["facteur_hors_promo"] or 1) != 1:
                notes.append(f"facturé {float(f['prix_hors_promo']):.3f} €/{f['unite_hors_promo']}, ".replace(".", ",")
                             + f"soit {hp:.3f} €".replace(".", ",") + f"/{t['unite_prix'].replace('€/', '')}")
            if facture_ecartee:
                notes.append(f"facturé {float(f['prix_hors_promo']):.3f} le {jour(f['date_hors_promo']):%d/%m/%Y}, "
                             "trop loin du cadencier : autre unité ? prix non retenu")
            wd.cell(r, 17, " ; ".join(notes))
            wd.cell(r, 18, jour(t["fin_validite"]))
            # Une offre est valable jusqu'à son dernier jour inclus
            wd.cell(r, 19, f'=IF(AND(R{r}<>"",O{r}="cadencier",R{r}<TODAY()),"expirée","")')
            wd.cell(r, 20, t.get("conditionnement", ""))
            retenu = hp if confirme and f["date_hors_promo"] > t["date_tarif"] else float(t["prix"])
            prevu[(nom, a["fournisseur"])] = round(retenu / qte, 4)
            cellule[(nom, a["fournisseur"])] = r
            r += 1
    dern_d = r - 1
    habiller(wd, 2, dern_d, len(tit))
    for rr in range(2, dern_d + 1):
        for c, fmt in ((6, "0.000"), (9, "0.000"), (11, "0.000"), (13, "0.00"), (12, "0.###")):
            wd.cell(rr, c).number_format = fmt
        for c in (8, 10, 14, 18):
            wd.cell(rr, c).number_format = "dd/mm/yyyy"
    for i, w in enumerate((36, 11, 19, 9, 44, 10, 8, 11, 11, 11, 10, 10, 12, 11, 10, 16, 40, 11, 9, 28), 1):
        wd.column_dimensions[L(i)].width = w
    wd.freeze_panes = "B2"
    if dern_d >= 2:
        wd.auto_filter.ref = f"A1:{L(len(tit))}{dern_d}"
        wd.conditional_formatting.add(f"P2:P{dern_d}", FormulaRule(formula=['$P2="à vérifier"'],
                                      fill=PatternFill("solid", bgColor="FEF3C7")))
        wd.conditional_formatting.add(f"S2:S{dern_d}", FormulaRule(formula=['$S2="expirée"'],
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
        wc["A3"] = "Résultat provisoire : les correspondances entre produits restent à valider (onglet À valider)."
        wc["A3"].font = Font(name=POLICE, size=10, bold=True, color="B3261E")

    # Bilan
    wc["A4"] = "Bilan (produits comparables)"
    wc["A4"].font = Font(name=POLICE, bold=True, size=11)
    rang_titres = 4 + n + 3
    premiere = rang_titres + 1
    titres = ["Produit", "Unité", ""] + fournisseurs + ["Moins cher", "Écart avec le 2e (par unité)", "Écart %",
                                                       "Mois entre les prix", "Prix à vérifier", "Offres expirées",
                                                       "Résultat solide", "Type", "Remarque"]
    # colonnes : A produit, B unité, C famille, D... prix
    entete(wc, rang_titres, titres)
    wc.cell(rang_titres, 3, "Famille")
    for i, f in enumerate(fournisseurs):
        wc.cell(rang_titres, c0 + i).fill = PatternFill("solid", fgColor=COULEURS[f])

    a, b = L(c0), L(c0 + n - 1)
    cm, ce, cp, cmo, cd, cx, cs, ct, cn = (c0 + n + i for i in range(9))
    r = premiere
    familles = list(OrderedDict.fromkeys(p["famille"] for p in produits.values()))
    ordre = sorted(produits.items(), key=lambda kv: (familles.index(kv[1]["famille"]), kv[1]["type"] == "different"))
    # Onglet À valider : d'abord ce qui reste à trancher, puis ce qui est tranché, puis les produits identiques
    a_valider = sorted(ordre, key=lambda kv: (2 if kv[1]["type"] == "identique" and not kv[1]["decision"]
                                              else 1 if kv[1]["decision"] else 0))
    RV = 6  # ligne des titres de l'onglet À valider
    rang_val = {nom: RV + 1 + i for i, (nom, _) in enumerate(a_valider)}
    col_dec = L(3 + n + 3)
    rang_comp = {}
    for nom, p in ordre:
        rang_comp[nom] = r
        wc.cell(r, 1, nom)
        wc.cell(r, 2, p["unite"])
        wc.cell(r, 3, p["famille"])
        for i, f in enumerate(fournisseurs):
            if (nom, f) in cellule:
                c = wc.cell(r, c0 + i, f"='Détail'!M{cellule[(nom, f)]}")
                c.font = Font(color="008000")
                c.number_format = "0.00"
        # Le choix du client dans l'onglet À valider l'emporte tout de suite : le bilan se recalcule.
        # Case vide : la proposition de Claude, même si un choix précédent a été reporté puis effacé.
        dec = f"'{ONGLET}'!${col_dec}${rang_val[nom]}"
        wc.cell(r, ct, f'=IF({dec}="{CHOIX["identique"]}","{LIBELLE_TYPE["identique"]}",'
                       f'IF({dec}="{CHOIX["equivalent"]}","{LIBELLE_TYPE["equivalent"]}",'
                       f'IF({dec}="{CHOIX["different"]}","{LIBELLE_TYPE["different"]}","{LIBELLE_TYPE[p["proposition"]]}")))')
        rng = f"{a}{r}:{b}{r}"
        typ = f"{L(ct)}{r}"
        wc.cell(r, cm, f'=IF({typ}="non comparable","non comparable",IF(COUNT({rng})<2,"seul",'
                       f'IF(MIN({rng})=SMALL({rng},2),"égalité",INDEX(${a}${rang_titres}:${b}${rang_titres},MATCH(MIN({rng}),{rng},0)))))')
        wc.cell(r, ce, f'=IF(OR({typ}="non comparable",COUNT({rng})<2),"",SMALL({rng},2)-MIN({rng}))').number_format = "0.00"
        wc.cell(r, cp, f'=IF({L(ce)}{r}="","",{L(ce)}{r}/MIN({rng}))').number_format = "0%"
        # Un dixième de mois près : 106 jours font 3,5 mois, donc plus de 3
        wc.cell(r, cmo, f"=ROUND((_xlfn.MAXIFS('Détail'!$N$2:$N${dern_d},'Détail'!$A$2:$A${dern_d},$A{r})"
                        f"-_xlfn.MINIFS('Détail'!$N$2:$N${dern_d},'Détail'!$A$2:$A${dern_d},$A{r}))/30.4,1)").number_format = "0.0"
        wc.cell(r, cd, f"=COUNTIFS('Détail'!$A$2:$A${dern_d},$A{r},'Détail'!$P$2:$P${dern_d},\"à vérifier\")").number_format = "0"
        wc.cell(r, cx, f"=COUNTIFS('Détail'!$A$2:$A${dern_d},$A{r},'Détail'!$S$2:$S${dern_d},\"expirée\")").number_format = "0"
        g = f"{L(cm)}{r}"
        wc.cell(r, cs, f'=IF(AND({g}<>"non comparable",{g}<>"seul",{L(cmo)}{r}<=3,{L(cd)}{r}=0,{L(cx)}{r}=0),"oui","")')
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
    if dern >= premiere:
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
        for c in (cd, cx):
            wc.conditional_formatting.add(f"{L(c)}{premiere}:{L(c)}{dern}", FormulaRule(
                formula=[f"{L(c)}{premiere}>0"], fill=PatternFill("solid", bgColor="FEF3C7")))
        wc.conditional_formatting.add(f"{L(cs)}{premiere}:{L(cs)}{dern}", FormulaRule(
            formula=[f'{L(cs)}{premiere}="oui"'], fill=PatternFill("solid", bgColor="D1FAE5")))
        wc.conditional_formatting.add(f"A{premiere}:{L(cn)}{dern}", FormulaRule(
            formula=[f'${L(ct)}{premiere}="non comparable"'], font=Font(color="8A94A3")))
        wc.auto_filter.ref = f"A{rang_titres}:{L(cn)}{dern}"

    largeurs = {1: 38, 2: 12, 3: 20}
    for i in range(n):
        largeurs[c0 + i] = 12
    largeurs.update({cm: 19, ce: 11, cp: 8, cmo: 8, cd: 8, cx: 8, cs: 9, ct: 14, cn: 46})
    for c, w in largeurs.items():
        wc.column_dimensions[L(c)].width = w
    wc.freeze_panes = wc.cell(premiere, 2)

    # ---------- Meilleurs prix : le fournisseur le moins cher de chaque produit ----------
    # Tout se lit dans le Comparatif et le Détail : corriger un prix change aussi cet onglet.
    # Les lignes sont rangées par fournisseur d'après les prix du moment.
    wm = wb.create_sheet("Meilleurs prix", 1)
    wm["A1"] = f"Meilleurs prix · {client}"
    wm["A1"].font = Font(name=POLICE, bold=True, size=14)
    wm["A2"] = ("Pour chaque produit comparable, le fournisseur le moins cher, sa référence, son conditionnement et "
                "son prix. Prix HT, hors franco, remises et frais de livraison.")
    wm["A2"].font = Font(name=POLICE, size=10, color="55606E")
    if config.get("provisoire", True):
        wm["A3"] = "Résultat provisoire : les correspondances entre produits restent à valider (onglet À valider)."
        wm["A3"].font = Font(name=POLICE, size=10, bold=True, color="B3261E")
    titres = ["Fournisseur", "Famille", "Produit", "Réf.", "Libellé du fournisseur", "Conditionnement", "Prix HT",
              "Unité de prix", "Prix par unité de comparaison", "Unité", "Moins cher que le 2e de", "Résultat solide",
              "Remarque"]
    rm = 5
    entete(wm, rm, titres)
    # Tous les produits y figurent, pour suivre les choix du client sans reconstruire : ceux qui ne sont
    # pas comparables à la construction vont en bas, et leur ligne affiche « non comparable »
    comparables = [(nom, p) for nom, p in ordre if len(p["lignes"]) >= 2]

    def gagnant(nom):
        prix = [(prevu[(nom, f)], i) for i, f in enumerate(fournisseurs) if (nom, f) in prevu]
        return min(prix)[1]

    comparables.sort(key=lambda kv: (kv[1]["type"] == "different", gagnant(kv[0])))
    i = rm + 1
    for nom, p in comparables:
        rc = rang_comp[nom]
        d1, d2 = rangs[nom]
        prix = f"'Comparatif'!${a}${rc}:${b}${rc}"
        # Un produit que le client déclare non comparable sort aussitôt de la liste
        hors = f"'Comparatif'!${L(ct)}${rc}=\"{LIBELLE_TYPE['different']}\""
        wm.cell(i, 1, f"=IF({hors},\"{LIBELLE_TYPE['different']}\","
                      f"IFERROR(INDEX('Comparatif'!${a}${rang_titres}:${b}${rang_titres},MATCH(MIN({prix}),{prix},0)),\"\"))")
        wm.cell(i, 2, p["famille"])
        wm.cell(i, 3, nom)

        def detail(col):
            return f"=IFERROR(INDEX('Détail'!${col}${d1}:${col}${d2},MATCH($A{i},'Détail'!$C${d1}:$C${d2},0)),\"\")"

        wm.cell(i, 4, detail("D"))
        wm.cell(i, 5, detail("E"))
        wm.cell(i, 6, detail("T"))
        wm.cell(i, 7, detail("K")).number_format = "0.000"
        wm.cell(i, 8, detail("G"))
        wm.cell(i, 9, f"=IF({hors},\"\",MIN({prix}))").number_format = "0.00"
        wm.cell(i, 10, p["unite"])
        # À égalité, MATCH garde le premier fournisseur : la colonne le dit
        wm.cell(i, 11, f"=IF('Comparatif'!{L(cm)}{rc}=\"égalité\",\"égalité\",'Comparatif'!{L(cp)}{rc})").number_format = "0%"
        wm.cell(i, 12, f"='Comparatif'!{L(cs)}{rc}")
        wm.cell(i, 13, p["note"]).alignment = Alignment(wrap_text=True)
        i += 1
    dern_m = i - 1
    habiller(wm, rm + 1, dern_m, len(titres))
    for rr in range(rm + 1, dern_m + 1):
        wm.cell(rr, 1).font = Font(name=POLICE, size=10, bold=True)
        wm.cell(rr, 9).font = Font(name=POLICE, size=10, bold=True)
    if dern_m > rm:
        for f in fournisseurs:
            wm.conditional_formatting.add(f"A{rm + 1}:A{dern_m}", FormulaRule(
                formula=[f'$A{rm + 1}="{f}"'], font=Font(bold=True, color=COULEURS[f])))
        wm.conditional_formatting.add(f"A{rm + 1}:{L(len(titres))}{dern_m}", FormulaRule(
            formula=[f'$A{rm + 1}="{LIBELLE_TYPE["different"]}"'], font=Font(color="8A94A3")))
        wm.conditional_formatting.add(f"L{rm + 1}:L{dern_m}", FormulaRule(
            formula=[f'$L{rm + 1}="oui"'], fill=PatternFill("solid", bgColor="D1FAE5")))
        wm.auto_filter.ref = f"A{rm}:{L(len(titres))}{dern_m}"
    for k, w in enumerate((20, 18, 34, 10, 42, 28, 10, 11, 12, 9, 11, 9, 46), 1):
        wm.column_dimensions[L(k)].width = w
    wm.freeze_panes = wm.cell(rm + 1, 1)

    # ---------- Mercuriale contre facture ----------
    wf = wb.create_sheet("Mercuriale contre facture")
    titres = ["Fournisseur", "Réf.", "Libellé", "Date du cadencier", "Prix cadencier", "Date de la facture",
              "Prix facturé", "Écart %", "Promo", "Nombre d'achats", "Unité", "Note"]
    entete(wf, 1, titres)
    controle = sorted(lire("controle_factures.csv") if factures else [], key=lambda x: (x["fournisseur"], -abs(float(x["ecart_pct"]))))
    for i, x in enumerate(controle, 2):
        # Le prix du cadencier est ramené à l'unité de la facture, pour comparer ce qui est comparable
        ramene = float(x.get("prix_cadencier_ramene") or x["prix_cadencier"])
        unite_cad = x.get("unite_cadencier", "")
        wf.cell(i, 1, x["fournisseur"])
        wf.cell(i, 2, x["ref"])
        wf.cell(i, 3, x["libelle"])
        wf.cell(i, 4, jour(x["date_tarif"])).number_format = "dd/mm/yyyy"
        wf.cell(i, 5, ramene).number_format = "0.000"
        wf.cell(i, 6, jour(x["date_facture"])).number_format = "dd/mm/yyyy"
        wf.cell(i, 7, float(x["prix_facture"])).number_format = "0.000"
        wf.cell(i, 8, f"=(G{i}-E{i})/E{i}").number_format = "+0.0%;-0.0%;0.0%"
        wf.cell(i, 9, "oui" if x["promo"] == "1" else "")
        wf.cell(i, 10, int(x["nb_achats"]))
        wf.cell(i, 11, x.get("unite_facture") or unite_cad)
        if float(x.get("facteur") or 1) != 1:
            wf.cell(i, 12, f"cadencier : {float(x['prix_cadencier']):.3f} €".replace(".", ",")
                           + f" les {unite_cad}, ramené à 1 {x['unite_facture']}")
    habiller(wf, 2, len(controle) + 1, len(titres))
    for i, w in enumerate((21, 9, 44, 11, 10, 11, 10, 9, 7, 9, 9, 44), 1):
        wf.column_dimensions[L(i)].width = w
    wf.freeze_panes = "A2"
    wf.auto_filter.ref = f"A1:{L(len(titres))}{len(controle) + 1}"

    # ---------- Évolutions entre deux cadenciers d'un même fournisseur ----------
    we = wb.create_sheet("Évolutions")
    titres = ["Fournisseur", "Réf.", "Libellé", "Ancien cadencier", "Ancien prix", "Nouveau cadencier", "Nouveau prix",
              "Évolution", "Unité"]
    entete(we, 1, titres)
    groupes = {}
    for t in tarifs:
        groupes.setdefault((t["fournisseur"], cle_ref(t["ref"])), []).append(t)
    i = 2
    for (four, ref), g in sorted(groupes.items()):
        if len(g) < 2:
            continue
        g.sort(key=lambda x: x["date_tarif"])
        avant, apres = g[0], g[-1]
        # On compare au kg ou au litre quand c'est possible : un pack et une bouteille ne se comparent pas au prix brut
        if avant["prix_unitaire"] and apres["prix_unitaire"] and avant["unite"] == apres["unite"]:
            v0, v1, unite = float(avant["prix_unitaire"]), float(apres["prix_unitaire"]), "€/" + apres["unite"]
        elif avant["unite_prix"] == apres["unite_prix"]:
            v0, v1, unite = float(avant["prix"]), float(apres["prix"]), apres["unite_prix"]
        else:
            continue
        if round(v0, 4) == round(v1, 4):
            continue
        we.cell(i, 1, four)
        we.cell(i, 2, apres["ref"])
        we.cell(i, 3, apres["libelle"])
        we.cell(i, 4, jour(avant["date_tarif"])).number_format = "dd/mm/yyyy"
        we.cell(i, 5, v0).number_format = "0.000"
        we.cell(i, 6, jour(apres["date_tarif"])).number_format = "dd/mm/yyyy"
        we.cell(i, 7, v1).number_format = "0.000"
        we.cell(i, 8, f"=(G{i}-E{i})/E{i}").number_format = "+0.0%;-0.0%;0.0%"
        we.cell(i, 9, unite)
        i += 1
    habiller(we, 2, i - 1, len(titres))
    for k, w in enumerate((21, 9, 44, 12, 10, 12, 10, 10, 9), 1):
        we.column_dimensions[L(k)].width = w
    we.freeze_panes = "A2"

    # ---------- Références sans équivalent ----------
    ws = wb.create_sheet("Sans équivalent")
    titres = ["Fournisseur", "Réf.", "Libellé", "Prix HT", "Unité de prix", "Prix par kg ou litre", "Unité", "Date du tarif"]
    entete(ws, 1, titres)
    pris = {(l[0]["fournisseur"], cle_ref(l[0]["ref"])) for p in produits.values() for l in p["lignes"]}
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

    # ---------- Mercuriale ----------
    outils = dict(entete=entete, habiller=habiller, jour=jour, Font=Font, PatternFill=PatternFill,
                  FormulaRule=FormulaRule, L=L)
    comp = dict(premiere=premiere, dern=dern, moins_cher=L(cm)) if dern >= premiere else None
    nb_merc = ajouter_onglets(wb, ICI, recent, produits, outils, comp)
    # Meilleurs prix juste après le Comparatif, qu'il résume
    wb.move_sheet("Meilleurs prix", wb.sheetnames.index("Comparatif") + 1 - wb.sheetnames.index("Meilleurs prix"))

    # ---------- À valider : le client tranche chaque correspondance ----------
    wv = wb.create_sheet(ONGLET, 0)
    wv["A1"] = f"Correspondances à valider · {client}"
    wv["A1"].font = Font(name=POLICE, bold=True, size=14)
    wv["A2"] = ("Pour chaque produit, choisir dans la colonne Décision : « Même produit » ou « Remplaçable en cuisine » "
                "(on compare les prix), ou « Pas comparable » (affiché pour mémoire, jamais compté). Le Comparatif se met "
                "à jour tout de suite. Ensuite, dire à Claude que c'est fait : il reporte les choix et met la page à jour.")
    wv["A2"].font = Font(name=POLICE, size=10, color="55606E")
    wv["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    wv.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3 + n + 4)
    wv.row_dimensions[2].height = 42
    titres = (["Famille", "Produit", "Unité"] + fournisseurs
              + ["Ce qui les distingue", "Proposition de Claude", "Décision du client", "Commentaire du client"])
    entete(wv, RV, titres)
    for i, f in enumerate(fournisseurs):
        wv.cell(RV, 4 + i).fill = PatternFill("solid", fgColor=COULEURS[f])
    cdist, cprop, cdec, ccom = (3 + n + k for k in range(1, 5))
    for nom, p in a_valider:
        rv = rang_val[nom]
        wv.cell(rv, 1, p["famille"])
        wv.cell(rv, 2, nom).font = Font(bold=True)
        wv.cell(rv, 3, p["unite"])
        for a_, t in p["lignes"]:
            c = wv.cell(rv, 4 + fournisseurs.index(a_["fournisseur"]), f"{t['libelle']} (réf. {t['ref']})")
            c.alignment = Alignment(wrap_text=True)
        wv.cell(rv, cdist, p["note"]).alignment = Alignment(wrap_text=True)
        wv.cell(rv, cprop, CHOIX[p["proposition"]])
        wv.cell(rv, cdec, p["decision"] or None).font = Font(bold=True)
        wv.cell(rv, ccom, p["commentaire"] or None).alignment = Alignment(wrap_text=True)
    dern_v = RV + len(a_valider)
    # Compteurs : ce qui reste à trancher, hors produits proposés identiques
    plage_dec, plage_prop = f"{L(cdec)}{RV + 1}:{L(cdec)}{max(dern_v, RV + 1)}", f"{L(cprop)}{RV + 1}:{L(cprop)}{max(dern_v, RV + 1)}"
    wv["A4"] = "Décisions prises"
    wv["B4"] = f"=COUNTIF({plage_dec},\"?*\")"
    wv["A5"] = "Encore à trancher"
    wv["B5"] = f"=COUNTIFS({plage_dec},\"\",{plage_prop},\"<>{CHOIX['identique']}\")"
    for c in ("A4", "A5"):
        wv[c].font = Font(name=POLICE, bold=True, size=10)
    for c in ("B4", "B5"):
        wv[c].font = Font(name=POLICE, bold=True, size=12)
        wv[c].alignment = Alignment(horizontal="left")
    habiller(wv, RV + 1, dern_v, len(titres))
    for rr in range(RV + 1, dern_v + 1):
        for c in [2] + list(range(4, ccom + 1)):
            wv.cell(rr, c).alignment = Alignment(wrap_text=True, vertical="center")
        wv.cell(rr, 2).font = Font(name=POLICE, size=10, bold=True)
        wv.cell(rr, cdec).font = Font(name=POLICE, size=10, bold=True)
    if dern_v > RV:
        liste = DataValidation(type="list", formula1='"' + ",".join(CHOIX.values()) + '"', allow_blank=True,
                               showErrorMessage=True, errorTitle="Choix inconnu",
                               error="Choisir Même produit, Remplaçable en cuisine ou Pas comparable.",
                               showInputMessage=True, promptTitle="Décision du client",
                               prompt="Laisser vide pour garder la proposition de Claude.")
        wv.add_data_validation(liste)
        liste.add(plage_dec)
        # Jaune : à trancher. Vert : tranché comme Claude le proposait. Orange : le client a changé d'avis.
        d0, p0 = f"${L(cdec)}{RV + 1}", f"${L(cprop)}{RV + 1}"
        wv.conditional_formatting.add(plage_dec, FormulaRule(
            formula=[f'AND({d0}="",{p0}<>"{CHOIX["identique"]}")'], fill=PatternFill("solid", bgColor="FEF3C7")))
        wv.conditional_formatting.add(plage_dec, FormulaRule(
            formula=[f'AND({d0}<>"",{d0}={p0})'], fill=PatternFill("solid", bgColor="D1FAE5")))
        wv.conditional_formatting.add(plage_dec, FormulaRule(
            formula=[f'AND({d0}<>"",{d0}<>{p0})'], fill=PatternFill("solid", bgColor="FED7AA")))
        wv.auto_filter.ref = f"A{RV}:{L(ccom)}{dern_v}"
    largeurs = {1: 18, 2: 30, 3: 10, cdist: 40, cprop: 16, cdec: 20, ccom: 30}
    for i in range(n):
        largeurs[4 + i] = 30
    for c, w in largeurs.items():
        wv.column_dimensions[L(c)].width = w
    wv.freeze_panes = wv.cell(RV + 1, 3)

    # ---------- Lisez-moi ----------
    wl = wb.create_sheet("Lisez-moi", 0)
    wl.column_dimensions["A"].width = 26
    wl.column_dimensions["B"].width = 92
    wl["A1"] = f"Comparatif à partir de cadenciers · {client}"
    wl["A1"].font = Font(name=POLICE, bold=True, size=14)
    lignes = [
        ("But", "Comparer les prix des fournisseurs à partir de leurs cadenciers" + (", contrôlés sur les factures." if factures else ". Aucune facture n'a servi au contrôle.")),
        ("À valider", "Les correspondances proposées par Claude, une ligne par produit, avec le libellé de chaque fournisseur. Le client choisit dans la colonne Décision : Même produit, Remplaçable en cuisine ou Pas comparable. Le Comparatif et les Meilleurs prix suivent tout de suite. Claude reporte ensuite ces choix dans ses fichiers : sans cela, ils seraient perdus à la prochaine mise à jour."),
        ("Comparatif", "Un produit par ligne, un fournisseur par colonne. Le prix le plus bas ressort en vert."),
        ("Meilleurs prix", "Pour chaque produit comparable, le fournisseur le moins cher, avec sa référence, son libellé, son conditionnement et son prix, rangés par fournisseur : ce qu'il vaut mieux acheter chez chacun. À relire avec la colonne Résultat solide. Les produits non comparables sont en bas, en gris, et s'y ajoutent si le client change d'avis dans l'onglet À valider. « égalité » : un autre fournisseur fait le même prix."),
        ("Mercuriale", "Tous les fournisseurs et tous leurs tarifs dans un seul tableau, classés par famille, avec le dernier prix payé (ramené à l'unité de prix du tarif) et le verdict du comparatif."),
        ("Détail", "D'où vient chaque prix : libellé, date, quantité, facture. Les prix en bleu se corrigent à la main, le reste suit."),
        ("Mercuriale contre facture", "Prix de la mercuriale (le tarif du fournisseur) face au dernier prix payé, pour chaque référence retrouvée sur une facture. Quand la facture compte dans une autre unité (au kg pour une pièce de 1,4 kg, à la brique pour un pack de 6), le tarif est ramené à l'unité facturée : la colonne Note le dit."),
        ("Évolutions", "Prix qui ont changé entre deux cadenciers d'un même fournisseur."),
        ("Sans équivalent", "Références lues mais comparées à rien, faute de produit équivalent chez un autre fournisseur."),
        ("Fond jaune", "« Mois entre les prix » supérieur à 3 : les prix comparés n'ont pas le même âge, le résultat est indicatif. « Prix à vérifier » : unité non confirmée par une facture. « Offres expirées » : prix lu sur une offre dont la fin de validité est passée."),
        ("Fond orange", "Écart supérieur à 50 % : souvent une erreur d'unité ou deux produits différents. À contrôler avant de s'y fier."),
        ("Résultat solide", "Prix à 3 mois d'écart au plus, aucune offre expirée, aucune unité à vérifier. Mêmes règles que la page interactive."),
        ("Prix retenu", f"Le prix du cadencier, ou celui de la dernière facture hors promotion si elle est plus récente et à moins de {SEUIL_UNITE:.0%} du cadencier. Une facture au kg, au litre ou à l'unité d'un lot est d'abord ramenée à l'unité du cadencier. Plus loin, la facture est sans doute dans une autre unité : son prix n'est pas retenu et la note de lecture le signale."),
        ("Non comparable", "Produit affiché pour mémoire, jamais compté dans le bilan."),
        ("Tableur", "Ce classeur s'ouvre dans Excel, LibreOffice ou Google Sheets. Pour Google Sheets : dans Google Drive, Nouveau, Importer un fichier, l'ouvrir, puis Fichier, Enregistrer au format Google Sheets. Pour une mise à jour, dans la feuille Google : Fichier, Importer, Remplacer la feuille de calcul, ce qui garde la même adresse. Une correction faite à la main dans le classeur est perdue à la mise à jour : la signaler à Claude pour qu'il la reporte dans ses fichiers."),
        ("Limites", "Les correspondances entre produits ont été proposées par Claude et restent à valider. Quand un document ne porte pas de date, c'est celle du fichier qui est utilisée. Ce sont des prix HT, hors franco, remises de fin d'année et frais de livraison."),
    ]
    for i, (k, v) in enumerate(lignes, 3):
        wl.cell(i, 1, k).font = Font(name=POLICE, bold=True, size=10)
        c = wl.cell(i, 2, v)
        c.font = Font(name=POLICE, size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        wl.cell(i, 1).alignment = Alignment(vertical="top")
    # Tant que le comparatif est provisoire, le classeur s'ouvre sur ce qui reste à valider
    wb.active = wb.sheetnames.index(ONGLET if config.get("provisoire", True) else "Comparatif")
    # Réglages lus dans config.json : la page vérifie qu'ils n'ont pas changé depuis
    wb.properties.keywords = json.dumps({k: config.get(k) for k in CONFIG_UTILE}, ensure_ascii=False)
    hors_liste = fonctions_incompatibles(wb)
    if hors_liste:
        print("ATTENTION : fonctions peut-être absentes de Google Sheets : " + ", ".join(hors_liste))
    wb.save(SORTIE)
    print(f"{len(produits)} produits, {n} fournisseurs, {dern_d - 1} lignes de détail, "
          f"mercuriale de {nb_merc} références -> {SORTIE.name}")


if __name__ == "__main__":
    main()

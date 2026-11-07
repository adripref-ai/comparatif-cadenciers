"""Mercuriale et cadencier de commande, tirés des tarifs et des achats relevés sur les factures.

La mercuriale met tous les fournisseurs et tous leurs tarifs dans un seul tableau.
Le cadencier de commande part de ce que le client achète vraiment : consommation estimée,
jours de livraison, stock, et quantité à commander pour tenir jusqu'à la livraison suivante.

Les réglages modifiables vivent dans le dossier du client :
- fournisseurs.csv : fournisseur, jours_livraison, livraisons_par_semaine, delai_jours, franco_eur, source
- stocks.csv : fournisseur, ref, libelle, stock, unite, date_stock, conso_semaine
- familles.csv (facultatif) : fournisseur, ref, famille
preparer_cadencier.py les crée et les complète sans effacer ce que le client a déjà saisi.
"""
import csv
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

from lecteurs import cle_ref, nombre

JOURS = ["lun", "mar", "mer", "jeu", "ven", "sam", "dim"]
CHAMPS_FOURNISSEURS = ["fournisseur", "jours_livraison", "livraisons_par_semaine", "delai_jours", "franco_eur", "source"]
CHAMPS_STOCKS = ["fournisseur", "ref", "libelle", "stock", "unite", "date_stock", "conso_semaine"]
DELAI_DEFAUT = 1
SECURITE_DEFAUT = 2


def lire_csv(chemin):
    chemin = Path(chemin)
    return list(csv.DictReader(open(chemin, encoding="utf-8-sig"))) if chemin.exists() else []


def ecrire_csv(chemin, champs, lignes):
    with open(chemin, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, champs, extrasaction="ignore")
        w.writeheader()
        w.writerows(lignes)


def nombre_ou_vide(v):
    try:
        return nombre(v) if str(v or "").strip() else None
    except ValueError:
        return None


def livraisons_observees(achats):
    """Par fournisseur : dates de livraison distinctes, jours de la semaine et rythme observé."""
    dates = defaultdict(set)
    for a in achats:
        dates[a["fournisseur"]].add(a["date_livraison"])
    out = {}
    for four, ds in dates.items():
        ds = sorted(date.fromisoformat(d) for d in ds)
        ecart = (ds[-1] - ds[0]).days
        par_semaine = round(7 * (len(ds) - 1) / ecart, 2) if len(ds) > 1 and ecart else None
        jours = [JOURS[j] for j in sorted({d.weekday() for d in ds})]
        out[four] = dict(dates=ds, jours=" ".join(jours), par_semaine=par_semaine)
    return out


def reglages_fournisseurs(ici):
    """fournisseurs.csv, avec un rythme de livraison utilisable pour chacun (ou None)."""
    out = {}
    for f in lire_csv(Path(ici) / "fournisseurs.csv"):
        lps = nombre_ou_vide(f.get("livraisons_par_semaine"))
        if lps is None and f.get("jours_livraison", "").strip():
            lps = float(len([j for j in f["jours_livraison"].replace(",", " ").split() if j.strip("?")]))
        out[f["fournisseur"]] = dict(f, lps=lps, delai=nombre_ou_vide(f.get("delai_jours")))
    return out


def estimer_consommations(achats, reglages):
    """Consommation par semaine de chaque référence achetée.

    Elle se mesure sur la période où l'on observe le fournisseur : de sa première à sa dernière
    livraison, plus un intervalle entre deux livraisons. Un sirop acheté une seule fois en trois
    semaines de livraisons compte donc pour un tiers de bouteille par semaine, pas pour une bouteille."""
    obs = livraisons_observees(achats)
    groupes = defaultdict(list)
    for a in achats:
        groupes[(a["fournisseur"], cle_ref(a["ref"]))].append(a)
    out = {}
    for (four, ref), lignes in groupes.items():
        lignes.sort(key=lambda a: a["date_livraison"])
        o = obs[four]
        n_four = len(o["dates"])
        lps = reglages.get(four, {}).get("lps")
        if n_four > 1:
            intervalle = (o["dates"][-1] - o["dates"][0]).days / (n_four - 1)
        elif lps:
            intervalle = 7 / lps
        else:
            intervalle = None
        qtes = [nombre_ou_vide(a["qte"]) for a in lignes]
        total = sum(q for q in qtes if q is not None)
        colisages = Counter(a["colisage"] for a in lignes if a["colisage"])
        dernier = lignes[-1]
        if intervalle and None not in qtes:
            periode = (o["dates"][-1] - o["dates"][0]).days + intervalle
            conso = round(total / periode * 7, 2)
            base = (f"{total:g} ({dernier['unite']}) achetés en {len(lignes)} fois sur {periode:g} jours, "
                    f"{n_four} livraison(s) de {four}")
            if n_four == 1:
                base += ", en supposant que la livraison couvre jusqu'à la suivante"
        else:
            conso = None
            base = ("quantités absentes des factures" if None in qtes else
                    "une seule livraison connue et rythme inconnu : saisir la consommation")
        # Le dernier prix hors promotion, s'il existe, est celui qu'on paiera à la prochaine commande
        normal = [a for a in lignes if str(a.get("promo", "0")) in ("0", "")] or lignes
        out[(four, ref)] = dict(fournisseur=four, ref=dernier["ref"], libelle=dernier["libelle"], unite=dernier["unite"],
                                colisage=nombre_ou_vide(colisages.most_common(1)[0][0]) if colisages else None,
                                conso=conso, base=base, nb=len(lignes), dernier_prix=float(normal[-1]["prix"]),
                                date_dernier=dernier["date_livraison"])
    return out


# ---------- Onglets du classeur ----------

def ajouter_onglets(wb, ici, config, recent, produits, outils, comparatif):
    """Ajoute Mercuriale et Cadencier de commande au classeur.
    outils : fonctions de mise en forme de comparer.py (entete, habiller, jour, Font, PatternFill, L).
    comparatif : rangs et colonnes de l'onglet Comparatif, pour reprendre son verdict par formule."""
    entete, habiller, jour = outils["entete"], outils["habiller"], outils["jour"]
    Font, PatternFill, FormulaRule, L = outils["Font"], outils["PatternFill"], outils["FormulaRule"], outils["L"]
    bleu = Font(color="0000FF")
    familles = {(f["fournisseur"], cle_ref(f["ref"])): f["famille"] for f in lire_csv(Path(ici) / "familles.csv")}
    appariement = {}
    for nom, p in produits.items():
        for a, _t in p["lignes"]:
            appariement[(a["fournisseur"], cle_ref(a["ref"]))] = (nom, p["famille"])
    achats = lire_csv(Path(ici) / "achats.csv")
    reglages = reglages_fournisseurs(ici)
    conso = estimer_consommations(achats, reglages)

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
        c = conso.get(cle)
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

    # ---------- Cadencier de commande : stocks, consommations, livraisons ----------
    wk = wb.create_sheet("Cadencier de commande", 2)
    tit = ["Famille", "Produit comparé", "Fournisseur", "Réf.", "Libellé", "Unité", "Colisage",
           "Consommation par semaine", "Base de l'estimation", "Jours de livraison", "Livraisons par semaine",
           "Délai de commande (jours)", "Sécurité (jours)", "Stock", "Date du stock",
           "Besoin jusqu'à la livraison suivante", "À commander", "Dernier prix payé", "Montant estimé",
           "Moins cher (comparatif)"]
    entete(wk, 1, tit)
    stocks = {(s["fournisseur"], cle_ref(s["ref"])): s for s in lire_csv(Path(ici) / "stocks.csv")}
    securite = config.get("securite_jours", SECURITE_DEFAUT)
    cles = list(dict.fromkeys(list(conso) + list(stocks)))
    lignes = []
    for cle in cles:
        c, s = conso.get(cle), stocks.get(cle, {})
        fam, nom = famille_produit(cle)
        lignes.append((fam == "À classer", fam, nom or "~", cle[0], (c or s)["libelle"], cle, c, s))
    lignes.sort(key=lambda x: x[:5])
    for r, (_, fam, nom, four, lib, cle, c, s) in enumerate(lignes, 2):
        reg = reglages.get(four, {})
        conso_saisie = nombre_ou_vide(s.get("conso_semaine"))
        valeurs = [fam, "" if nom == "~" else nom, four, (c or s)["ref"], lib, (c or s).get("unite", ""),
                   c["colisage"] if c else None,
                   conso_saisie if conso_saisie is not None else (c["conso"] if c else None),
                   "saisie dans stocks.csv" if conso_saisie is not None else (c["base"] if c else "aucun achat relevé"),
                   reg.get("jours_livraison", ""), reg.get("lps"),
                   reg.get("delai") if reg.get("delai") is not None else DELAI_DEFAUT, securite,
                   nombre_ou_vide(s.get("stock")), jour(s.get("date_stock") or ""),
                   f'=IF(OR(H{r}="",K{r}="",K{r}=0),"",ROUND(H{r}/7*(7/K{r}+L{r}+M{r}),2))',
                   f'=IF(P{r}="","conso. ou livraisons à saisir",IF(N{r}="","stock à saisir",'
                   f'CEILING(MAX(0,P{r}-N{r}),IF(G{r}="",1,G{r}))))',
                   c["dernier_prix"] if c else None,
                   f'=IF(AND(ISNUMBER(Q{r}),ISNUMBER(R{r})),ROUND(Q{r}*R{r},2),"")',
                   verdict(r)]
        for k, v in enumerate(valeurs, 1):
            cell = wk.cell(r, k, v)
            if k in (7, 8, 11, 12, 13, 14):
                cell.font = bleu
        for k, fmt in ((8, "0.00"), (11, "0.##"), (16, "0.00"), (18, "0.000"), (19, "0.00")):
            wk.cell(r, k).number_format = fmt
        wk.cell(r, 15).number_format = "dd/mm/yyyy"
    dern_k = len(lignes) + 1
    habiller(wk, 2, dern_k, len(tit))
    # Total estimé de la prochaine commande, par fournisseur
    r = dern_k + 2
    wk.cell(r, 3, "Prochaine commande estimée").font = Font(bold=True)
    for four in dict.fromkeys(x[3] for x in lignes):
        r += 1
        wk.cell(r, 3, four)
        wk.cell(r, 19, f'=SUMIF($C$2:$C${dern_k},C{r},$S$2:$S${dern_k})').number_format = "0.00"
    for k, w in enumerate((18, 26, 19, 9, 36, 8, 9, 11, 34, 12, 10, 10, 9, 9, 11, 12, 13, 10, 11, 19), 1):
        wk.column_dimensions[L(k)].width = w
    wk.freeze_panes = "F2"
    if dern_k >= 2:
        wk.auto_filter.ref = f"A1:{L(len(tit))}{dern_k}"
        wk.conditional_formatting.add(f"Q2:Q{dern_k}", FormulaRule(formula=['AND(ISNUMBER($Q2),$Q2>0)'],
                                      fill=PatternFill("solid", bgColor="D1FAE5"), font=Font(bold=True)))
        wk.conditional_formatting.add(f"Q2:Q{dern_k}", FormulaRule(formula=['ISTEXT($Q2)'],
                                      fill=PatternFill("solid", bgColor="FEF3C7")))
    return len(recent), len(lignes)

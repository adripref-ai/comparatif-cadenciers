#!/usr/bin/env python3
"""Reprend dans les fichiers du client ce qu'il a saisi dans son tableur ou sur la page.

Usage : reprendre_stocks.py <dossier du client> <fichier> [--date JJ/MM/AAAA]
<fichier> peut être :
- le classeur rempli par le client dans Excel ou LibreOffice, ou téléchargé depuis Google Sheets
  (Fichier, Télécharger, Microsoft Excel) ;
- l'onglet « Cadencier de commande » exporté en CSV (virgules ou points-virgules, UTF-8, Windows ou Mac),
  ou relu dans Google Sheets et recopié en CSV, ligne de titres comprise ;
- le texte du bouton « Copier les stocks » de la page, enregistré en CSV.

Dans un classeur ou un export du cadencier, seules les cases que le client a changées sont reprises :
la colonne cachée « Valeurs d'origine » garde ce que le classeur contenait à sa construction.
- stock et date du stock, consommation, colisage -> stocks.csv ;
- livraisons par semaine et délai -> fournisseurs.csv, quand toutes les lignes d'un fournisseur disent pareil ;
- sécurité : signalée seulement, car elle vaut pour tout le cadencier (config.json, securite_jours).
Un stock changé sans changer sa date prend la date de --date, sinon celle du jour. Une case vidée
n'efface rien. Un stock plus ancien que celui de stocks.csv n'est pas repris. Tout ce qui n'est pas
repris est signalé.

comparer.py fait cette reprise tout seul sur le classeur du dossier avant de le reconstruire.
"""
import csv
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

sys.dont_write_bytecode = True

from lecteurs import cle_ref
from suivi import (CHAMPS_FOURNISSEURS, CHAMPS_STOCKS, ESTIMATION, ORIGINE, ecrire_csv, lire_csv,
                   nombre_ou_vide)

ONGLET = "Cadencier de commande"
NOMS = {"fournisseur": "Fournisseur", "ref": "Réf.", "stock": "Stock", "date_stock": "Date du stock",
        "conso": "Consommation par semaine", "colis": "Colisage", "lps": "Livraisons par semaine",
        "delai": "Délai de commande (jours)", "secu": "Sécurité (jours)", "origine": ORIGINE,
        "estimation": ESTIMATION}
# Écart toléré : un export CSV écrit les nombres tels qu'ils sont affichés (0,13 pour 0,125)
TOLERANCE = 0.006
MOIS = {"janv": 1, "jan": 1, "févr": 2, "fevr": 2, "fév": 2, "fev": 2, "mars": 3, "mar": 3, "avr": 4, "mai": 5,
        "juin": 6, "juil": 7, "août": 8, "aout": 8, "sept": 9, "sep": 9, "oct": 10, "nov": 11, "déc": 12, "dec": 12}


class Illisible(Exception):
    pass


def lire_texte(chemin):
    """CSV en UTF-8, ou enregistré par Excel sous Windows ou sur Mac."""
    brut = chemin.read_bytes()
    try:
        return brut.decode("utf-8-sig")
    except UnicodeDecodeError:
        pass
    for codage in ("cp1252", "mac_roman"):
        texte = brut.decode(codage, errors="replace")
        if "Réf." in texte.split("\n", 1)[0] or "Libellé" in texte.split("\n", 1)[0]:
            return texte
    return brut.decode("cp1252", errors="replace")


def lire_saisies(chemin):
    """(lignes, a_origine) : lignes sous forme de dictionnaires aux clés de NOMS."""
    if chemin.suffix.lower() in (".xlsx", ".xlsm"):
        import openpyxl
        try:
            wb = openpyxl.load_workbook(chemin, data_only=True)
        except Exception as e:
            raise Illisible(f"{chemin.name} n'est pas un classeur lisible ({e.__class__.__name__}). "
                            "Le réenregistrer au format .xlsx.")
        if ONGLET not in wb.sheetnames:
            raise Illisible(f"Onglet « {ONGLET} » absent de {chemin.name}.")
        brut = [list(r) for r in wb[ONGLET].iter_rows(values_only=True)]
    else:
        texte = lire_texte(chemin)
        tete = texte.split("\n", 1)[0]
        brut = list(csv.reader(texte.splitlines(), delimiter=";" if tete.count(";") > tete.count(",") else ","))
    if not brut:
        raise Illisible(f"{chemin.name} est vide.")
    entete = [str(h or "").strip() for h in brut[0]]
    index = {}
    for cle, nom in NOMS.items():
        for i, h in enumerate(entete):
            if h and h in (nom, cle):
                index[cle] = i
    if not {"fournisseur", "ref", "stock"} <= set(index):
        raise Illisible("Colonnes Fournisseur, Réf. et Stock introuvables en première ligne : "
                        "est-ce bien le cadencier de commande ?")
    out = []
    for r in brut[1:]:
        x = {cle: (r[i] if i < len(r) else None) for cle, i in index.items()}
        if str(x["fournisseur"] or "").strip() and str(x["ref"] or "").strip():
            out.append(x)
    return out, "origine" in index


def vide(v):
    return v is None or str(v).strip() == ""


def lire_nombre(v):
    """Nombre saisi, ou Illisible. « 1 200 » et « 3,5 » sont acceptés, « 12 packs » non."""
    if isinstance(v, (int, float)):
        return float(v)
    t = re.sub(r"[\s  ]", "", str(v)).replace(",", ".")
    if not re.fullmatch(r"-?\d+(\.\d+)?", t):
        raise Illisible(str(v).strip())
    return float(t)


def lire_date(v):
    """Date saisie en AAAA-MM-JJ, ou Illisible. Accepte 06/11/2026, 06.11.2026, 6-11-26, 2026-11-06,
    une date avec son heure, « 6 nov. 2026 » et le numéro de série d'un tableur (46332)."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    t = str(v).strip()
    if isinstance(v, (int, float)) or re.fullmatch(r"\d{5}(\.\d+)?", t):
        n = float(t)
        if 20000 < n < 80000:
            return (date(1899, 12, 30) + timedelta(days=int(n))).isoformat()
        raise Illisible(t)
    t = re.sub(r"[ T]\d{1,2}:\d{2}(:\d{2})?$", "", t)
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", t)
    if m:
        a, mo, j = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})", t)
        if m:
            j, mo, a = map(int, m.groups())
        else:
            m = re.fullmatch(r"(\d{1,2})\s+([a-zéû]+)\.?\s+(\d{4})", t.lower())
            if not m or m.group(2)[:4].rstrip(".") not in MOIS and m.group(2)[:3] not in MOIS:
                raise Illisible(t)
            j, a = int(m.group(1)), int(m.group(3))
            mo = MOIS.get(m.group(2)[:4].rstrip("."), MOIS.get(m.group(2)[:3]))
    if a < 100:
        a += 2000
    try:
        return date(a, mo, j).isoformat()
    except ValueError:
        raise Illisible(t)


def differe(a, b):
    if a is None or b is None:
        return (a is None) != (b is None)
    return abs(a - b) > TOLERANCE


def reprendre(ici, lignes, a_origine, defaut, nom, auto=False):
    """Reporte les saisies dans stocks.csv et fournisseurs.csv. Renvoie les lignes du compte rendu."""
    ici = Path(ici)
    stocks = lire_csv(ici / "stocks.csv")
    par_cle = {(s["fournisseur"], cle_ref(s["ref"])): s for s in stocks}
    msg, repris, illisibles, inconnues, perimes, consos, colis, secus = [], [], [], [], [], [], [], []
    reglages = {}  # fournisseur -> {"lps": {valeurs}, "delai": {valeurs}}
    a_estimation = any("estimation" in x for x in lignes)
    for x in lignes:
        four, ref = str(x["fournisseur"]).strip(), str(x["ref"]).strip()
        if isinstance(x["ref"], float) and x["ref"].is_integer():
            ref = str(int(x["ref"]))
        nom_l = f"{four} {ref}"
        s = par_cle.get((four, cle_ref(ref)))
        try:
            o = json.loads(x["origine"]) if a_origine and not vide(x.get("origine")) else None
        except ValueError:
            o = None
        if s is None:  # faute de frappe, ou référence pas encore achetée : on ne l'invente pas
            if not vide(x.get("stock")) and (o is None or differe_brut(x.get("stock"), o.get("stock"))):
                inconnues.append(nom_l)
            continue
        # ----- Stock et date -----
        ancien = nombre_ou_vide(s.get("stock"))
        ancienne_date = s.get("date_stock") or ""
        stock, d_saisie = None, None
        if not vide(x.get("stock")):
            try:
                stock = lire_nombre(x["stock"])
            except Illisible as e:
                illisibles.append(f"{nom_l} « {e} »")
        if not vide(x.get("date_stock")):
            try:
                d_saisie = lire_date(x["date_stock"])
            except Illisible as e:
                if stock is not None:
                    msg.append(f"Date illisible pour {nom_l} : « {e} ». Stock daté du {defaut}.")
        if stock is not None:
            if o is not None:
                change = differe(stock, o.get("stock"))
                d_origine = o.get("date") or ""
                nouvelle_d = d_saisie and d_saisie != d_origine
                if change or nouvelle_d:
                    d = d_saisie if nouvelle_d else defaut
                    if not differe(stock, ancien) and (not d or d <= ancienne_date):
                        pass  # déjà dans stocks.csv
                    # stocks.csv a changé depuis la construction du classeur, et sa date est plus récente
                    elif differe(ancien, o.get("stock")) and ancienne_date and ancienne_date >= d:
                        perimes.append(f"{nom_l} : {stock:g} au {d}, mais stocks.csv a {ancien:g} au {ancienne_date}")
                    else:
                        s["stock"], s["date_stock"] = f"{stock:g}", d
                        repris.append(nom_l)
            else:
                if not differe(stock, ancien):
                    if d_saisie and d_saisie > ancienne_date:
                        s["date_stock"] = d_saisie
                        repris.append(nom_l)
                else:
                    d = d_saisie if d_saisie and d_saisie != ancienne_date else defaut
                    if ancienne_date and d < ancienne_date:
                        perimes.append(f"{nom_l} : {stock:g} au {d}, mais stocks.csv a {ancien:g} au {ancienne_date}")
                    else:
                        s["stock"], s["date_stock"] = f"{stock:g}", d
                        repris.append(nom_l)
        # ----- Consommation, colisage et réglages : seulement s'ils ont changé -----
        reference = o.get("conso") if o is not None else (nombre_ou_vide(x.get("estimation")) if a_estimation else None)
        if not vide(x.get("conso")) and (o is not None or a_estimation):
            try:
                conso = lire_nombre(x["conso"])
                if differe(conso, reference):
                    s["conso_semaine"] = f"{conso:g}"
                    consos.append(f"{nom_l} : {conso:g} par semaine")
            except Illisible as e:
                illisibles.append(f"{nom_l} (consommation) « {e} »")
        if o is None:
            continue
        for cle, titre in (("colis", "colisage"), ("lps", "livraisons par semaine"), ("delai", "délai"),
                           ("secu", "sécurité")):
            if vide(x.get(cle)):
                continue
            try:
                v = lire_nombre(x[cle])
            except Illisible as e:
                illisibles.append(f"{nom_l} ({titre}) « {e} »")
                continue
            if not differe(v, o.get(cle)):
                continue
            if cle == "colis":
                s["colisage"] = f"{v:g}"
                colis.append(f"{nom_l} : colis de {v:g}")
            elif cle == "secu":
                secus.append(f"{nom_l} : {v:g} jours")
            else:
                reglages.setdefault(four, {}).setdefault(cle, set()).add(v)

    stocks.sort(key=lambda s: (s["fournisseur"], s["libelle"]))
    ecrire_csv(ici / "stocks.csv", CHAMPS_STOCKS, stocks)

    if reglages:
        fournisseurs = lire_csv(ici / "fournisseurs.csv")
        champ = {"lps": "livraisons_par_semaine", "delai": "delai_jours"}
        for f in fournisseurs:
            for cle, valeurs in reglages.get(f["fournisseur"], {}).items():
                if len(valeurs) == 1:
                    v = valeurs.pop()
                    f[champ[cle]] = f"{v:g}"
                    f["source"] = f"corrigé dans le tableur ({nom})"
                    msg.append(f"{f['fournisseur']} : {champ[cle].replace('_', ' ')} = {v:g}, repris dans fournisseurs.csv.")
                else:
                    msg.append(f"{f['fournisseur']} : {champ[cle].replace('_', ' ')} différent selon les lignes "
                               f"({', '.join(f'{v:g}' for v in sorted(valeurs))}), non repris : demander la bonne valeur.")
        ecrire_csv(ici / "fournisseurs.csv", CHAMPS_FOURNISSEURS, fournisseurs)

    out = []
    if repris or not auto:
        out.append(f"{len(repris)} stock(s) repris de {nom} dans stocks.csv" + (" : " + ", ".join(repris) if repris else "."))
    if consos:
        out.append("Consommations corrigées à la main, reprises :\n  " + "\n  ".join(consos))
    if colis:
        out.append("Colisages corrigés, repris :\n  " + "\n  ".join(colis))
    out += msg
    if secus:
        out.append("Sécurité changée sur certaines lignes, non reprise (elle vaut pour tout le cadencier, "
                   "config.json securite_jours) :\n  " + "\n  ".join(secus))
    if perimes:
        out.append("Stocks plus anciens que ceux de stocks.csv, non repris :\n  " + "\n  ".join(perimes))
    if illisibles:
        out.append("Valeurs illisibles, non reprises (écrire un nombre seul, dans l'unité de la ligne) :\n  "
                   + "\n  ".join(illisibles))
    if inconnues:
        out.append("Références absentes de stocks.csv, non reprises (vérifier le fournisseur et la référence) : "
                   + ", ".join(inconnues))
    est_cadencier = any("conso" in x for x in lignes)
    if est_cadencier and not a_origine and not a_estimation:
        out.append("Colonne « Valeurs d'origine » absente (classeur d'une version précédente) : consommations, "
                   "colisages et réglages non repris. Les vérifier avec le client.")
    elif est_cadencier and not a_origine:
        out.append("Classeur de la version 0.7.0 : colisages et réglages non repris.")
    return out


def differe_brut(v, origine):
    try:
        return differe(lire_nombre(v), origine)
    except Illisible:
        return True


def main():
    args = sys.argv[1:]
    defaut = date.today().isoformat()
    if "--date" in args:
        i = args.index("--date")
        if i + 1 >= len(args):
            sys.exit("--date demande une date : --date 08/11/2026")
        try:
            defaut = lire_date(args[i + 1])
        except Illisible:
            sys.exit(f"--date illisible : « {args[i + 1]} ». Écrire 08/11/2026 ou 2026-11-08.")
        del args[i:i + 2]
    if len(args) != 2:
        sys.exit(__doc__)
    ici, chemin = Path(args[0]).expanduser(), Path(args[1]).expanduser()
    if not chemin.exists():
        sys.exit(f"Fichier introuvable : {chemin}")
    try:
        lignes, a_origine = lire_saisies(chemin)
    except Illisible as e:
        sys.exit(str(e))
    for ligne in reprendre(ici, lignes, a_origine, defaut, chemin.name):
        print(ligne)


if __name__ == "__main__":
    main()

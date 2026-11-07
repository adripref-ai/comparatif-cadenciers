"""Outils communs et lecteurs des mises en page déjà connues.

Un lecteur prend le chemin d'un fichier et renvoie (lignes, nombre de lignes annoncées).
Rien n'est inventé : une ligne dont le prix ou la quantité n'est pas certain
reste dans le fichier, avec confiance = « à vérifier ».
"""
import csv
import os
import re
import warnings
from datetime import date, datetime

import fitz  # PyMuPDF
import openpyxl

warnings.filterwarnings("ignore")

COLONNES = ["fournisseur", "document", "date_tarif", "source_date", "fin_validite", "ref",
            "libelle", "marque", "conditionnement", "prix", "unite_prix", "qte", "unite",
            "prix_unitaire", "confiance", "note"]


# ---------- outils communs ----------

def lignes_pdf(page, tol=3):
    """Regroupe les mots d'une page en lignes, d'après leur hauteur sur la page.
    Évite le mélange de colonnes qu'on obtient en lisant le texte brut."""
    mots = sorted(page.get_text("words"), key=lambda m: ((m[1] + m[3]) / 2, m[0]))
    lignes, courante, y = [], [], None
    for m in mots:
        milieu = (m[1] + m[3]) / 2
        if y is None or abs(milieu - y) <= tol:
            courante.append(m)
            y = milieu if y is None else y
        else:
            lignes.append(sorted(courante, key=lambda a: a[0]))
            courante, y = [m], milieu
    if courante:
        lignes.append(sorted(courante, key=lambda a: a[0]))
    return lignes


def nombre(txt):
    return float(str(txt).strip().replace(" ", "").replace(",", "."))


def date_fichier(p):
    """Date interne du PDF si elle existe : celle du fichier change dès qu'on le recopie."""
    if str(p).lower().endswith(".pdf"):
        m = re.match(r"D:(\d{4})(\d{2})(\d{2})", fitz.open(p).metadata.get("creationDate") or "")
        if m:
            return "-".join(m.groups())
    return date.fromtimestamp(os.path.getmtime(p)).isoformat()


def date_fr(txt):
    return datetime.strptime(txt, "%d/%m/%Y").date().isoformat()


NB = r"(\d+(?:[.,]\d+)?)"
FACTEUR = {"kg": (1, "kg"), "k": (1, "kg"), "g": (0.001, "kg"), "gr": (0.001, "kg"),
           "l": (1, "L"), "lit": (1, "L"), "litre": (1, "L"), "litres": (1, "L"),
           "cl": (0.01, "L"), "ml": (0.001, "L")}
MESURE = re.compile(NB + r"\s*(kg|k|gr|g|litres|litre|lit|l|cl|ml)(?![a-wyzà-ÿ]|x(?!\s*\d))", re.I)


def quantite(texte, sans_multiple=False):
    """Trouve la quantité vendue dans un libellé. Renvoie (qte, unité, sûr ?).
    Gère 1K, 500g, 75cl, 60gx12. Avec sans_multiple, le x12 de « 1KGx12 » est le
    nombre d'unités au colis, pas une quantité vendue : on l'ignore."""
    mesures = list(MESURE.finditer(texte))
    if not mesures:
        return None, None, False
    m = mesures[-1]
    f, unite = FACTEUR[m.group(2).lower()]
    qte = nombre(m.group(1)) * f
    apres = texte[m.end():]
    mult = re.match(r"\s*[xX]\s*(\d+)", apres)
    if mult and sans_multiple:
        return round(qte, 4), unite, True
    if mult:  # 60gx12 : 12 pièces de 60 g vendues ensemble
        return round(qte * int(mult.group(1)), 4), unite, False
    return round(qte, 4), unite, len({(x.group(1), x.group(2).lower()) for x in mesures}) == 1


def ligne(fournisseur, doc, date_tarif, source_date, fin, ref, libelle, cond, prix, unite_prix,
          qte=None, unite=None, confiance="sûr", note="", marque=""):
    pu = round(prix / qte, 4) if qte else None
    return dict(fournisseur=fournisseur, document=doc, date_tarif=date_tarif, source_date=source_date,
                fin_validite=fin, ref=str(ref), libelle=libelle.strip(), marque=marque.strip(),
                conditionnement=cond, prix=prix, unite_prix=unite_prix, qte=qte or "",
                unite=unite or "", prix_unitaire=pu if pu is not None else "",
                confiance=confiance, note=note)


# ---------- Sysco : offre tarifaire PDF ----------

SYSCO = re.compile(r"^([A-Z]) (\d{5}) (.+?) (UE|FR|FRI|UN|BE|[A-Z]{2,3}) (\d+ [a-z]{2}\b.*?) "
                   r"(KG|L|PC|ST|BT|CT|BO|BI) (\d+\.\d{2})( X)?$")


def lire_sysco(p):
    doc = fitz.open(p)
    texte = doc[0].get_text()
    debut, fin = re.search(r"valable du (\S+) au (\S+)", texte).groups()
    out, annonces = [], 0
    for page in doc:
        for mots in lignes_pdf(page):
            t = " ".join(m[4] for m in mots)
            if re.match(r"^[A-Z] \d{5} ", t):
                annonces += 1
            m = SYSCO.match(t)
            if not m:
                continue
            etat, ref, lib, _ori, cond, uf, prix, _ = m.groups()
            prix = nombre(prix)
            if etat == "S":  # pastille S de l'offre : un surgelé ne se rapproche pas d'un sec ou d'un frais
                lib += " (surgelé)"
            if uf in ("KG", "L"):
                out.append(ligne("Sysco", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, cond,
                                 prix, "€/" + uf.replace("KG", "kg"), 1, uf.replace("KG", "kg")))
                continue
            # Vendu à la pièce : le poids est dans le conditionnement ou le calibre
            q, u, sur = quantite(cond if MESURE.search(cond) else lib)
            if q:
                out.append(ligne("Sysco", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, cond,
                                 prix, "€/" + uf, q, u, "sûr" if sur else "à vérifier"))
            else:
                out.append(ligne("Sysco", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, cond,
                                 prix, "€/" + uf, 1, uf.lower(), "à vérifier", "quantité absente du conditionnement"))
    return out, annonces


# ---------- Pro à Pro : offre PDF ----------

UV_PAP = r"(SCH|KG|BTE|BTL|PLQ|LOT|PQT|PLT|5/1|4/4|PCE|BRIC|BQT|BID|POT|SAC|CAR|FLA|SEAU|PC|BAC|STK|TUB|ROL|UN|BOC|CT|DOS|BOI)"
PAP = re.compile(r"^(\d{4,6}) " + UV_PAP + r" (.+?) 1\.000 (\d+\.\d{4}) (\d+\.\d{2})$")


def lire_proapro(p):
    doc = fitz.open(p)
    texte = " ".join(page.get_text() for page in doc)
    dates = sorted(re.findall(r"\d{2}/\d{2}/\d{4}", texte)[:2], key=date_fr)
    debut, fin = dates[0], dates[1]
    out, annonces = [], 0
    for page in doc:
        # Colonne Marque séparée par sa position, pour ne pas la coller au libellé
        for mots in lignes_pdf(page, tol=4):
            t = " ".join(m[4] for m in mots)
            if not re.match(r"^\d{4,6} \S+ ", t) or "1.000" not in t:
                continue
            annonces += 1
            m = PAP.match(t)
            if not m:
                continue
            ref, uv, reste, prix, _ = m.groups()
            lib = " ".join(w[4] for w in mots if 104 <= w[0] < 283)
            marque = " ".join(w[4] for w in mots if 283 <= w[0] < 346)
            prix = nombre(prix)
            kw = dict(marque=marque)
            if uv == "KG":
                out.append(ligne("Pro à Pro", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, uv,
                                 prix, "€/kg", 1, "kg", **kw))
                continue
            q, u, sur = quantite(lib)
            if q:
                out.append(ligne("Pro à Pro", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, uv,
                                 prix, "€/" + uv, q, u, "sûr" if sur else "à vérifier", **kw))
            else:
                out.append(ligne("Pro à Pro", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, uv,
                                 prix, "€/" + uv, 1, uv.lower(), "à vérifier", "quantité absente du libellé", **kw))
    return out, annonces


# ---------- Comptoir du Littoral : cadencier PDF ----------

CPT = re.compile(r"^(?:\d{1,3} )?(\d{1,5}) (.+) (\d+,\d{4})$")


def lire_comptoir(p):
    doc = fitz.open(p)
    out, annonces = [], 0
    d = date_fichier(p)
    for page in doc:
        for mots in lignes_pdf(page):
            t = " ".join(m[4] for m in mots)
            if not re.search(r"\d,\d{4}$", t):
                continue
            annonces += 1
            m = CPT.match(t)
            if not m:
                continue
            ref, lib, prix = m.groups()
            prix = nombre(prix)
            if re.search(r"F[UÛ]T", lib, re.I):
                out.append(ligne("Comptoir du Littoral", p.name, d, "date du fichier", "", ref, lib, "fût",
                                 prix, "€/L", 1, "L", "si facture", "prix au litre supposé, à confirmer sur facture"))
                continue
            fin = re.search(NB + r"\s*L$", lib)
            if fin:
                q = nombre(fin.group(1))
                out.append(ligne("Comptoir du Littoral", p.name, d, "date du fichier", "", ref, lib,
                                 "cubi" if lib.startswith("BIB") else "bouteille", prix, "€/pièce", q, "L",
                                 "si facture", "prix à la pièce supposé, à confirmer sur facture"))
            else:
                q, u, _ = quantite(lib)
                out.append(ligne("Comptoir du Littoral", p.name, d, "date du fichier", "", ref, lib, "",
                                 prix, "€/pièce", q, u, "à vérifier", "quantité déduite du libellé"))
    return out, annonces


# ---------- Sodiboissons : cadencier PDF ----------

SODI = re.compile(r"^(\d+) (.+) (\d*\.?\d+)$")


def lire_sodiboissons(p):
    doc = fitz.open(p)
    out, annonces = [], 0
    d = date_fichier(p)
    for page in doc:
        lignes = lignes_pdf(page)
        # Les deux premières colonnes portent le client : on lit à partir de la colonne ARTICLE
        x0 = next(w[0] for l in lignes for w in l if w[4] == "ARTICLE") - 2
        for mots in lignes:
            if not mots[0][4].isdigit():
                continue
            annonces += 1
            m = SODI.match(" ".join(w[4] for w in mots if w[0] >= x0))
            if not m:
                continue
            ref, lib, prix = m.groups()
            prix = float(prix)
            if lib.startswith(("FUT", "BIB")):
                q = re.search(r"(\d+)L", lib)
                nom = "fût" if lib.startswith("FUT") else "cubi"
                out.append(ligne("Sodiboissons", p.name, d, "date du fichier", "", ref, lib,
                                 f"{nom} {q.group(1)} L" if q else nom, prix, "€/L", 1, "L",
                                 "à vérifier", "prix au litre supposé, à confirmer sur facture"))
                continue
            pack = re.search(r"(\d+)\s*X\s*" + NB + r"\s*CL", lib, re.I)
            if pack:
                q = nombre(pack.group(2)) / 100
                out.append(ligne("Sodiboissons", p.name, d, "date du fichier", "", ref, lib,
                                 f"{pack.group(1)} x {pack.group(2)} cl", prix, "€/pièce", q, "L",
                                 "à vérifier", "prix à la bouteille supposé, à confirmer sur facture"))
            else:
                q, u, _ = quantite(lib)
                out.append(ligne("Sodiboissons", p.name, d, "date du fichier", "", ref, lib, "", prix,
                                 "€/pièce", q, u, "à vérifier", "quantité déduite du libellé"))
    return out, annonces


# ---------- Episaveurs : offre Excel ----------

def lire_episaveurs(p):
    ws = openpyxl.load_workbook(p, data_only=True).active
    out, annonces, entetes, d = [], 0, None, None
    for r in ws.iter_rows(values_only=True):
        v = [c for c in r if c is not None]
        if v and isinstance(v[0], str) and "Offre de prix à compter du" in v[0]:
            d = date_fr(re.search(r"\d{2}/\d{2}/\d{4}", v[0]).group())
        if "Réf article" in v:
            entetes = {c: i for i, c in enumerate(r) if c is not None}
            continue
        if not entetes or r[entetes["Réf article"]] in (None, ""):
            continue
        ref = str(r[entetes["Réf article"]]).strip()
        if not re.fullmatch(r"\d+", ref):
            continue
        annonces += 1
        lib = str(r[entetes["Désignation"]])
        prix = nombre(r[entetes["Prix unitaire"]])
        ut = str(r[entetes["Unité de tarif"]]).strip().upper()
        uv = str(r[entetes["Unité de vente"]]).strip()
        nb = r[entetes["Nb UV/Colis"]]
        cond = f"{uv}, {nb} par colis"
        if ut in ("KG", "L", "LT"):
            u = "kg" if ut == "KG" else "L"
            out.append(ligne("Episaveurs", p.name, d, "document", "", ref, lib, cond, prix, "€/" + u, 1, u))
            continue
        # Prix à l'unité (bidon, sac...) : le x4 de « 5Lx4 » ne compte pas. Prix au colis : il compte.
        q, u, sur = quantite(lib, sans_multiple=ut != "COL")
        if lib.startswith("Sac poub"):  # 50L = contenance du sac, pas une quantité vendue
            q = None
        if q:
            out.append(ligne("Episaveurs", p.name, d, "document", "", ref, lib, cond, prix, "€/" + ut, q, u,
                             "sûr" if sur else "à vérifier"))
        else:
            out.append(ligne("Episaveurs", p.name, d, "document", "", ref, lib, cond, prix, "€/" + ut, 1, ut.lower(),
                             "à vérifier", "quantité absente du libellé"))
    return out, annonces


# ---------- Saisie : lignes relevées par Claude sur un scan, une photo ou une mise en page rare ----------

def lire_saisie(p):
    """CSV : fournisseur,date,ref,libelle,conditionnement,prix,unite_prix,qte,unite.
    Toutes les lignes restent « à vérifier » : personne n'a contrôlé la lecture."""
    out, annonces = [], 0
    for x in csv.DictReader(open(p, encoding="utf-8-sig")):
        annonces += 1
        q = nombre(x["qte"]) if x.get("qte") else None
        out.append(ligne(x["fournisseur"], p.name, x.get("date") or date_fichier(p), "saisie", "", x["ref"], x["libelle"],
                         x.get("conditionnement", ""), nombre(x["prix"]), x["unite_prix"], q, x.get("unite"),
                         "à vérifier", "relevé à la main sur le document"))
    return out, annonces


LECTEURS = {"saisie": lire_saisie, "sysco": lire_sysco, "proapro": lire_proapro, "comptoir": lire_comptoir,
            "sodiboissons": lire_sodiboissons, "episaveurs": lire_episaveurs}

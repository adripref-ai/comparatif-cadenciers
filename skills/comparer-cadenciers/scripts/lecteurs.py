"""Outils communs et lecteurs des mises en page déjà connues.

Un lecteur prend le chemin d'un fichier et renvoie (lignes, nombre de lignes annoncées).
Rien n'est inventé : une ligne dont le prix ou la quantité n'est pas certain
reste dans le fichier, avec confiance = « à vérifier ».
"""
import csv
import os
import re
import unicodedata
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


def date_iso(txt):
    """Accepte 2025-12-08 comme 08/12/2025 et renvoie AAAA-MM-JJ."""
    txt = str(txt).strip()
    for forme in ("%Y-%m-%d", "%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(txt, forme).date().isoformat()
        except ValueError:
            pass
    raise ValueError(f"date illisible : « {txt} ». Écrire 08/12/2025 ou 2025-12-08.")


def cle_ref(ref):
    """0002641 et 2641 désignent le même article : on compare les références sans leurs zéros en tête."""
    return str(ref).strip().lstrip("0") or "0"


# Une facture confirme l'unité du cadencier si son prix hors promotion en reste proche.
# Au-delà, c'est sans doute une autre unité (un carton au lieu d'une bouteille) : rien n'est confirmé.
SEUIL_UNITE = 0.30

# Un résultat solide : un écart de plus de 50 % est souvent une erreur d'unité, et un prix de plus de
# 6 mois n'est plus celui qu'on paiera. Mêmes règles dans le classeur et sur la page.
ECART_MAX = 0.5
AGE_MAX_JOURS = 183


def unite_confirmee(prix_cadencier, prix_facture):
    return prix_facture not in (None, "") and abs(float(prix_facture) - float(prix_cadencier)) <= SEUIL_UNITE * float(prix_cadencier)


# Unités de facture reconnues pour ramener une facture au cadencier. Toute autre unité (colis, carton,
# unité de commande, abréviation inconnue) ne convertit rien : mieux vaut un écart signalé qu'un faux prix.
# Une facture au gramme ou au centilitre est trop rare pour être un indice sûr : elle n'est pas convertie.
UNITES_MESURE = {"KG": "kg", "KGS": "kg", "KGNET": "kg", "KILO": "kg", "KILOS": "kg", "KILOGRAMME": "kg",
                 "KILOGRAMMES": "kg", "L": "L", "LT": "L", "LTR": "L", "LITRE": "L", "LITRES": "L"}
UNITES_PETITES = {"G", "GR", "GRAMME", "GRAMMES", "CL", "ML"}
UNITES_PIECE = {"PC", "PCE", "PIECE", "PIECES"}


def sans_accent(txt):
    return "".join(c for c in unicodedata.normalize("NFD", txt) if unicodedata.category(c) != "Mn")


def facteur_facture(tarif, unite_facture):
    """Nombre d'unités facturées dans une unité de prix du cadencier. Le prix facturé multiplié par ce facteur
    se lit dans l'unité du cadencier. Exemples : 19,46 € le demi-pain de 1,4 kg facturé 15,763 €/kg (facteur
    1,4) ; 6,444 € les 6 bricks facturé 1,074 € la brique (facteur 6). Sans indice sûr, le facteur vaut 1."""
    u = re.sub(r"[\s.]", "", sans_accent(unite_facture or "").upper())
    if not u:
        return 1
    if u in UNITES_PETITES:
        return 1
    if u in UNITES_MESURE:
        # Facture au kg ou au litre : cadencier ramené au kg ou au litre
        try:
            qte = float(tarif["qte"])
        except (TypeError, ValueError):
            return 1
        return qte if qte > 0 and tarif["unite"] == UNITES_MESURE[u] else 1
    # Cadencier vendu par lot de N (« 6 Bricks », « 8 Poches ») et facture à la pièce de ce lot :
    # l'unité facturée doit nommer la pièce (PCE) ou le même objet que le lot (BRIC pour Bricks)
    m = re.fullmatch(r"(\d+)\s+([^\W\d_]+)", tarif["unite_prix"].replace("€/", "").strip())
    if not m or int(m.group(1)) < 2:
        return 1
    objet = sans_accent(m.group(2)).upper()
    if objet in UNITES_MESURE or objet in UNITES_PETITES:
        return 1
    meme_objet = len(u) >= 3 and u.isalpha() and (objet.startswith(u) or u.startswith(objet.rstrip("S")))
    return int(m.group(1)) if u in UNITES_PIECE or meme_objet else 1


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
        precedente, x_lib = None, None
        for mots in lignes_pdf(page):
            t = " ".join(m[4] for m in mots)
            if re.match(r"^[A-Z] \d{5} ", t):
                annonces += 1
            elif precedente is not None and abs(mots[0][0] - x_lib) < 4:
                # Suite de la désignation, en retrait sous le libellé : « 8 % d'acidité », « 44 % de cacao »
                lib = precedente["libelle"]
                fin_lib = " (surgelé)" if lib.endswith(" (surgelé)") else ""
                precedente["libelle"] = lib[:len(lib) - len(fin_lib)] + " " + t + fin_lib
                continue
            precedente = None
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
            else:
                # Vendu à la pièce : le poids est dans le conditionnement ou le calibre
                q, u, sur = quantite(cond if MESURE.search(cond) else lib)
                if q:
                    out.append(ligne("Sysco", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, cond,
                                     prix, "€/" + uf, q, u, "sûr" if sur else "à vérifier"))
                else:
                    out.append(ligne("Sysco", p.name, date_fr(debut), "document", date_fr(fin), ref, lib, cond,
                                     prix, "€/" + uf, 1, uf.lower(), "à vérifier", "quantité absente du conditionnement"))
            precedente, x_lib = out[-1], mots[2][0]
    return out, annonces


# ---------- Pro à Pro : offre PDF ----------

UV_PAP = r"(SCH|KG|BTE|BTL|PLQ|LOT|PQT|PLT|5/1|4/4|PCE|BRIC|BQT|BID|POT|SAC|CAR|FLA|SEAU|PC|BAC|STK|TUB|ROL|UN|BOC|CT|DOS|BOI)"
PAP = re.compile(r"^(\d{4,6}) " + UV_PAP + r" (.+?) \d+\.\d{3} (\d+\.\d{4}) (\d+\.\d{2})$")


def lire_proapro(p):
    doc = fitz.open(p)
    texte = " ".join(page.get_text() for page in doc)
    dates = sorted(re.findall(r"\d{2}/\d{2}/\d{4}", texte)[:2], key=date_fr)
    debut, fin = dates[0], dates[1]
    out, annonces, total = [], 0, None
    for page in doc:
        # Colonne Marque séparée par sa position, pour ne pas la coller au libellé
        for mots in lignes_pdf(page, tol=4):
            t = " ".join(m[4] for m in mots)
            # L'offre annonce elle-même son nombre de lignes : c'est le meilleur garde-fou
            nb = re.search(r"Nb Lignes\s*:\s*(\d+)\b", t)
            if nb:
                total = int(nb.group(1))
            # Compté sans regarder la quantité commandée, pour qu'une ligne mal lue se voie
            if not re.match(r"^\d{4,6} \S+ ", t) or not re.search(r"\d\.\d{2}$", t):
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
    return out, max(annonces, total or 0)


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
    x0 = None
    for page in doc:
        lignes = lignes_pdf(page)
        # Les deux premières colonnes portent le client : on lit à partir de la colonne ARTICLE.
        # Une page sans en-tête (suite, ou page blanche) garde la colonne de la page d'avant.
        x0 = next((w[0] - 2 for l in lignes for w in l if w[4] == "ARTICLE"), x0)
        for mots in lignes:
            if x0 is None:
                break
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
        date_offre = next((re.search(r"\d{2}/\d{2}/\d{4}", c) for c in v if isinstance(c, str)
                           and re.search(r"offre|valable|compter|partir", c, re.I) and re.search(r"\d{2}/\d{2}/\d{4}", c)), None)
        if date_offre and not d:
            d = date_fr(date_offre.group())
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
            out.append(ligne("Episaveurs", p.name, d, "document","", ref, lib, cond, prix, "€/" + u, 1, u))
            continue
        # Prix à l'unité (bidon, sac...) : le x4 de « 5Lx4 » ne compte pas. Prix au colis : il compte.
        q, u, sur = quantite(lib, sans_multiple=ut != "COL")
        if lib.startswith("Sac poub"):  # 50L = contenance du sac, pas une quantité vendue
            q = None
        if q:
            out.append(ligne("Episaveurs", p.name, d, "document","", ref, lib, cond, prix, "€/" + ut, q, u,
                             "sûr" if sur else "à vérifier"))
        else:
            out.append(ligne("Episaveurs", p.name, d, "document","", ref, lib, cond, prix, "€/" + ut, 1, ut.lower(),
                             "à vérifier", "quantité absente du libellé"))
    for x in out:  # Offre sans date reconnue : date du fichier, à confirmer
        if not x["date_tarif"]:
            x["date_tarif"], x["source_date"] = (d, "document") if d else (date_fichier(p), "date du fichier")
    return out, annonces


# ---------- Saisie : lignes relevées par Claude sur un scan, une photo ou une mise en page rare ----------

def lire_saisie(p):
    """CSV : fournisseur,date,ref,libelle,conditionnement,prix,unite_prix,qte,unite.
    Toutes les lignes restent « à vérifier » : personne n'a contrôlé la lecture."""
    out, annonces = [], 0
    for x in csv.DictReader(open(p, encoding="utf-8-sig")):
        annonces += 1
        q = nombre(x["qte"]) if x.get("qte") else None
        out.append(ligne(x["fournisseur"], p.name, date_iso(x["date"]) if x.get("date") else date_fichier(p), "saisie", "", x["ref"], x["libelle"],
                         x.get("conditionnement", ""), nombre(x["prix"]), x["unite_prix"], q, x.get("unite"),
                         "à vérifier", "relevé à la main sur le document"))
    return out, annonces


LECTEURS = {"saisie": lire_saisie, "sysco": lire_sysco, "proapro": lire_proapro, "comptoir": lire_comptoir,
            "sodiboissons": lire_sodiboissons, "episaveurs": lire_episaveurs}

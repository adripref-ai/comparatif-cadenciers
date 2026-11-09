"""Validation des correspondances entre produits par le client.

Claude propose un type pour chaque produit (identique, equivalent, different). Le client tranche dans
l'onglet « À valider » du classeur, dans Excel, LibreOffice ou Google Sheets, avec une liste déroulante.
reprendre_validations.py reporte ensuite ses choix dans appariements.csv, qui reste la seule source.
"""
import openpyxl

ONGLET = "À valider"
# Les mots du client, dans l'ordre de la liste déroulante
CHOIX = {"identique": "Même produit", "equivalent": "Remplaçable en cuisine", "different": "Pas comparable"}
DECISIONS = {v.lower(): k for k, v in CHOIX.items()}
# Libellés de la colonne Type du Comparatif
LIBELLE_TYPE = {"identique": "identique", "equivalent": "équivalent", "different": "non comparable"}
COL_PRODUIT, COL_DECISION, COL_COMMENTAIRE = "Produit", "Décision du client", "Commentaire du client"


def decision_type(texte):
    """Type interne d'une décision du client, None si la case est vide. ValueError si illisible."""
    t = (texte or "").strip()
    if not t:
        return None
    if t.lower() not in DECISIONS:
        raise ValueError(t)
    return DECISIONS[t.lower()]


def lire_decisions(classeur):
    """Renvoie {produit: (décision, commentaire)} lus dans l'onglet À valider, ou None sans cet onglet."""
    wb = openpyxl.load_workbook(classeur)
    if ONGLET not in wb.sheetnames:
        return None
    ws = wb[ONGLET]
    for rang in range(1, 30):
        titres = [c.value for c in ws[rang]]
        if COL_PRODUIT in titres and COL_DECISION in titres:
            break
    else:
        return None
    ip, idec, icom = (titres.index(COL_PRODUIT), titres.index(COL_DECISION),
                      titres.index(COL_COMMENTAIRE) if COL_COMMENTAIRE in titres else None)
    out = {}
    for r in ws.iter_rows(min_row=rang + 1, values_only=True):
        if not r[ip]:
            continue
        com = r[icom] if icom is not None else ""
        out[str(r[ip]).strip()] = (str(r[idec] or "").strip(), str(com or "").strip())
    return out


def etat(a):
    """Ce qu'appariements.csv dit d'un produit, d'après sa première ligne : (décision affichée, commentaire)."""
    valide = (a.get("valide") or "").strip().lower() == "oui"
    typ = {"équivalent": "equivalent", "différent": "different"}.get(a["type"].strip().lower(), a["type"].strip().lower())
    return (CHOIX.get(typ, "") if valide else "", (a.get("commentaire") or "").strip())

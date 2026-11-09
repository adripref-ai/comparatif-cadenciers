"""Validation des correspondances entre produits par le client.

Claude propose un type pour chaque produit (identique, equivalent, qualite, different). Le client tranche dans
l'onglet « À valider » du classeur, dans Excel, LibreOffice ou Google Sheets, avec une liste déroulante.
reprendre_validations.py reporte ensuite ses choix dans appariements.csv, qui reste la seule source.
"""
import sys
from pathlib import Path

import openpyxl

ONGLET = "À valider"
# Les mots du client, dans l'ordre de la liste déroulante
# - Même produit : même produit, même conditionnement.
# - Autre conditionnement : même produit, seul le conditionnement change ; comparé au kg, au litre ou à l'unité.
# - Qualité différente : pas tout à fait le même produit (marque, qualité, élevage) ; comparé quand même,
#   mais signalé partout.
# - Ne pas comparer : erreur d'appariement ; affiché pour mémoire, jamais compté.
CHOIX = {"identique": "Même produit", "equivalent": "Autre conditionnement", "qualite": "Qualité différente",
         "different": "Ne pas comparer"}
DECISIONS = {v.lower(): k for k, v in CHOIX.items()}
# Mots des versions 0.9.0 et 0.9.1, encore acceptés dans un classeur rempli avant
DECISIONS.update({"remplaçable en cuisine": "equivalent", "pas comparable": "different"})
# Libellés de la colonne Type du Comparatif
LIBELLE_TYPE = {"identique": "identique", "equivalent": "autre conditionnement", "qualite": "qualité différente",
                "different": "non comparable"}
# Ce qu'on peut écrire dans la colonne type d'appariements.csv
TYPES = {"identique": "identique", "equivalent": "equivalent", "équivalent": "equivalent",
         "qualite": "qualite", "qualité": "qualite", "different": "different", "différent": "different"}
COL_PRODUIT, COL_DECISION, COL_COMMENTAIRE = "Produit", "Décision du client", "Commentaire du client"
# Colonnes cachées : ce que Claude a écrit à la construction. Une case différente est un choix du client ;
# une case identique est laissée telle quelle, même si appariements.csv a changé depuis (reprise depuis
# Google Sheets, remarque orale reportée par Claude).
COL_ECRIT_DEC, COL_ECRIT_COM = "Décision écrite par Claude (ne pas modifier)", "Commentaire écrit par Claude (ne pas modifier)"


def decision_type(texte):
    """Type interne d'une décision du client, None si la case est vide. ValueError si illisible."""
    t = (texte or "").strip()
    if not t:
        return None
    if t.lower() not in DECISIONS:
        raise ValueError(t)
    return DECISIONS[t.lower()]


def normaliser(choix):
    """(décision, commentaire) avec la décision ramenée à son type : « Pas comparable » (0.9.x) vaut « Ne pas comparer »."""
    dec, com = choix
    try:
        return decision_type(dec), com
    except ValueError:
        return dec, com


def lire_decisions(classeur):
    """{produit: (décision, commentaire, écrit)} lus dans l'onglet À valider, ou None sans cet onglet.
    écrit : (décision, commentaire) écrits par Claude à la construction, ou None pour un classeur d'avant 0.11."""
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

    def col(nom):
        return titres.index(nom) if nom in titres else None

    ip, idec, icom, iedec, iecom = map(col, (COL_PRODUIT, COL_DECISION, COL_COMMENTAIRE, COL_ECRIT_DEC, COL_ECRIT_COM))
    texte = lambda r, i: str(r[i] or "").strip() if i is not None else ""
    out = {}
    for r in ws.iter_rows(min_row=rang + 1, values_only=True):
        if not r[ip]:
            continue
        ecrit = (texte(r, iedec), texte(r, iecom)) if iedec is not None else None
        out[str(r[ip]).strip()] = (texte(r, idec), texte(r, icom), ecrit)
    return out


def refuser_si_ouvert(classeur):
    """Excel laisse un fichier ~$<nom> tant que le classeur est ouvert. L'écraser à ce moment-là, c'est risquer
    qu'Excel réenregistre l'ancienne version par-dessus, ou lire une version que le client n'a pas enregistrée."""
    classeur = Path(classeur)
    verrou = next((v for v in classeur.parent.glob("~$*") if classeur.name.endswith(v.name[2:])), None)
    if verrou:
        sys.exit(f"Le classeur {classeur.name} est ouvert dans Excel ({verrou.name}). Demander au client de "
                 "l'enregistrer puis de le fermer, et relancer. S'il est déjà fermé, ce fichier est un reste "
                 "d'Excel : le supprimer.")


def change_par_le_client(dec, com, ecrit):
    return ecrit is None or normaliser((dec, com)) != normaliser(ecrit)


def choix_non_reportes(classeur, reportes):
    """Produits dont le client a changé le choix dans le classeur sans qu'il soit dans appariements.csv.
    reportes : {produit: (décision, commentaire)} d'après appariements.csv. Un produit renommé ou retiré
    depuis la construction compte aussi : son choix serait perdu."""
    oublies = []
    for nom, (dec, com, ecrit) in (lire_decisions(classeur) or {}).items():
        if not change_par_le_client(dec, com, ecrit):
            continue
        if nom not in reportes:
            if ecrit is not None and (dec or com):
                oublies.append(f"{nom} (absent d'appariements.csv : renommé ou retiré ?)")
        elif normaliser((dec, com)) != normaliser(reportes[nom]):
            oublies.append(nom)
    return oublies


def etat(a):
    """Ce qu'appariements.csv dit d'un produit, d'après sa première ligne : (décision affichée, commentaire)."""
    valide = (a.get("valide") or "").strip().lower() == "oui"
    typ = TYPES.get(a["type"].strip().lower(), a["type"].strip().lower())
    return (CHOIX.get(typ, "") if valide else "", (a.get("commentaire") or "").strip())

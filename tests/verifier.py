#!/usr/bin/env python3
"""Rejoue toute la chaîne sur un faux client et vérifie ce que le client verrait.

Usage : python3 tests/verifier.py [--excel]
Crée un client fictif dans un dossier temporaire, avec des dates calculées à partir d'aujourd'hui,
lance les scripts du plugin et vérifie les résultats et les refus. --excel fait aussi calculer le classeur
par Excel ou LibreOffice, puis vérifie que le classeur et la page disent la même chose.
Affiche OK ou la liste des échecs ; le code de sortie vaut 1 en cas d'échec.
"""
import csv
import json
import re
import shutil
import subprocess
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

import openpyxl

SCRIPTS = Path(__file__).resolve().parent.parent / "skills/comparer-cadenciers/scripts"
CLASSEUR = "Comparatif-fournisseurs.xlsx"
echecs = []


def verifier(condition, message):
    if not condition:
        echecs.append(message)


def lancer(script, *args, attendu=0):
    r = subprocess.run([sys.executable, str(SCRIPTS / script), *map(str, args)], capture_output=True, text=True)
    sortie = r.stdout + r.stderr
    if (r.returncode == 0) != (attendu == 0):
        echecs.append(f"{script} {' '.join(map(str, args[1:]))} : code {r.returncode}, attendu {attendu}\n{sortie}")
    return sortie


def jour(n):
    return (date.today() - timedelta(days=n)).isoformat()


LECTEURS_CLIENT = '''import csv
from lecteurs import ligne, nombre


def lire_essai(p):
    out = []
    for x in csv.DictReader(open(p, encoding="utf-8")):
        out.append(ligne(x["fournisseur"], p.name, x["date"], "indiquée", "", x["ref"], x["libelle"], "",
                         nombre(x["prix"]), x["unite_prix"], nombre(x["qte"]), x["unite"], "sûr", ""))
    return out, len(out)


def facture_essai(p):
    if p.suffix != ".csv":
        return None
    lignes = list(csv.DictReader(open(p, encoding="utf-8")))
    return lignes[0]["fournisseur"], lignes[0]["date"], [
        (x["ref"], x["libelle"], float(x["prix"]), x["promo"] == "1", None, None, x["unite"], None) for x in lignes]


LECTEURS = {"essai": lire_essai}
FACTURES = [facture_essai]
'''


def ecrire_csv(chemin, lignes):
    with open(chemin, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerows(lignes)


def creer_client(ici):
    (ici / "cadenciers").mkdir(parents=True)
    (ici / "factures").mkdir()
    (ici / "config.json").write_text(json.dumps(dict(client="Essai", factures="factures", provisoire=True, tableur="")))
    (ici / "lecteurs_client.py").write_text(LECTEURS_CLIENT)
    entete = ["fournisseur", "date", "ref", "libelle", "prix", "unite_prix", "qte", "unite"]
    ecrire_csv(ici / "cadenciers/alpha.csv", [entete,
               ["Alpha", jour(10), "001", "Farine T55 25 kg", "20.00", "€/sac", "25", "kg"],
               ["Alpha", jour(10), "002", "Beurre doux 1 kg", "9.00", "€/kg", "1", "kg"],
               ["Alpha", jour(10), "003", "Crème 35 % 1 L", "4.00", "€/L", "1", "L"],
               ["Alpha", jour(10), "004", "Sel fin 1 kg", "1.00", "€/kg", "1", "kg"],
               ["Alpha", jour(250), "005", "Sucre 1 kg", "1.10", "€/kg", "1", "kg"]])
    ecrire_csv(ici / "cadenciers/beta.csv", [entete,
               ["Beta", jour(5), "10", "Farine T55 10 kg", "9.00", "€/sac", "10", "kg"],
               ["Beta", jour(5), "11", "Beurre doux 500 g", "4.40", "€/plaquette", "0.5", "kg"],
               ["Beta", jour(5), "12", "Crème 30 % 1 L", "3.80", "€/L", "1", "L"],
               ["Beta", jour(5), "13", "Sel de mer 1 kg", "2.00", "€/kg", "1", "kg"],
               ["Beta", jour(5), "14", "Sucre semoule 1 kg", "1.20", "€/kg", "1", "kg"]])
    ecrire_csv(ici / "sources.csv", [["lecteur", "fichier", "date"],
               ["essai", "cadenciers/alpha.csv", ""], ["essai", "cadenciers/beta.csv", ""]])
    ecrire_csv(ici / "factures/alpha-1.csv", [["fournisseur", "date", "ref", "libelle", "prix", "promo", "unite"],
               ["Alpha", jour(3), "1", "FARINE T55 25KG", "20.00", "0", "SAC"]])
    (ici / "factures/scan-inconnu.txt").write_text("facture d'un fournisseur sans lecteur")
    ecrire_csv(ici / "appariements.csv", [["famille", "produit", "unite", "type", "fournisseur", "ref", "qte_forcee", "note"],
               ["Épicerie", "Farine T55", "€/kg", "equivalent", "Alpha", "001", "", "sac de 25 kg contre 10 kg"],
               ["Épicerie", "Farine T55", "€/kg", "equivalent", "Beta", "10", "", "sac de 25 kg contre 10 kg"],
               ["Crèmerie", "Beurre doux", "€/kg", "different", "Alpha", "002", "", ""],
               ["Crèmerie", "Beurre doux", "€/kg", "different", "Beta", "11", "", ""],
               ["Crèmerie", "Crème liquide", "€/L", "qualite", "Alpha", "003", "", "35 % contre 30 %"],
               ["Crèmerie", "Crème liquide", "€/L", "qualite", "Beta", "12", "", "35 % contre 30 %"],
               ["Épicerie", "Sel fin", "€/kg", "identique", "Alpha", "004", "", ""],
               ["Épicerie", "Sel fin", "€/kg", "identique", "Beta", "13", "", ""],
               ["Épicerie", "Sucre", "€/kg", "equivalent", "Alpha", "005", "", ""],
               ["Épicerie", "Sucre", "€/kg", "equivalent", "Beta", "14", "", ""]])


def page(ici):
    html = (ici / "page-comparatif.html").read_text(encoding="utf-8")
    d = json.loads(re.search(r'id="donnees">(.*?)</script>', html, re.S).group(1).replace("<\\/", "</"))
    return {p["nom"]: p for p in d["produits"]}


def decider(classeur, choix):
    wb = openpyxl.load_workbook(classeur)
    ws = wb["À valider"]
    titres = [c.value for c in ws[6]]
    ip, idec = titres.index("Produit") + 1, titres.index("Décision du client") + 1
    for r in range(7, ws.max_row + 1):
        nom = ws.cell(r, ip).value
        if nom in choix:
            ws.cell(r, idec).value = choix[nom] or None
    wb.save(classeur)


def appariements(ici):
    with open(ici / "appariements.csv", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def reconstruire(ici):
    lancer("comparer.py", ici)
    lancer("generer_page.py", ici)
    return page(ici)


def main():
    avec_excel = "--excel" in sys.argv
    ici = Path(tempfile.mkdtemp(prefix="essai-comparatif-")) / "Essai"
    creer_client(ici)
    classeur = ici / CLASSEUR

    lancer("lire_cadenciers.py", ici)
    sortie = lancer("controler_factures.py", ici)
    verifier("scan-inconnu.txt" in sortie and "ATTENTION" in sortie, "facture non reconnue non signalée :\n" + sortie)
    verifier(not (ici / "achats.csv").exists(), "achats.csv ne devrait plus être écrit")
    with open(ici / "controle_factures.csv", encoding="utf-8-sig") as f:
        verifier([x["ref"] for x in csv.DictReader(f)] == ["001"], "la référence contrôlée doit garder ses zéros")

    p = reconstruire(ici)
    wb = openpyxl.load_workbook(classeur)
    verifier(wb.sheetnames == ["Lisez-moi", "À valider", "Comparatif", "Meilleurs prix", "Mercuriale", "Détail",
                               "Mercuriale contre facture", "Évolutions", "Sans équivalent"], f"onglets : {wb.sheetnames}")
    verifier(p["Farine T55"]["gagnant"] == "Alpha" and p["Farine T55"]["solide"], f"Farine : {p['Farine T55']}")
    verifier(p["Beurre doux"]["gagnant"] == "non comparable", "Beurre doux devrait être non comparable")
    verifier(p["Crème liquide"]["type"] == "qualité différente" and p["Crème liquide"]["gagnant"] == "Beta",
             f"Crème : {p['Crème liquide']['type']} {p['Crème liquide']['gagnant']}")
    verifier(p["Sel fin"]["gagnant"] == "Alpha" and not p["Sel fin"]["solide"], "un écart de 100 % ne doit pas être solide")
    verifier(p["Sucre"]["ancien"] and not p["Sucre"]["solide"], "un prix de plus de 6 mois ne doit pas être solide")
    formule = next(c.value for row in wb["Comparatif"].iter_rows() for c in row
                   if isinstance(c.value, str) and c.value.startswith("=") and "À valider" in c.value)
    verifier("MATCH(" in formule, "le Type doit chercher la décision par le nom du produit : " + formule[:120])

    # Un choix fait dans le classeur bloque la page et la reconstruction tant qu'il n'est pas reporté
    decider(classeur, {"Crème liquide": "Ne pas comparer"})
    lancer("generer_page.py", ici, attendu=1)
    lancer("comparer.py", ici, attendu=1)
    lancer("reprendre_validations.py", ici)
    verifier(appariements(ici)[4]["type"] == "different" and appariements(ici)[4]["valide"] == "oui", "choix non reporté")
    p = reconstruire(ici)
    verifier(p["Crème liquide"]["gagnant"] == "non comparable", "Crème devrait être non comparable après report")

    # Une case vidée rend le produit à la proposition de Claude
    decider(classeur, {"Crème liquide": ""})
    lancer("reprendre_validations.py", ici)
    p = reconstruire(ici)
    verifier(p["Crème liquide"]["type"] == "qualité différente", "case vidée : la proposition devrait revenir")

    # Une remarque orale reportée dans appariements.csv ne bloque pas la reconstruction
    lignes = appariements(ici)
    for a in lignes:
        if a["produit"] == "Farine T55":
            a.update(type="qualite", valide="oui", commentaire="le client préfère la farine Alpha")
    with open(ici / "appariements.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, list(lignes[0]))
        w.writeheader()
        w.writerows(lignes)
    p = reconstruire(ici)
    verifier(p["Farine T55"]["type"] == "qualité différente", "remarque orale perdue à la reconstruction")

    # Choix repris depuis une feuille Google téléchargée : le classeur local se reconstruit sans perte
    telecharge = ici.parent / "telecharge.xlsx"
    shutil.copy(classeur, telecharge)
    decider(telecharge, {"Sel fin": "Qualité différente"})
    lancer("reprendre_validations.py", ici, telecharge)
    p = reconstruire(ici)
    verifier(p["Sel fin"]["type"] == "qualité différente", "choix Google Sheets perdu")
    verifier(p["Farine T55"]["type"] == "qualité différente", "la reprise Google a effacé la remarque orale")

    # Un produit renommé après un choix du client : refusé, pour ne pas perdre ce choix
    decider(classeur, {"Sucre": "Même produit"})
    lignes = appariements(ici)
    for a in lignes:
        if a["produit"] == "Sucre":
            a["produit"] = "Sucre en poudre"
    with open(ici / "appariements.csv", "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, list(lignes[0]))
        w.writeheader()
        w.writerows(lignes)
    sortie = lancer("comparer.py", ici, attendu=1)
    verifier("renommé" in sortie, "produit renommé non signalé :\n" + sortie)
    decider(classeur, {"Sucre": ""})
    lancer("comparer.py", ici)

    # Un classeur ouvert dans Excel n'est ni reconstruit ni écrasé par le calcul
    verrou = ici / ("~$" + CLASSEUR)
    verrou.write_text("")
    sortie = lancer("comparer.py", ici, attendu=1)
    verifier("ouvert dans Excel" in sortie, "classeur ouvert non signalé :\n" + sortie)
    verifier("PAS_CALCULE" in lancer("calculer_classeur.py", classeur), "calculer_classeur.py a écrasé un classeur ouvert")
    verrou.unlink()

    # Dossier des factures absent : erreur claire
    (ici / "config.json").write_text(json.dumps(dict(client="Essai", factures="introuvable", provisoire=True)))
    verifier("introuvable" in lancer("controler_factures.py", ici, attendu=1), "dossier des factures absent non signalé")
    (ici / "config.json").write_text(json.dumps(dict(client="Essai", factures="factures", provisoire=True, tableur="")))

    if avec_excel:
        p = reconstruire(ici)
        reponse = lancer("calculer_classeur.py", classeur)
        if "CALCULE" in reponse and "PAS_" not in reponse:
            wb = openpyxl.load_workbook(classeur, data_only=True)
            ws = wb["Comparatif"]
            rang = next(r for r in range(1, 40) if ws.cell(r, 1).value == "Produit")
            titres = [c.value for c in ws[rang]]
            for row in ws.iter_rows(min_row=rang + 1, values_only=True):
                if not row[0]:
                    continue
                x = dict(zip(titres, row))
                attendu = p[row[0]]
                verifier(x["Moins cher"] == attendu["gagnant"], f"{row[0]} : classeur {x['Moins cher']}, page {attendu['gagnant']}")
                verifier((x["Résultat solide"] == "oui") == attendu["solide"], f"{row[0]} : solide différent entre classeur et page")
                verifier(x["Type"] == attendu["type"], f"{row[0]} : type différent entre classeur et page")
        else:
            print("Calcul impossible ici, comparaison classeur et page sautée : " + reponse.strip())

    if echecs:
        print(f"{len(echecs)} ÉCHEC(S) (dossier : {ici})")
        print("\n---\n".join(echecs))
        sys.exit(1)
    shutil.rmtree(ici.parent)
    print("OK")


if __name__ == "__main__":
    main()

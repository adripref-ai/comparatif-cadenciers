#!/usr/bin/env python3
"""Fait calculer les formules du classeur, pour que ses valeurs soient lisibles partout.

Usage : calculer_classeur.py <classeur.xlsx>
Utilise Microsoft Excel sur Mac, sinon LibreOffice. Sans l'un ni l'autre, ou si le calcul échoue, ce n'est
pas bloquant : les formules se calculeront à la première ouverture, et la page interactive n'en dépend pas.
Réponses possibles : CALCULE, PAS_CALCULE (avec la raison) ou PAS_DE_TABLEUR. Le code de sortie est toujours 0.
"""
import os
import shutil
import subprocess
import sys
from pathlib import Path

p = Path(sys.argv[1]).expanduser().resolve()
excel = Path("/Applications/Microsoft Excel.app")
conteneur = Path.home() / "Library/Containers/com.microsoft.Excel/Data"


def excel_mac():
    # Excel est cloisonné : hors de son propre dossier, il demande l'accès dans une fenêtre ou refuse
    # sans rien dire. On fait donc le calcul sur une copie rangée dans son dossier, puis on la rapporte.
    travail = conteneur / "comparatif-cadenciers"
    travail.mkdir(exist_ok=True)
    copie = travail / f"calcul-{os.getpid()}.xlsx"
    shutil.copy2(p, copie)
    script = f'''
tell application "Microsoft Excel"
  open workbook workbook file name "{copie}"
  set wb to active workbook
  if wb is missing value then error "Excel n'a pas ouvert le classeur"
  try
    if name of wb is not "{copie.name}" then error "mauvais classeur actif"
    calculate
    save wb
    close wb saving no
  on error e
    try
      close wb saving no
    end try
    error e
  end try
end tell
return "CALCULE"'''
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=90)
        if r.returncode == 0 and r.stdout.strip() == "CALCULE":
            shutil.copy2(copie, p)
            return "CALCULE"
        return "PAS_CALCULE : " + (r.stderr.strip() or r.stdout.strip() or "Excel n'a rien répondu")
    except subprocess.TimeoutExpired:
        return "PAS_CALCULE : Excel n'a pas répondu en 90 secondes (une fenêtre attend peut-être une réponse)"
    finally:
        copie.unlink(missing_ok=True)


if sys.platform == "darwin" and excel.exists() and conteneur.exists():
    reponse = excel_mac()
elif shutil.which("soffice"):
    try:
        r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                            "--outdir", str(p.parent), str(p)], capture_output=True, text=True, timeout=120)
        reponse = "CALCULE" if r.returncode == 0 else "PAS_CALCULE : " + r.stderr.strip()
    except subprocess.TimeoutExpired:
        reponse = "PAS_CALCULE : LibreOffice n'a pas répondu en 120 secondes"
else:
    reponse = "PAS_DE_TABLEUR"
if reponse != "CALCULE":
    reponse += ". Les formules se calculeront à l'ouverture du fichier."
print(reponse)

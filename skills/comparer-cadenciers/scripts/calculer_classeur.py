#!/usr/bin/env python3
"""Fait calculer les formules du classeur, pour que ses valeurs soient lisibles partout.

Usage : calculer_classeur.py <classeur.xlsx>
Utilise Microsoft Excel sur Mac, sinon LibreOffice. Sans l'un ni l'autre, ce n'est pas bloquant :
les formules se calculeront à la première ouverture, et la page interactive n'en dépend pas.
"""
import shutil
import subprocess
import sys
from pathlib import Path

p = Path(sys.argv[1]).expanduser().resolve()
excel = Path("/Applications/Microsoft Excel.app")
if sys.platform == "darwin" and excel.exists():
    # Le classeur est fermé même en cas d'erreur, pour ne rien laisser ouvert chez l'utilisateur
    script = f'''
tell application "Microsoft Excel"
  open workbook workbook file name "{p}"
  set wb to active workbook
  try
    if name of wb is not "{p.name}" then error "mauvais classeur actif"
    calculate
    save wb
    close wb saving no
  on error e
    close wb saving no
    error e
  end try
end tell
return "CALCULE"'''
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=120)
    print(r.stdout.strip() or r.stderr.strip())
    sys.exit(r.returncode)
if shutil.which("soffice"):
    r = subprocess.run(["soffice", "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                        "--outdir", str(p.parent), str(p)], capture_output=True, text=True, timeout=120)
    print("CALCULE" if r.returncode == 0 else r.stderr.strip())
    sys.exit(r.returncode)
print("PAS_DE_TABLEUR : les formules se calculeront à l'ouverture du fichier.")

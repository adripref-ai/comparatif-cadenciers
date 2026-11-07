#!/usr/bin/env python3
"""Affiche un PDF ligne par ligne, avec la position de chaque ligne, pour écrire un nouveau lecteur.

Usage : apercu_pdf.py <fichier.pdf> [nombre de lignes]
Une page sans aucun mot est blanche ou scannée : la regarder à l'image.
"""
import sys

sys.dont_write_bytecode = True

import fitz

from lecteurs import lignes_pdf

doc = fitz.open(sys.argv[1])
limite = int(sys.argv[2]) if len(sys.argv) > 2 else 60
for i, page in enumerate(doc, 1):
    lignes = lignes_pdf(page)
    print(f"--- page {i} : {len(lignes)} lignes" + ("  (aucun texte : page blanche ou scan)" if not lignes else ""))
    for mots in lignes[:limite]:
        print(f"x={mots[0][0]:>5.0f}  " + " ".join(m[4] for m in mots))

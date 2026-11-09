# Ajouter un cadencier inconnu

Deux possibilités, sans jamais modifier le plugin :

- **Relevé** (scan, photo, moins de 80 lignes, ou document vu une seule fois) : relever les lignes dans un CSV rangé dans `cadenciers/` et utiliser le lecteur `saisie`. Colonnes : `fournisseur,date,ref,libelle,conditionnement,prix,unite_prix,qte,unite`. Sans référence, en inventer une courte et stable (`K1`, `K2`). `unite` vaut `kg`, `L` ou une unité libre, `qte` est la quantité vendue pour ce prix.
- **Lecteur** (document long, ou qui reviendra) : l'écrire dans `lecteurs_client.py`, dans le dossier du client, en suivant la suite de cette page.

Après un relevé, recompter les lignes du document et les comparer à celles du CSV.

## 1. Regarder le fichier

- PDF : `scripts/apercu_pdf.py <fichier>` affiche chaque ligne et sa position. Le texte brut mélange souvent les colonnes, d'où la lecture par lignes.
- Scan (aucun texte) : lire la page comme une image, saisir les lignes dans un CSV et les marquer toutes « à vérifier ».
- Excel : repérer la ligne d'en-tête et lire par nom de colonne.

## 2. Écrire le lecteur

Un lecteur prend un chemin et renvoie `(lignes, annonces)`. `annonces` compte les lignes qui ressemblent à un produit, de façon indépendante de l'analyse détaillée : c'est le garde-fou contre les lignes perdues.

```python
import re
import fitz
from lecteurs import date_fichier, lignes_pdf, ligne, nombre, quantite

MOTIF = re.compile(r"^(\d+) (.+) (\d+,\d{2})$")


def lire_mon_fournisseur(p):
    out, annonces = [], 0
    d = date_fichier(p)
    for page in fitz.open(p):
        for mots in lignes_pdf(page):
            t = " ".join(m[4] for m in mots)
            if not re.search(r"\d,\d{2}$", t):
                continue
            annonces += 1
            m = MOTIF.match(t)
            if not m:
                continue
            ref, lib, prix = m.groups()
            q, u, sur = quantite(lib)
            out.append(ligne("Mon fournisseur", p.name, d, "date du fichier", "", ref, lib, "",
                             nombre(prix), "€/pièce", q, u, "sûr" if sur else "à vérifier"))
    return out, annonces


LECTEURS = {"mon-fournisseur": lire_mon_fournisseur}
```

Marquer « à vérifier » toute unité supposée : prix peut-être à la bouteille ou au colis, au litre ou au fût, ou quantité devinée parmi plusieurs mesures. Une quantité lue sans ambiguïté dans le libellé peut rester « sûr » : c'est la troisième valeur que renvoie `quantite()`. Si `q` est vide, la ligne reste dans `tarifs.csv` sans prix unitaire et devra recevoir une `qte_forcee`.

## 3. Facture d'un format inconnu

Ajouter à `lecteurs_client.py` une fonction qui prend un chemin et renvoie `None` si la facture n'est pas de ce fournisseur, sinon `(fournisseur, "AAAA-MM-JJ", [(ref, libellé, prix, promo, qte, date_livraison, unite, colisage), ...])`. Les quatre derniers champs sont facultatifs : `qte` est la quantité achetée dans l'unité du prix, `date_livraison` la date de la livraison quand une facture en regroupe plusieurs, `colisage` le nombre d'unités par colis. Ignorer les avoirs. Elle reçoit tous les fichiers du dossier des factures, PDF ou non (CSV, Excel, image) : renvoyer `None` pour ceux qu'elle ne sait pas lire.

```python
FACTURES = [lire_facture_mon_fournisseur]
```

Le nom du fournisseur doit être écrit exactement comme dans le lecteur du cadencier.

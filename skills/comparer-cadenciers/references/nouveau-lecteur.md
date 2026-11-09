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

Ajouter à `lecteurs_client.py` une fonction qui prend un chemin et renvoie `None` si la facture n'est pas de ce fournisseur, sinon `(fournisseur, "AAAA-MM-JJ", [(ref, libellé, prix, promo, qte, date_livraison, unite, colisage), ...])`. Seuls les quatre premiers champs et `unite` servent : `qte`, `date_livraison` et `colisage` peuvent valoir `None`. Pour un avoir, renvoyer `(fournisseur, None, [])` : il est reconnu, donc pas signalé comme inconnu, mais rien n'est contrôlé. Un fichier pour lequel tous les lecteurs renvoient `None` est signalé par `controler_factures.py` comme non reconnu.

`unite` est l'unité du prix facturé, recopiée de la facture. Elle sert à ramener la facture à l'unité du cadencier, et seulement dans deux cas :
- au kg ou au litre (`KG`, `KILOS`, `KG NET`, `L`, `LT`, `LITRES`…), face à un tarif dont l'unité de comparaison est le kg ou le litre : 15,763 €/kg contre 19,46 € le demi-pain de 1,4 kg ;
- à la pièce d'un tarif vendu par lot (« €/6 Bricks ») : l'unité doit être `PCE`, `PIECE` ou nommer le même objet que le lot (`BRIC` pour Bricks).

Toute autre unité (`COL`, `CTN`, `UC`, gramme, centilitre, abréviation inconnue) ne convertit rien. Une unité vide non plus. Elle reçoit tous les fichiers du dossier des factures, PDF ou non (CSV, Excel, image) : renvoyer `None` pour ceux qu'elle ne sait pas lire.

```python
FACTURES = [lire_facture_mon_fournisseur]
```

Le nom du fournisseur doit être écrit exactement comme dans le lecteur du cadencier.

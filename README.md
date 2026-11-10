# Comparatif cadenciers

Compare les prix des fournisseurs d'un restaurant à partir de leurs cadenciers.

## Ce qu'il fait

1. Lit les cadenciers PDF ou Excel, et compte les lignes pour n'en perdre aucune.
2. Ramène les prix au kg, au litre ou à l'unité.
3. Contrôle les prix et les unités sur les factures.
4. Met en face les produits équivalents. Le client valide lui-même chaque correspondance dans l'onglet À valider du classeur, avec une liste déroulante : même produit, autre conditionnement (comparé au kg, au litre ou à l'unité), qualité différente (comparé mais signalé) ou ne pas comparer. Le bilan suit tout de suite.
5. Produit un classeur, qui s'ouvre dans Excel, LibreOffice ou Google Sheets : mode d'emploi, correspondances à valider, comparatif, meilleurs prix (le fournisseur le moins cher de chaque produit, avec son conditionnement), mercuriale, détail, mercuriale contre facture, historique des prix, références sans équivalent.
6. Tient une mercuriale : tous les fournisseurs et tous leurs tarifs dans un seul tableau, avec le dernier prix payé.
7. Met la mercuriale face aux factures : pour chaque référence achetée, le tarif du fournisseur face au dernier prix payé.
8. Garde l'historique des prix : chaque prix connu d'une référence, cadenciers et factures, dans l'ordre des dates, avec la hausse ou la baisse depuis le premier prix, au kg ou au litre. Plus on ajoute de cadenciers et de factures au fil des mois, plus il est long.
9. Publie une page interactive avec les mêmes chiffres : une couleur par fournisseur, barres de prix, filtres et recherche.

## Installer

Dans un terminal, sur l'ordinateur à équiper :

```bash
claude plugin marketplace add adripref-ai/comparatif-cadenciers
```

```bash
claude plugin install comparatif-cadenciers@outils-restaurant
```

Ouvrir ensuite une nouvelle conversation.

## Comment s'en servir

Joindre les cadenciers à la conversation ou indiquer leur dossier, puis dire « compare ces cadenciers », ou taper `/comparatif-cadenciers:comparer-cadenciers`. Installer le plugin ne lance rien tout seul.

Chaque client a son propre dossier : aucune donnée de client n'est gardée dans le plugin.

## Ce qu'il faut

- Python 3 avec PyMuPDF et openpyxl.
- Un tableur au choix :
  - Microsoft Excel ou LibreOffice, qui calculent aussi le classeur à l'avance. Sur Mac, le calcul se fait sur une copie rangée dans le dossier d'Excel, sans fenêtre d'autorisation.
  - Google Sheets, pour un client sans Excel : Claude met le classeur dans son Google Drive, avec Chrome ou le connecteur Google Drive, ou lui dit comment l'importer.
- Sans tableur, la page interactive fonctionne quand même.

## Limites

- Lecture seule : aucune commande, aucun envoi.
- Un cadencier est une photo à une date. Le franco, les remises et les frais de livraison n'y figurent pas.
- Cadenciers lus sans rien ajouter : Sysco, Pro à Pro, Episaveurs, Comptoir du Littoral et Sodiboissons. Factures lues sans rien ajouter : Sysco, Pro à Pro, Episaveurs et Comptoir du Littoral. Une autre mise en page demande un lecteur, que Claude écrit dans le dossier du client. Une facture non reconnue est signalée, jamais ignorée en silence.
- Fermer le classeur dans Excel avant que Claude le reconstruise : il refuse de travailler sur un classeur ouvert.

## Vérifier le plugin

Après une modification, lancer le test, qui rejoue toute la chaîne sur un faux client :

```bash
python3 tests/verifier.py --excel
```

Il affiche `OK`, ou la liste de ce qui ne va pas. Sans `--excel`, il ne fait pas calculer le classeur.

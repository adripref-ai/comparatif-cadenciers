# Comparatif cadenciers

Compare les prix des fournisseurs d'un restaurant à partir de leurs cadenciers.

## Ce qu'il fait

1. Lit les cadenciers PDF ou Excel, et compte les lignes pour n'en perdre aucune.
2. Ramène les prix au kg, au litre ou à l'unité.
3. Contrôle les prix et les unités sur les factures.
4. Met en face les produits équivalents, avec validation par une personne du métier.
5. Produit un classeur Excel : comparatif, mercuriale, cadencier de commande, détail, cadencier contre facture, évolutions.
6. Tient une mercuriale : tous les fournisseurs et tous leurs tarifs dans un seul tableau, avec le dernier prix payé.
7. Tient un cadencier de commande : consommations estimées sur les factures, jours de livraison, stocks, et quantités à commander arrondies au colis.
8. Publie une page interactive avec les mêmes chiffres : une couleur par fournisseur, barres de prix, filtres, recherche, et saisie des stocks pour voir tout de suite quoi commander.

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
- Microsoft Excel ou LibreOffice pour calculer le classeur à l'avance. Sur Mac, le calcul se fait sur une copie rangée dans le dossier d'Excel, sans fenêtre d'autorisation. Sans tableur, le classeur se calcule à la première ouverture et la page interactive fonctionne quand même.

## Limites

- Lecture seule : aucune commande, aucun envoi. Le cadencier propose des quantités, le client décide.
- Les consommations sont estimées sur les achats : plus il y a de factures, plus elles sont justes. Elles se corrigent à la main.
- Un cadencier est une photo à une date. Le franco, les remises et les frais de livraison n'y figurent pas.
- Les mises en page connues viennent de Sysco, Pro à Pro, Episaveurs, Comptoir du Littoral et Sodiboissons. Une nouvelle mise en page demande un lecteur, que Claude écrit dans le dossier du client.

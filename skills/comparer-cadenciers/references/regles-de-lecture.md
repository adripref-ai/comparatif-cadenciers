# Règles de lecture

Pièges rencontrés sur de vrais cadenciers. Relire chaque ligne de `tarifs.csv` avec cette liste.

## Unité du prix

- Prix affiché à la bouteille alors que le libellé indique le colis : « EAU 20X50CL » à moins de 1 € est le prix d'une bouteille de 50 cl.
- Fûts et cubis parfois au litre, parfois au contenant : un cubi 10 L à moins de 5 € est au litre, à plus de 15 € il est au cubi. Comparer aux factures ou à un autre fournisseur.
- Sysco : colonne UF. KG et L donnent déjà le prix unitaire, ST, PC, BT et CT donnent un prix à la pièce.
- Pro à Pro : prix par UV. UV `KG` : prix au kg, même si le libellé indique le poids de la pièce.

## Quantité dans le libellé

- `1KGx12` ou `5Lx4` chez Episaveurs : le multiple est le nombre d'unités au colis. Il ne compte que si le prix est au colis (`COL`).
- `60gx12` chez Pro à Pro : 12 pièces vendues ensemble, le multiple compte.
- `5x5x25 1K` : les dimensions ne sont pas une quantité.
- `Sac poubelle 50L` : une contenance, pas une quantité.
- Boîtes 4/4 et 5/1 : pas de poids net, comparer à la boîte.

## Prix

- « Prix brut » : le prix payé peut être plus bas. Contrôler sur facture.
- Une offre peut être valable à durée limitée : garder ses dates.
- Une facture avec promotion ne vaut pas prix normal : le classeur retient le dernier prix hors promotion.
- Chez Sysco, un prix barré est une référence catalogue, pas le prix normal du client.

## État du produit

- Sysco : la pastille S marque un surgelé, le lecteur ajoute « (surgelé) » au libellé. Ne jamais le rapprocher d'un produit sec ou frais.

## Dates

- Un document sans date prend sa date interne de PDF, ou à défaut celle du fichier. Demander la vraie date et la mettre dans `sources.csv`.
- Une offre dont la fin de validité est passée s'affiche « Offre expirée ». Demander un tarif à jour.
- Plus de 3 mois d'écart entre deux prix comparés : résultat indicatif.

## Signaux d'erreur

- Écart supérieur à 50 % : d'abord suspecter une unité ou deux produits différents.
- Prix au kg incohérent avec le métier, comme du sel à 0,04 €/kg.
- Une facture confirme un prix, jamais une quantité.

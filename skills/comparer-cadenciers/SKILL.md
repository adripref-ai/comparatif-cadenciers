---
name: comparer-cadenciers
description: Compare les prix des fournisseurs d'un restaurant ou d'un commerce à partir de leurs cadenciers (PDF ou Excel), ramène tout au kg ou au litre, contrôle sur les factures et produit un classeur Excel et une page interactive en couleurs. À utiliser quand l'utilisateur dit « compare mes cadenciers », « comparatif fournisseurs », « quel fournisseur est le moins cher », « compare ces tarifs », « ajoute ce cadencier », « mets à jour le comparatif », ou fournit plusieurs tarifs de fournisseurs à mettre en face.
---

# Comparer des cadenciers

Produire un comparatif de prix fiable à partir de fichiers de tarifs. Répondre en français, en langage simple. Ne passer aucune commande et ne rien envoyer à un fournisseur.

## Déroulé

L'utilisateur est souvent chez son client, avec lui. Enchaîner les étapes 0 à 5 sans demander de validation intermédiaire, jusqu'à un premier résultat marqué provisoire, puis passer en revue les doutes avec le client et republier. Ne poser une question avant le résultat que si elle bloque vraiment. Dire en une ligne où en est l'analyse à chaque étape.

Traiter tous les fichiers fournis, même d'un format inconnu. Un fichier illisible ne doit pas arrêter les autres : le signaler et continuer. Un document sans prix (catalogue, plaquette) va dans `autres/` et se signale au compte rendu, avec le tarif à demander.

Les scripts sont dans `scripts/`, à côté de ce fichier. Lancer chacun avec le dossier du client en argument.

## Étape 0 · Préparer

1. Vérifier les outils : `python3 -c "import fitz, openpyxl"`. S'ils manquent, préfixer les commandes par `uv run --no-project --with pymupdf --with openpyxl`, ou à défaut les installer avec pip.
2. Créer un dossier pour le client, hors du plugin : dans le dossier de travail s'il y en a un, sinon `~/Documents/Comparatif <client>/`. Y copier le contenu de `modele/` et créer `cadenciers/` et `factures/`.
3. Y copier les cadenciers et les factures, qu'ils soient joints à la conversation ou dans un dossier indiqué. Si un fichier joint n'a pas de chemin sur le disque, demander de le déposer dans `cadenciers/`. Ranger chaque fichier d'après son contenu, pas d'après son nom.
4. Remplir `config.json` : `client` (nom affiché, à trouver dans les cadenciers si personne ne l'a dit) et `factures` (dossier des factures, parcouru en entier).

Ne jamais écrire les données d'un client dans le dossier du plugin.

## Étape 1 · Lire les cadenciers

1. Ouvrir chaque fichier et reconnaître sa mise en page. Lecteurs fournis : `sysco` (offre tarifaire), `proapro` (offre), `comptoir`, `sodiboissons`, `episaveurs` (offre Excel), `saisie` (lignes relevées à la main).
2. Mise en page inconnue, scan ou photo : suivre `references/nouveau-lecteur.md`. Un même fournisseur peut changer de mise en page d'un client à l'autre, donc vérifier avant de réutiliser un lecteur.
3. Remplir `sources.csv` : `lecteur,fichier,date`. Renseigner `date` quand le document n'en porte pas et que la date est connue. Sinon le lecteur prend la date interne du PDF : le signaler comme à confirmer.
4. Lancer `scripts/lire_cadenciers.py <dossier>`. Ne pas continuer tant que « annoncées » et « lues » diffèrent.
5. Relire toutes les lignes de `tarifs.csv` avec `references/regles-de-lecture.md`. Corriger le lecteur, jamais le résultat à la main : un lecteur fourni se remplace en donnant le même nom dans `lecteurs_client.py`.

## Étape 2 · Contrôler sur les factures

Lancer `scripts/controler_factures.py <dossier>`, puis lire `controle_factures.csv` :

- un prix identique confirme l'unité du cadencier ;
- un écart signale un prix brut, un cadencier périmé ou une promotion.

Facture d'un format inconnu : voir `references/nouveau-lecteur.md`. Sans facture, garder les unités déduites « à vérifier » et le dire.

## Étape 3 · Apparier les produits

Écrire `appariements.csv`, une ligne par produit et par fournisseur :

`famille,produit,unite,type,fournisseur,ref,qte_forcee,note`

- `unite` : `€/kg`, `€/L`, ou une unité libre comme `€/boîte 4/4`.
- `type` : `identique` (même marque, même format), `equivalent` (substituable en cuisine), `different` (affiché mais jamais compté).
- `qte_forcee` : obligatoire si l'unité n'est pas le kg ou le litre. Mettre 1 pour comparer à la pièce.
- `note` : ce qui distingue les produits (marque, format, pourcentage).

Commencer par ce que le client achète vraiment, d'après les factures, ou à défaut d'après ses cadenciers personnalisés, plutôt que les tarifs généraux. Dans le doute, choisir `different`. Ne jamais rapprocher un surgelé d'un frais, ni un lyophilisé d'un surgelé.

Ne mettre que les produits présents chez au moins deux fournisseurs, avec une seule référence par fournisseur et par produit. Les autres apparaissent d'eux-mêmes dans l'onglet Sans équivalent du classeur. `type` et `note` se lisent sur la première ligne de chaque produit. Si très peu de produits se recoupent, le dire clairement : c'est un résultat, pas une erreur.

Continuer sans attendre. Seule une personne du métier sait si deux produits se remplacent : tant qu'elle n'a pas validé les `equivalent` et `different`, annoncer le comparatif comme provisoire. Présenter cette liste à l'étape 6.

## Étape 4 · Construire le classeur

1. Lancer `scripts/comparer.py <dossier>`. Il refuse une référence absente ou une unité incohérente : corriger `appariements.csv`.
2. Lancer `scripts/calculer_classeur.py <dossier>/Comparatif-fournisseurs.xlsx`. La réponse `PAS_DE_TABLEUR` n'est pas bloquante.
3. Si le classeur a été calculé, vérifier qu'il n'a aucune erreur de formule. Expliquer ou corriger chaque écart supérieur à 50 %.

## Étape 5 · Publier la page interactive

1. Lancer `scripts/generer_page.py <dossier>`. La page est construite à partir du classeur et ne dépend pas d'Excel.
2. Publier `<dossier>/page-comparatif.html` avec l'outil `Artifact`, sans modifier la page. Noter l'adresse obtenue dans `config.json`, clé `artefact`.
3. Pour une mise à jour, republier à cette même adresse : faire d'abord `Artifact` avec `action: "read"` et cette `url`, puis publier avec `url` et sans `icon`.

L'artefact reste privé tant que l'utilisateur ne le partage pas. Il contient les prix négociés du client : ne pas proposer de le rendre public. Si l'outil `Artifact` n'est pas disponible, remettre le fichier HTML, qui s'ouvre dans un navigateur.

## Étape 6 · Rendre compte

Dire simplement :

- le lien de la page et le chemin du classeur ;
- combien de fichiers ont été lus, et ceux qui ne l'ont pas été ;
- combien de lignes et de produits ont été comparés, et combien de références n'ont aucun équivalent chez un autre fournisseur ;
- qui est le moins cher, en rappelant les limites ;
- combien de résultats sont solides : prix à moins de 3 mois d'écart et unités confirmées ;
- les écarts entre cadencier et factures ;
- la liste des correspondances à valider, à passer en revue avec le client.

Corriger `appariements.csv` d'après ses réponses, mettre `provisoire` à `false` dans `config.json` une fois tout validé, puis relancer les étapes 4 et 5, à la même adresse.

Ne pas présenter comme une économie certaine une comparaison de prix d'âges différents. Le franco, les remises de fin d'année et les frais de livraison ne sont pas dans les cadenciers.

## Mettre à jour

Pour un nouveau cadencier, l'ajouter à `sources.csv` sans retirer l'ancien : le plus récent sert au comparatif, et l'onglet Évolutions montre les changements. Relancer les étapes 1, 2, 4 et 5.

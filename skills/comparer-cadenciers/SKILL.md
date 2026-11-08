---
name: comparer-cadenciers
description: Compare les prix des fournisseurs d'un restaurant ou d'un commerce à partir de leurs cadenciers (PDF ou Excel), ramène tout au kg ou au litre, contrôle sur les factures et produit un classeur (Excel, LibreOffice ou Google Sheets) et une page interactive en couleurs, avec une mercuriale de tous les fournisseurs et un cadencier de commande (stocks, consommations, jours de livraison, quantités à commander). À utiliser quand l'utilisateur dit « compare mes cadenciers », « comparatif fournisseurs », « quel fournisseur est le moins cher », « compare ces tarifs », « ajoute ce cadencier », « mets à jour le comparatif », « fais ma mercuriale », « cadencier de commande », « qu'est-ce que je dois commander », « mets à jour mes stocks », « je n'ai pas Excel », « mets-le dans Google Sheets », ou fournit plusieurs tarifs de fournisseurs à mettre en face.
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
5. Demander dès le début : « Tu préfères ouvrir le classeur dans Excel ou dans Google Sheets ? », avec l'outil de question s'il existe, sinon en une phrase. Noter la réponse dans `config.json`, clé `tableur` : `excel` ou `google_sheets`. Continuer sans attendre : la réponse ne sert qu'à l'étape 4. Si le client ne sait pas, ou n'arrive pas à ouvrir Excel, proposer Google Sheets.

Ne jamais écrire les données d'un client dans le dossier du plugin.

## Étape 1 · Lire les cadenciers

1. Ouvrir chaque fichier et reconnaître sa mise en page. Lecteurs fournis : `sysco` (offre tarifaire), `proapro` (offre), `comptoir`, `sodiboissons`, `episaveurs` (offre Excel), `saisie` (lignes relevées à la main).
2. Mise en page inconnue, scan ou photo : suivre `references/nouveau-lecteur.md`. Un même fournisseur peut changer de mise en page d'un client à l'autre, donc vérifier avant de réutiliser un lecteur.
3. Remplir `sources.csv` : `lecteur,fichier,date`. Renseigner `date` (`08/12/2025` ou `2025-12-08`) quand le document n'en porte pas et que la date est connue. Sinon le lecteur prend la date interne du PDF, ou celle du fichier : le script liste ces documents, à signaler comme dates à confirmer.
4. Lancer `scripts/lire_cadenciers.py <dossier>`. Un fichier illisible est marqué `ILLISIBLE` et les autres sont quand même lus. Ne pas continuer tant que « annoncées » et « lues » diffèrent ou qu'un fichier reste illisible : écrire un lecteur ou faire un relevé.
5. Relire toutes les lignes de `tarifs.csv` avec `references/regles-de-lecture.md`. Corriger le lecteur, jamais le résultat à la main : un lecteur fourni se remplace en donnant le même nom dans `lecteurs_client.py`.

## Étape 2 · Contrôler sur les factures

Lancer `scripts/controler_factures.py <dossier>`, puis lire `controle_factures.csv` :

- un prix hors promotion à moins de 30 % du cadencier confirme son unité. Un petit écart signale un prix brut, un cadencier périmé ou une promotion ;
- au-delà de 30 %, la facture est sans doute dans une autre unité (un carton pour une bouteille) : son prix n'est pas retenu, la ligne reste « à vérifier » et la note de lecture le dit.

Le script écrit aussi `achats.csv` : chaque ligne achetée, avec sa quantité et sa date de livraison. Il sert au cadencier de commande.

Facture d'un format inconnu, PDF ou non : voir `references/nouveau-lecteur.md`. Une unité supposée par le lecteur (prix à la bouteille ou au colis, au litre ou au fût) reste « à vérifier » tant qu'aucune facture ne la confirme : le dire. Une quantité lue sans ambiguïté dans le libellé (`1L`, `500g`) est sûre.

## Étape 3 · Apparier les produits

Écrire `appariements.csv` avec le module `csv` de Python, car les libellés et les notes contiennent des virgules. Une ligne par produit et par fournisseur :

`famille,produit,unite,type,fournisseur,ref,qte_forcee,note`

- `unite` : `€/kg`, `€/L`, ou une unité libre comme `€/boîte 4/4`.
- `unite` : la même sur toutes les lignes d'un produit.
- `type` : `identique` (même marque, même format), `equivalent` (substituable en cuisine), `different` (affiché mais jamais compté). Les accents sont acceptés.
- `qte_forcee` : obligatoire si l'unité n'est pas celle du cadencier. Mettre 1 pour comparer à la pièce. `0,75` et `0.75` sont acceptés. Elle ne confirme pas une unité supposée : seule une facture le fait. Seule exception, `1` quand la quantité manquait au cadencier, puisque rien n'est alors déduit.
- `note` : ce qui distingue les produits (marque, format, pourcentage).

Commencer par ce que le client achète vraiment, d'après les factures, ou à défaut d'après ses cadenciers personnalisés, plutôt que les tarifs généraux. Dans le doute, choisir `different`. Ne jamais rapprocher un surgelé d'un frais, ni un lyophilisé d'un surgelé.

Ne mettre que les produits présents chez au moins deux fournisseurs, avec une seule référence par fournisseur et par produit. Les autres apparaissent d'eux-mêmes dans l'onglet Sans équivalent du classeur. `type` et `note` se lisent sur la première ligne de chaque produit. Si très peu de produits se recoupent, le dire clairement : c'est un résultat, pas une erreur. Avec aucun produit commun, le classeur est quand même produit, avec l'onglet Sans équivalent.

Continuer sans attendre. Seule une personne du métier sait si deux produits se remplacent : tant qu'elle n'a pas validé les `equivalent` et `different`, annoncer le comparatif comme provisoire. Présenter cette liste à l'étape 6.

## Étape 3 bis · Préparer la mercuriale et le cadencier de commande

1. Lancer `scripts/preparer_cadencier.py <dossier>`. Il crée ou complète, sans jamais effacer une saisie :
   - `fournisseurs.csv` : jours de livraison et livraisons par semaine déduits des factures, délai de commande (1 jour par défaut), franco. Compléter ce qu'un document donne : Sysco imprime ses « Jours de livraison » (`Ma J V` s'écrit `mar jeu ven`). Un jour suivi de `?` vient d'une seule livraison : le faire confirmer.
   - `stocks.csv` : une ligne par référence achetée. Le client y met son stock et sa date, et peut y imposer une consommation par semaine (`conso_semaine`) quand l'estimation est fausse ou absente, ou un colisage (`colisage`).
2. Écrire `familles.csv` (`fournisseur,ref,famille`) pour les références qui ne sont pas dans `appariements.csv`, avec les mêmes familles. Sinon elles restent « À classer » dans la mercuriale.
3. La consommation par semaine s'estime sur les quantités achetées, rapportées à la période où l'on observe le fournisseur. Elle suppose que le stock ne bouge pas d'une livraison à l'autre : la faire valider. Avec une seule livraison et un rythme inconnu, elle reste à saisir.
4. `securite_jours` dans `config.json` règle les jours de stock de sécurité (2 par défaut).

## Étape 4 · Construire le classeur

1. Lancer `scripts/comparer.py <dossier>`. Il ajoute les onglets Mercuriale (tous les fournisseurs, tous les tarifs, dernier prix payé, verdict du comparatif) et Cadencier de commande (consommation, livraisons, stock, quantité à commander arrondie au colis, montant estimé par fournisseur). Les cases en bleu se corrigent dans le tableur. Il refuse une référence absente, une unité incohérente ou mélangée, deux références d'un même fournisseur pour un produit, un type ou une `qte_forcee` illisibles : corriger `appariements.csv`.
2. Lancer `scripts/calculer_classeur.py <dossier>/Comparatif-fournisseurs.xlsx`. Il répond `CALCULE`, `PAS_CALCULE` avec la raison, ou `PAS_DE_TABLEUR`. Les deux dernières réponses ne sont pas bloquantes. Si `tableur` vaut `google_sheets`, passer à l'étape 4 bis. S'il n'est pas encore choisi, reposer la question avant l'étape 6.
3. Si le classeur a été calculé, vérifier qu'il n'a aucune erreur de formule. Expliquer ou corriger chaque écart supérieur à 50 %.

Le classeur n'utilise que des fonctions reconnues par Excel, LibreOffice et Google Sheets. `comparer.py` affiche `ATTENTION` s'il en trouve une autre : la remplacer dans le script avant de continuer.

## Étape 4 bis · Ouvrir le classeur dans Google Sheets

Quand le client a choisi Google Sheets, ou n'arrive pas à se connecter à Excel. Il lui faut un compte Google. Google Sheets calcule les formules lui-même : `PAS_DE_TABLEUR` n'y change rien.

1. Mettre `Comparatif-fournisseurs.xlsx` dans le Google Drive du client, par le premier moyen disponible :
   - **Chrome** (outils `claude-in-chrome`), le client connecté à son compte Google : ouvrir drive.google.com, Nouveau, Importer un fichier, et donner le classeur avec l'outil d'envoi de fichier ;
   - **connecteur Google Drive** : créer le fichier avec le contenu du classeur en base64, type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, sans désactiver la conversion. Le fichier devient directement une feuille Google ;
   - **sinon, le client le fait** : lui remettre le classeur et lui dire « Sur drive.google.com : Nouveau, Importer un fichier, choisir le classeur ».
2. Ouvrir le fichier puis faire Fichier, Enregistrer au format Google Sheets, sauf s'il a été converti en route. La mise à jour par remplacement n'existe que sur une feuille Google.
3. Vérifier que les onglets Comparatif et Cadencier de commande n'affichent aucune erreur (`#N/A`, `#ERROR!`, `#NAME?`).
4. Noter l'adresse de la feuille dans `config.json`, clé `google_sheets`. Cette clé, comme `artefact` et `tableur`, ne demande pas de reconstruire le classeur.

Ne jamais saisir le mot de passe Google du client : il se connecte lui-même. La feuille contient ses prix négociés : la laisser privée, et ne la partager qu'avec les personnes qu'il désigne.

Pour une mise à jour : reprendre d'abord les saisies de la feuille (voir « Commander ») : sans cela, la remplacer les efface. Reconstruire le classeur, puis, dans la feuille Google : Fichier, Importer, Importer, choisir le nouveau classeur, « Remplacer la feuille de calcul ». L'adresse reste la même.

## Étape 5 · Publier la page interactive

1. Lancer `scripts/generer_page.py <dossier>`. La page est construite à partir du classeur et ne dépend pas d'Excel. Elle refuse un classeur plus ancien que ses données : relancer d'abord l'étape 4.
2. Publier `<dossier>/page-comparatif.html` avec l'outil `Artifact`, sans modifier la page. Noter l'adresse obtenue dans `config.json`, clé `artefact`.
3. Pour une mise à jour, republier à cette même adresse : faire d'abord `Artifact` avec `action: "read"` et cette `url`, puis publier avec `url` et sans `icon`.

L'artefact reste privé tant que l'utilisateur ne le partage pas. Il contient les prix négociés du client : ne pas proposer de le rendre public. Si l'outil `Artifact` n'est pas disponible, remettre le fichier HTML, qui s'ouvre dans un navigateur.

## Étape 6 · Rendre compte

Dire simplement :

- le lien de la page, le chemin du classeur et, s'il y en a une, l'adresse de la feuille Google Sheets ;
- combien de fichiers ont été lus, et ceux qui ne l'ont pas été ;
- combien de lignes et de produits ont été comparés, et combien de références n'ont aucun équivalent chez un autre fournisseur ;
- qui est le moins cher, en rappelant les limites ;
- combien de résultats sont solides : prix à 3 mois d'écart au plus, au dixième de mois près, aucune offre expirée (une offre vaut jusqu'à son dernier jour inclus) et unités confirmées. Le classeur (colonne Résultat solide) et la page appliquent la même règle ;
- les écarts entre cadencier et factures ;
- la liste des correspondances à valider, à passer en revue avec le client.

Corriger `appariements.csv` d'après ses réponses, mettre `provisoire` à `false` dans `config.json` une fois tout validé, puis relancer les étapes 4 et 5, à la même adresse.

Ne pas présenter comme une économie certaine une comparaison de prix d'âges différents. Le franco, les remises de fin d'année et les frais de livraison ne sont pas dans les cadenciers.

Pour le cadencier de commande, dire aussi quelles consommations sont estimées et sur combien de livraisons, quels fournisseurs n'ont pas de jours de livraison connus, et que les stocks sont à saisir. La page permet de saisir un stock et de voir tout de suite la quantité à commander. Ces saisies restent dans le navigateur : le bouton « Copier les stocks » les copie pour `stocks.csv`.

## Commander

Quand le client donne ses stocks (liste, photo d'inventaire, collage depuis la page) : les reporter dans `stocks.csv` avec la date, puis relancer les étapes 4 et 5.

S'il les a saisis dans son tableur ou sur la page, lancer `scripts/reprendre_stocks.py <dossier> <fichier> [--date JJ/MM/AAAA]` avant de reconstruire le classeur :

- Excel ou LibreOffice : inutile, `comparer.py` reprend tout seul le classeur du dossier avant de le refaire ;
- Google Sheets : la feuille téléchargée (Fichier, Télécharger, Microsoft Excel), ou l'onglet Cadencier de commande relu avec le connecteur Google Sheets et recopié en CSV, ligne de titres et colonne cachée comprises ;
- la page : le texte du bouton « Copier les stocks », enregistré en CSV. Il ne contient que les stocks tapés sur la page, avec leur date.

Seules les cases que le client a changées sont reprises : la colonne cachée « Valeurs d'origine » garde ce que le classeur contenait à sa construction.

- Stock, date, consommation et colisage vont dans `stocks.csv` ; livraisons par semaine et délai dans `fournisseurs.csv`, quand les lignes d'un fournisseur disent la même chose.
- Un stock changé sans changer sa date prend celle de `--date`, sinon celle du jour. Un stock plus ancien que celui de `stocks.csv` n'est pas repris.
- Une case vidée n'efface rien. La sécurité, commune à tout le cadencier, est seulement signalée.

Lire tout le compte rendu au client : une valeur illisible (« 12 packs »), une date illisible, une référence inconnue, un stock plus ancien, un réglage différent selon les lignes ou un classeur d'une version précédente y sont signalés. Ces cas se règlent avec lui.

Ensuite, donner par fournisseur les quantités à commander et le montant estimé. Rappeler le prochain jour de livraison et, quand un autre fournisseur est moins cher sur un produit, le signaler sans changer la commande. Ne jamais passer la commande soi-même.

## Mettre à jour

Pour un nouveau cadencier, l'ajouter à `sources.csv` sans retirer l'ancien : le plus récent sert au comparatif, et l'onglet Évolutions montre les changements, au kg ou au litre quand c'est possible. Une référence écrite avec ou sans ses zéros en tête (`0002641`, `2641`) reste la même. Si le client a saisi des stocks dans son tableur, les reprendre d'abord (voir « Commander »). Relancer les étapes 1, 2, 3 bis, 4 et 5, et 4 bis s'il travaille dans Google Sheets. Pour de nouvelles factures, les déposer dans `factures/` : les consommations s'affinent à chaque livraison.

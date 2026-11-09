---
name: comparer-cadenciers
description: Compare les prix des fournisseurs d'un restaurant ou d'un commerce à partir de leurs cadenciers (PDF ou Excel), ramène tout au kg, au litre ou à l'unité, contrôle sur les factures et produit un classeur (Excel, LibreOffice ou Google Sheets) et une page interactive en couleurs, avec une mercuriale de tous les fournisseurs et une mercuriale contre facture (tarif face au dernier prix payé). À utiliser quand l'utilisateur dit « compare mes cadenciers », « comparatif fournisseurs », « quel fournisseur est le moins cher », « compare ces tarifs », « ajoute ce cadencier », « mets à jour le comparatif », « fais ma mercuriale », « mercuriale contre facture », « le client a validé », « reporte ses choix », « j'ai rempli l'onglet À valider », « je n'ai pas Excel », « mets-le dans Google Sheets », ou fournit plusieurs tarifs de fournisseurs à mettre en face.
---

# Comparer des cadenciers

Produire un comparatif de prix fiable à partir de fichiers de tarifs. Répondre en français, en langage simple. Ne passer aucune commande et ne rien envoyer à un fournisseur.

## Déroulé

L'utilisateur est souvent chez son client, avec lui. Enchaîner les étapes 0 à 6 sans demander de validation intermédiaire, jusqu'à un premier résultat marqué provisoire, puis faire valider les correspondances par le client dans le classeur (étape 7) et republier. Ne poser une question avant le résultat que si elle bloque vraiment. Dire en une ligne où en est l'analyse à chaque étape.

Traiter tous les fichiers fournis, même d'un format inconnu. Un fichier illisible ne doit pas arrêter les autres : le signaler et continuer. Un document sans prix (catalogue, plaquette) va dans `autres/` et se signale au compte rendu, avec le tarif à demander.

Les scripts sont dans `scripts/`, à côté de ce fichier. Lancer chacun avec le dossier du client en argument, sauf `calculer_classeur.py`, qui prend le chemin du classeur. Ne jamais modifier le plugin : ce qui est propre à un client va dans son dossier.

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
4. Lancer `scripts/lire_cadenciers.py <dossier>`. Un fichier illisible est marqué `ILLISIBLE` et les autres sont quand même lus. Ne pas continuer tant que « annoncées » et « lues » diffèrent, qu'un fichier reste illisible ou qu'une référence est en double : écrire un lecteur ou faire un relevé.
5. Relire toutes les lignes de `tarifs.csv` avec `references/regles-de-lecture.md`. Corriger le lecteur, jamais le résultat à la main : un lecteur fourni se remplace en donnant le même nom dans `lecteurs_client.py`.

## Étape 2 · Contrôler sur les factures

Lancer `scripts/controler_factures.py <dossier>`, puis lire `controle_factures.csv` :

- un prix hors promotion à moins de 30 % du cadencier confirme son unité. Un petit écart signale un prix brut, un cadencier périmé ou une promotion ;
- au-delà de 30 %, la facture est sans doute dans une autre unité (un carton pour une bouteille) : son prix n'est pas retenu, la ligne reste « à vérifier » et la note de lecture le dit.

Quand la facture dit qu'elle compte au kg ou au litre (le kg pour un demi-pain de 1,4 kg), ou à l'unité d'un lot vendu par N au cadencier (la brique pour un pack de 6), le prix facturé est d'abord ramené à l'unité du cadencier. Les colonnes `facteur`, `prix_cadencier_ramene` et `facteur_hors_promo` le montrent. Sans unité reconnue (voir `references/nouveau-lecteur.md`), rien n'est converti. `comparer.py` refuse un `controle_factures.csv` d'une version précédente : relancer alors cette étape.

Factures lues sans rien ajouter : Sysco, Pro à Pro, Episaveurs et Comptoir du Littoral, en PDF. Le script liste sous `ATTENTION` chaque fichier qu'il ne reconnaît pas : ces factures ne contrôlent rien. Écrire un lecteur pour elles, ou les sortir du dossier, et le dire au compte rendu. Il s'arrête si le dossier des factures n'existe pas. La Mercuriale prend son dernier prix payé dans `controle_factures.csv`, avec la même règle que le Détail.

Facture d'un format inconnu, PDF ou non : voir `references/nouveau-lecteur.md`. Une unité supposée par le lecteur (prix à la bouteille ou au colis, au litre ou au fût) reste « à vérifier » tant qu'aucune facture ne la confirme : le dire. Une quantité lue sans ambiguïté dans le libellé (`1L`, `500g`) est sûre.

## Étape 3 · Apparier les produits

Écrire `appariements.csv` avec le module `csv` de Python, car les libellés et les notes contiennent des virgules. Une ligne par produit et par fournisseur :

`famille,produit,unite,type,fournisseur,ref,qte_forcee,note`

- `unite` : `€/kg`, `€/L`, ou une unité libre comme `€/boîte 4/4`, la même sur toutes les lignes d'un produit.
- `type` : `identique` (même produit, même conditionnement), `equivalent` (même produit, seul le conditionnement change : bouteille ou bidon, 1 L ou 5 L ; comparé au kg, au litre ou à l'unité), `qualite` (pas tout à fait le même produit : autre marque, autre qualité, autre élevage ; comparé quand même, mais signalé en orange), `different` (erreur d'appariement : affiché pour mémoire, jamais compté). Les accents sont acceptés.
- `qte_forcee` : obligatoire si l'unité n'est pas celle du cadencier. Mettre 1 pour comparer à la pièce. `0,75` et `0.75` sont acceptés. Elle ne confirme pas une unité supposée : seule une facture le fait. Seule exception, `1` quand la quantité manquait au cadencier, puisque rien n'est alors déduit.
- `note` : ce qui distingue les produits (marque, format, pourcentage). Le client la lit pour trancher : l'écrire pour lui.
- `proposition`, `valide`, `commentaire` : remplies par `reprendre_validations.py`, ne pas les écrire à la main.

Commencer par ce que le client achète vraiment, d'après les factures, ou à défaut d'après ses cadenciers personnalisés, plutôt que les tarifs généraux. Dans le doute entre deux produits de même nature (des œufs plein air et des œufs de cage, deux vinaigres balsamiques de marques différentes), choisir `qualite` : le client voit quand même la différence de prix. Garder `different` pour deux produits qui ne sont pas de même nature. Ne jamais rapprocher un surgelé d'un frais, ni un lyophilisé d'un surgelé.

Ne mettre que les produits présents chez au moins deux fournisseurs, avec une seule référence par fournisseur et par produit. Les autres apparaissent d'eux-mêmes dans l'onglet Sans équivalent du classeur. `type` et `note` se lisent sur la première ligne de chaque produit. Si très peu de produits se recoupent, le dire clairement : c'est un résultat, pas une erreur. Avec aucun produit commun, le classeur est quand même produit, avec l'onglet Sans équivalent.

Continuer sans attendre. Seule une personne du métier sait si deux produits se valent : tant qu'elle n'a pas validé les `equivalent`, `qualite` et `different`, annoncer le comparatif comme provisoire. Elle tranche à l'étape 7, dans l'onglet À valider du classeur.

## Étape 3 bis · Préparer la mercuriale

Écrire `familles.csv` (`fournisseur,ref,famille`) pour les références qui ne sont pas dans `appariements.csv`, avec les mêmes familles. Sinon elles restent « À classer » dans la mercuriale.

## Étape 4 · Construire le classeur

1. Lancer `scripts/comparer.py <dossier>`. Il ajoute les onglets Meilleurs prix (pour chaque produit comparable, le fournisseur le moins cher, avec sa référence, son conditionnement et son prix, rangés par fournisseur), Mercuriale (tous les fournisseurs, tous les tarifs, dernier prix payé, verdict du comparatif) et Mercuriale contre facture (le tarif face au dernier prix payé, pour chaque référence retrouvée sur une facture). Il refuse une référence absente, une unité incohérente ou mélangée, deux références d'un même fournisseur pour un produit, un type ou une `qte_forcee` illisibles : corriger `appariements.csv`.
2. Lancer `scripts/calculer_classeur.py <dossier>/Comparatif-fournisseurs.xlsx`. Il répond `CALCULE`, `PAS_CALCULE` avec la raison, ou `PAS_DE_TABLEUR`. Les deux dernières réponses ne sont pas bloquantes. Si `tableur` vaut `google_sheets`, passer à l'étape 4 bis. S'il n'est pas encore choisi, reposer la question maintenant.
3. Si le classeur a été calculé, vérifier qu'il n'a aucune erreur de formule. Expliquer ou corriger chaque écart supérieur à 50 %.

Le classeur n'utilise que des fonctions reconnues par Excel, LibreOffice et Google Sheets. `comparer.py` affiche `ATTENTION` s'il en trouve une autre : ne pas modifier le plugin, le signaler à l'utilisateur. Le classeur reste utilisable dans Excel.

## Étape 4 bis · Ouvrir le classeur dans Google Sheets

Quand le client a choisi Google Sheets, ou n'arrive pas à se connecter à Excel. Il lui faut un compte Google. Google Sheets calcule les formules lui-même : `PAS_DE_TABLEUR` n'y change rien.

1. Mettre `Comparatif-fournisseurs.xlsx` dans le Google Drive du client, par le premier moyen disponible :
   - **Chrome** (outils `claude-in-chrome`), le client connecté à son compte Google : ouvrir drive.google.com, Nouveau, Importer un fichier, et donner le classeur avec l'outil d'envoi de fichier ;
   - **connecteur Google Drive** : créer le fichier avec le contenu du classeur en base64, type `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`, sans désactiver la conversion. Le fichier devient directement une feuille Google ;
   - **sinon, le client le fait** : lui remettre le classeur et lui dire « Sur drive.google.com : Nouveau, Importer un fichier, choisir le classeur ».
2. Ouvrir le fichier puis faire Fichier, Enregistrer au format Google Sheets, sauf s'il a été converti en route. La mise à jour par remplacement n'existe que sur une feuille Google.
3. Vérifier que les onglets Comparatif, Meilleurs prix et Mercuriale n'affichent aucune erreur (`#N/A`, `#ERROR!`, `#NAME?`).
4. Noter l'adresse de la feuille dans `config.json`, clé `google_sheets`. Cette clé, comme `artefact` et `tableur`, ne demande pas de reconstruire le classeur.

Ne jamais saisir le mot de passe Google du client : il se connecte lui-même. La feuille contient ses prix négociés : la laisser privée, et ne la partager qu'avec les personnes qu'il désigne.

Pour une mise à jour : le remplacement efface la feuille. D'abord télécharger la feuille et reprendre les choix du client (étape 7.2), et demander s'il a corrigé autre chose à la main, pour le reporter dans ses fichiers. Reconstruire le classeur, puis, dans la feuille Google : Fichier, Importer, Importer, choisir le nouveau classeur, « Remplacer la feuille de calcul ». L'adresse reste la même.

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
- qui est le moins cher, en rappelant les limites, et que l'onglet Meilleurs prix donne la liste de ce qu'il vaut mieux acheter chez chacun ;
- combien de résultats sont solides : prix à 3 mois d'écart au plus entre eux, au dixième de mois près, aucun de plus de 6 mois, écart de 50 % au plus (au-delà, c'est souvent une erreur d'unité), aucune offre expirée (une offre vaut jusqu'à son dernier jour inclus) et unités confirmées. Le classeur (colonne Résultat solide) et la page appliquent la même règle ;
- les écarts de l'onglet Mercuriale contre facture : tarif face au dernier prix payé ;
- combien de correspondances restent à valider, et où : l'onglet À valider du classeur, et la vue À valider de la page.

## Étape 7 · Faire valider les correspondances par le client

Le client tranche lui-même, dans le classeur, pas dans la conversation. L'onglet À valider liste une ligne par produit : le libellé de chaque fournisseur, ce qui les distingue, la proposition de Claude, puis une colonne Décision avec une liste déroulante (Même produit, Autre conditionnement, Qualité différente, Ne pas comparer) et une colonne Commentaire. Les produits à trancher sont en haut, en jaune. Le client peut trier et filtrer cet onglet. Dans le Comparatif et le Détail, il peut filtrer mais pas trier : leurs formules pointent vers des lignes précises. Une case vide garde la proposition de Claude. Le Comparatif et les Meilleurs prix suivent tout de suite : le client voit le bilan changer.

1. Ouvrir le classeur sur l'onglet À valider, dans Excel, LibreOffice ou la feuille Google Sheets, et le passer en revue avec le client. Commencer par les plus gros achats et par les écarts de plus de 50 %. Quand il a fini, il enregistre puis ferme le classeur : les scripts refusent un classeur encore ouvert dans Excel, qui risquerait d'être réenregistré par-dessus la nouvelle version.
2. Quand il a fini, récupérer ses choix :
   - **Excel ou LibreOffice** : enregistrer le classeur, puis lancer `scripts/reprendre_validations.py <dossier>` ;
   - **Google Sheets** : télécharger la feuille au format Excel (Fichier, Télécharger, Microsoft Excel), avec Chrome ou le connecteur Google Drive, ou demander au client de le faire. Puis lancer `scripts/reprendre_validations.py <dossier> <fichier téléchargé>`.
3. Le script ne reprend que les cases que le client a changées : une colonne cachée garde ce que Claude avait écrit. Il reporte ces décisions et ces commentaires dans `appariements.csv`, garde la proposition de Claude dans la colonne `proposition` et liste les choix différents de cette proposition : les relire, car un « Ne pas comparer » sur un produit acheté change le bilan. Dans le compte rendu, citer les produits « Qualité différente » où le moins cher change de qualité. Quand plus rien n'est à trancher, il passe `provisoire` à `false` dans `config.json`.
4. Relancer les étapes 4, 4 bis s'il travaille dans Google Sheets, et 5, à la même adresse, puis refaire le compte rendu de l'étape 6.

`comparer.py` et `generer_page.py` refusent un classeur où le client a changé des choix pas encore reportés, y compris sur un produit renommé ou retiré depuis : lancer d'abord `reprendre_validations.py`, ou supprimer le classeur pour abandonner ces choix. Un produit renommé perd son choix : le reporter à la main sous le nouveau nom. Une remarque donnée à l'oral se reporte dans `appariements.csv` : changer `type`, mettre `valide` à `oui` et la remarque dans `commentaire`. Elle ne bloque pas la reconstruction, puisque le client n'a pas touché la case.

Ne pas présenter comme une économie certaine une comparaison de prix d'âges différents. Le franco, les remises de fin d'année et les frais de livraison ne sont pas dans les cadenciers.

## Mettre à jour

Pour un nouveau cadencier, l'ajouter à `sources.csv` sans retirer l'ancien : le plus récent sert au comparatif, et l'onglet Évolutions montre les changements, au kg ou au litre quand c'est possible. Une référence écrite avec ou sans ses zéros en tête (`0002641`, `2641`) reste la même. Relancer les étapes 1, 2, 3 (apparier les nouvelles références), 3 bis, 4, 4 bis s'il travaille dans Google Sheets, et 5. S'il travaille dans Google Sheets, reprendre d'abord ses choix (étape 7.2). Si de nouveaux produits sont appariés alors que tout était validé, remettre `provisoire` à `true` dans `config.json` et faire valider ces produits (étape 7). Pour de nouvelles factures, les déposer dans `factures/`, puis relancer les étapes 2, 4 et 5 : le dernier prix payé se met à jour.

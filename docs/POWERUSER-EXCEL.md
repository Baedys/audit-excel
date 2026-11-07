# Power-user Excel — inventaire pour Clair

Périmètre limité à Excel de bureau. Les bibliothèques partagées sont recensées selon les contenus que l’on peut insérer et modifier dans un classeur.

Périmètre confirmé par l’utilisateur le 7 novembre 2026 : les automatismes PowerPoint sont traités dans une autre conversation. Les graphiques, formes, icônes et cartes insérés dans Excel restent inclus. Cette passe relève les capacités et les contrats utiles à une reconstruction indépendante. Les fichiers, images, modèles et assemblages de l’éditeur ne sont pas intégrés au projet.

Relevé du 6 novembre 2026 sur **Power-user 2609.1.3.454** dans Excel de bureau. Le guide recense **26 articles Excel**, dont deux historiques. L’assemblage local déclare **13 groupes de ruban**, **92 identifiants de bouton** et **28 fonctions de feuille**. Les identifiants incluent des variantes, boutons par défaut et entrées d’aide. Le classement documentaire distingue **17 compléments** au périmètre Arixcel relevé et **9 recouvrements partiels**. [Guide Excel](https://support.powerusersoftware.com/support/solutions/folders/80000693836) · [Historique éditeur](https://www.powerusersoftwares.com/version-history)

**Testé** indique un résultat observé dans la session locale. **Déclaré** indique une présence dans le binaire ou le guide. **Proposé pour Clair** désigne un choix de reconstruction.

Contre-vérification du 7 novembre : les trois pages du dossier officiel contiennent les mêmes 26 articles et l’article des fonctions contient 28 noms. Le [catalogue structuré des contrôles](poweruser-excel-controls.json) donne les 135 identifiants, leurs groupes, points CIL et actions directement résolues. Il ne transforme pas les variantes d’un bouton en fonctions distinctes. [Guide, page 1](https://support.powerusersoftware.com/support/solutions/folders/80000693836) · [page 2](https://support.powerusersoftware.com/support/solutions/folders/80000693836/page/2) · [page 3](https://support.powerusersoftware.com/support/solutions/folders/80000693836/page/3)

## Essais dans Excel

Cinq commandes distinctes et cinq fonctions de feuille exécutées dans Excel sur deux classeurs fictifs. Le volet Bibliothèque a aussi été ouvert. La première session a été fermée sans enregistrer. Pour le témoin filtré, seule la copie a été enregistrée afin de vérifier les valeurs et la formule dans le fichier Excel.

| Essai | Résultat observé |
| --- | --- |
| Nettoyer les formules | Retire `Probe!` des références de la feuille courante et conserve `Other!`. Sur `=LEN("Probe!B2")+Probe!B3`, retire aussi le qualificatif **dans le texte**. Excel calcule alors 22 au lieu de 28. [Nettoyage](https://support.powerusersoftware.com/support/solutions/articles/80001022178-clean-formulas) |
| Cellules fusionnées | Liste `C2:C3` et `C4:C5`, adresses cliquables, ignorer, dissocier ou dissocier et dupliquer, par zone ou en lot. Après duplication ciblée, `C2` et `C3` contiennent chacun `Nord`, et `C4:C5` reste dans la liste. [Gestionnaire](https://support.powerusersoftware.com/support/solutions/articles/80001022184-merged-cells-manager) |
| Dépivoter | La sélection inclut les en-têtes. Deux produits × trois années donnent six triplets produit, année, montant sur une nouvelle feuille. La première ligne contient déjà les données. [Dépivoter](https://support.powerusersoftware.com/support/solutions/articles/80001021824-unpivot-table) |
| Copier-coller les visibles | Deux sélecteurs. Avec `C1:C6` et la ligne 2 masquée manuellement, la destination saisie est `H8:H12`. Les six valeurs 10, **12**, 16, 7, 8, 9 apparaissent jusqu’à `H13`. Le 12 vient de la ligne masquée. Autre essai, `A1:C6` vers `E1:G6` avec la même ligne masquée. Seules `E1` et `G1` apparaissent renseignées dans la zone visible. Ces témoins concernent le masquage manuel. [Collage visible](https://support.powerusersoftware.com/support/solutions/articles/80001022168-copy-paste-visible-cells-only) |
| Copier-coller avec filtre actif | Filtre `Statut = Oui` sur `A1:C6`, lignes 3 et 5 filtrées. Copie `B2:B6` vers `C2:C6`. `C2`, `C4`, `C6` reçoivent 10, 30, 50. `C3` et `C5` gardent 102 et 104. **`C7` est écrasée par 60**, valeur de `B7`. Les deux cellules sont hors des plages choisies et hors du filtre. Le fichier enregistré confirme ces six valeurs. [Collage visible](https://support.powerusersoftware.com/support/solutions/articles/80001022168-copy-paste-visible-cells-only) |
| Transposer avec lien dynamique | Sélection `B2:C4`, destination proposée `B8:D9`. Le mode dynamique désactive les choix de format et écrit `=TRANSPOSE(B2:C4)` en `B8`. Le déversement donne `10, 20, 30` puis `10, 102, 30`. Il inclut la ligne filtrée 3. Les trois autres modes sont visibles dans la boîte. [Transposer](https://support.powerusersoftware.com/support/solutions/articles/80001041850-transpose-ranges) |
| Fonctions locales | Sur des fonds jaunes avec valeurs 10 et 20, `SumColor` renvoie 30 et `CountColor` renvoie 2. `HasFormulas(F3)` et `IsMerged(C4)` renvoient des booléens vrais. `SheetExists("Probe")` renvoie le **texte** `True`, confirmé par `ESTLOGIQUE=FAUX` et `ESTTEXTE=VRAI`. [Fonctions avancées](https://support.powerusersoftware.com/support/solutions/articles/80001023454-advanced-excel-functions) |
| Bibliothèque | Volet ancré à droite. Onglets Classeurs, Icônes, Images, Graphiques et Cartes. Recherche, sélecteur de source et filtre de catégorie sur Cartes. Galerie de cartes visible. [Bibliothèque Excel](https://support.powerusersoftware.com/support/solutions/articles/80001021415-library-excel-templates) |

## Bibliothèque de classeurs présente sur le poste

**21 modèles** avec aperçu PNG dans le cache Excel de Power-user. La lecture des conteneurs Office Open XML, sans ouvrir les classeurs dans Excel, établit la présence de 19 `.xlsx` et de deux `.xlsm`. Ce relevé porte sur les fichiers locaux, indépendamment de la disponibilité des collections distantes. [Bibliothèque Excel](https://support.powerusersoftware.com/support/solutions/articles/80001021415-library-excel-templates)

| Famille | Modèles repérés |
| --- | --- |
| Pilotage et projets (7) | Tableau de bord Excel, portefeuille et dimensionnement de projets, RACI, gestion des risques, fiche KPI, liste de tâches. |
| Planning (7) | Gantt quotidien, hebdomadaire et mensuel, calendrier annuel, planning des congés, feuilles de temps hebdomadaire et mensuelle. |
| Finance et sensibilité (4) | Bilan, compte de résultat, analyse de sensibilité, analyse boursière. |
| Données et cartes (3) | Tableaux de bord Covid et macroéconomique, carte des comtés américains. |

Les modèles Covid et macroéconomique contiennent chacun un projet VBA et des connexions de données. Le bilan et le compte de résultat contiennent chacun deux parties `externalLink` dans leur paquet Excel. Pour une bibliothèque Clair, prévoir un aperçu avant insertion avec signalement des macros, connexions et liens externes. Utiliser des modèles originaux ou fournis par l’utilisateur.

## Graphiques et contenus dans Excel

- **Galerie de graphiques.** Modèles Office éditables, avec recherche et catégories. Exemples transposables à Clair avec des données fictives originales. Cascade avec écarts, Gantt à jalons, barres divergentes, graphique en pente, haltères, matrices et gaufres. Le Gantt se décline aussi en grille Excel de planification. L’éditeur documente également les modèles d’entreprise au format `.crtx`. [Galerie](https://support.powerusersoftware.com/support/solutions/articles/80001031593-library-charts) · [Gantt](https://support.powerusersoftware.com/support/solutions/articles/80001025374-gantt-charts) · [Modèles graphiques d’entreprise](https://support.powerusersoftware.com/support/solutions/articles/80000962510-configuring-chart-templates-for-your-organization)
- **Générateurs Mekko et Sankey.** La plage source comprend les en-têtes. Dans le binaire Excel, les générateurs groupent les formes et enregistrent la référence de plage avec les métadonnées du groupe. Le menu contextuel appelle explicitement le rafraîchissement. Les déclenchements automatiques restent à vérifier en démo. Les deux générateurs et la galerie constituent des chantiers séparés pour Clair. [Mekko](https://support.powerusersoftware.com/support/solutions/articles/80001021624-mekko-charts) · [Sankey](https://support.powerusersoftware.com/support/solutions/articles/80001021629-sankey-charts)
- **Dimensions de graphiques natifs.** Le binaire relève largeur et hauteur de `ChartArea`, puis les quatre coordonnées de `PlotArea` et de son intérieur. Il applique ces dix mesures au graphique cible. Cas à vérifier sur des axes, titres ou légendes différents. [Dimensions](https://support.powerusersoftware.com/support/solutions/articles/80001022841-copy-paste-chart-dimensions)
- **Modèles personnels.** Le guide décrit une bibliothèque liée à des fichiers existants. Le binaire ouvre le classeur pour fabriquer l’aperçu quand le fichier d’aperçu manque. Pour Clair, afficher le contenu et les propriétés du paquet avant toute ouverture. [Modèles Excel](https://support.powerusersoftware.com/support/solutions/articles/80001021415-library-excel-templates)

Preuve statique de la version 2609.1.3.454. REA `ev_dabec8d850ce9d8c43a5d665239072a713a70e600b8dd01e320d32260305a212`. Générateurs Sankey et Mekko aux RVA 112716 et 119436, commandes contextuelles de rafraîchissement aux RVA 158104 et 153212, dimensions aux RVA 190188 et 190664, aperçu aux RVA 24612.

## Lots proposés pour Clair

Ruban court pour ouvrir les vues. Analyse des formules, constats et aperçus dans le volet Excel ancré à droite. Les neuf recouvrements du guide concernent **Arixcel** et ne désignent pas neuf fonctions déjà livrées dans Clair. Clair fournit le parseur, la navigation avec restauration du masquage, l’édition avec annulation, les couleurs et la copie d’écarts du comparateur, contenus et formats compris. Les six utilitaires U01 de la branche de développement ajoutent largeurs, ajustement explicite, zoom, quadrillage, pointillés d’impression et nouvelle fenêtre. Le développement U02 raccorde huit commandes de formats avec annulation ; leur validation détaillée et leurs limites restent distinctes des préréglages financiers et styles de ligne de Power-user. Les familles Arixcel U03–U06 restent largement ouvertes. [Fonctions Clair](../README.md) · [Registre de parité](PARITY.md)

Le volet de la bibliothèque a été observé à droite dans Excel. L’inspection de `TaskPaneManagerExcel` confirme que Power-user attache le `CustomTaskPane` natif à la fenêtre Excel active et conserve ses modèles de volet par classeur. La bibliothèque, l’explorateur de classeurs, les cartes, les graphiques et les styles de ligne passent par ce mécanisme. Clair crée déjà son volet natif à droite par fenêtre. Les futures vues peuvent partager ce volet avec l’analyse des formules et le comparateur. [Volet natif Clair](<../src/AddInTaskPane.cs>)

Preuve statique dans `PowerUser_Excel.dll`. REA `ev_dabec8d850ce9d8c43a5d665239072a713a70e600b8dd01e320d32260305a212`. Les RVA 59054 et 62384 donnent le rattachement à la fenêtre active et la création du volet. Les RVA 59071 et 61288 identifient le stockage par classeur. Le RVA 60608 initialise les différentes vues. Le comportement sur plusieurs fenêtres d’un même classeur reste à tester.

| Ordre | Lot | Comportement à viser |
| --- | --- | --- |
| 1 | **Qualité de feuille** | Volet avec adresses des fusions, masquages, vrais blancs et plages utilisées. Aller à la cellule, ignorer un constat, prévisualiser une dissociation. Distinguer valeur, formule et format avant duplication. |
| 2 | **Réécriture contrôlée** | Nettoyage par nœud syntaxique, collage exact, conversion des ancrages ligne/colonne et recalcul ciblé. Pour les modifications, afficher avant/après et périmètre, préserver texte littéral, références externes, noms, formules de tableau et déversements. Appliquer avec annulation. |
| 3 | **Préparation de données** | Collage visible avec liste explicite des cellules source et cible. Vérifier les dimensions et les bornes avant toute écriture. Dépivotage avec en-têtes explicites et sortie séparée. Transposition en valeurs, formules, liens ou formule dynamique, avec indication des lignes filtrées incluses. |
| 4 | **Mesures d’audit** | Dans le volet, compter visibles, fusions, formules et distincts, mesurer par couleur et par feuille. Séparer fond d’origine et fond conditionnel affiché. Afficher le périmètre de calcul. |
| 5 | **Confort et restitution** | Formats financiers, styles de ligne, commandes rapides, bibliothèque de modèles, puis graphiques Mekko et Sankey. Réutiliser les capacités existantes de Clair ou d’Arixcel lorsqu’un comportement est déjà couvert. |

Les essais discriminants pour les trois premiers lots incluent filtres et masquages manuels des deux côtés, plusieurs zones, tailles incompatibles, références à des feuilles avec apostrophes, fusion contenant une formule et plage protégée. Le volet latéral sert à examiner les résultats et à choisir les corrections. [Registre Arixcel de Clair](<PARITY.md>)

- **Point de départ dans Clair.** Le comparateur copie déjà les contenus avec références A1 exactes ou recalées et journalise l’annulation. Le collage exact hors comparaison peut réutiliser ce moteur avec les bornes choisies, blancs compris. Le collage visible demande un plan source–cible dédié. La navigation existante traite les lignes filtrées différemment des lignes masquées manuellement. [Copie des écarts](<../src/ExcelComparisonCopy.cs>)
- **Témoin de régression pour le collage visible.** Dans le cas filtré ci-dessus, la cible `C2:C7` doit devenir `[10,102,30,104,50,106]`. `C3` et `C5` sont filtrées. `C7` reste la sentinelle hors plage. En cas de dimensions visibles incompatibles, la commande s’arrête avant la première écriture. Tester séparément colonnes masquées, zones disjointes et source qui chevauche la destination.
- **Témoin de régression pour les vrais blancs.** Seul un blanc réel doit être retiré. Une formule `=""`, un zéro et un espace doivent suivre le déplacement et conserver leur contenu. Vérifier fusion et protection avant l’opération.

Le nettoyage syntaxique des références reste une nouvelle capacité : le parseur existant ne la fournit pas à lui seul. L’ajustement automatique de Power-user après modification d’une cellule reste distinct de la commande explicite AutoFit de Clair. Le développement U03 ajoute la somme copiée, cinq collages spéciaux et la duplication au volet. **Dupliquer → Copie A1 exacte** conserve le texte A1 ; **Conserver les références internes** déplace les liens internes avec le bloc. La source copiée doit appartenir à un classeur ouvert dans la même instance ; le vrai parcours du presse-papiers et certaines grandes structures restent à valider. Ce raccord ne déclare pas une équivalence générale avec les fonctions Power-user. [Preuves U03](CLIPBOARD-EVIDENCE.md)

## Éléments complémentaires déclarés par le binaire

| Commande | Contrat statique ou distinction |
| --- | --- |
| Recalcul de sélection | `formulas.selection.calculate` calcule cellule par cellule. Témoin utile en calcul manuel avec dépendance hors de la sélection. |
| Absolu / relatif | Quatre modes sur ligne et colonne. Appelle la conversion Excel avec un chemin séparé pour les anciennes formules matricielles. |
| Collage exact | Source et destination, mémorisation des formules avant écriture. Le contrôle de dimensions consulte `SpecialCells(12)`, alors que l’écriture parcourt la plage entière. Tester filtres et masquages séparément du collage visible. `WorksheetHelper.PasteExactFormula`, RVA `0x30974` et `0x309C4`. |
| Sélection par propriété | Fond et texte comparés à la première cellule de la sélection, puis balayage de `UsedRange`. Le fond est lu via `Interior.Color`. Gras, italique, souligné et fusions ont aussi leurs sélections. Le chemin Visibles utilise `SpecialCells(12)` sur les cellules de toute la feuille. `WorksheetHelper.Select`, RVA 194588. |
| Compacter les vides | La commande refuse une sélection avec fusion, puis utilise `SpecialCells(4)` et `Range.Delete` pour décaler à gauche ou vers le haut. `DeleteEmptyCellsViewModel.DeleteEmptyCells`, RVA `0xAF4C`. |
| Réafficher la feuille | Le chemin CIL met lignes et colonnes sur visible et sélectionne A1. Effet sur un filtre actif à vérifier. |
| Image liée | `cells.linked_picture.add` délègue à la commande native Excel `Camera`. [Image liée Excel](https://support.microsoft.com/en-us/excel/dynamically-display-the-contents-of-a-cell-or-range-in-a-graphic-object) |
| Déplacements et formats | Échanger positions ou formats, retourner par axe, formats et devises par échelle. Les commandes de bordure reconstruisent les quatre bords et ajoutent les séparations verticales, horizontales ou les deux. Les styles de ligne ont des bibliothèques personnelle et d’entreprise. Le gestionnaire prévoit import `.xlsx`, aperçu, choix du jeu, ouverture du fichier de définition, renommage et suppression. |
| Navigation et fichiers | Explorateur d’onglets, réglages d’édition dans les cellules, renommage local, enregistrement, envoi en pièce jointe, mot de passe à l’ouverture du classeur, raccourcis. |
| Modèle de nouveau classeur | Définir ou retirer un modèle par défaut pour les nouveaux classeurs. Chemin distinct de l’insertion d’un modèle depuis le volet. `ExcelFileTemplatesManager` aux RVA 90828 et 91140. |
| Pointeur | Le ruban utilise `MenuFocusCell` si la version d’Excel le propose. L’autre branche affiche `settings.pointer`. |

Le ruban Excel est construit par `ExcelRibbonCreator`. Le relevé compte **135 identifiants de contrôle** des familles étudiées, dont **92 boutons**, **13 groupes**, **13 boutons fractionnés**, **15 menus** et **2 séparateurs**. La classe `ExcelActionsIds` expose séparément **48 constantes d’action**. Ces deux décomptes ont des périmètres distincts et recouvrent des commandes communes.

### Index par groupe de ruban

Les nombres ci-dessous viennent des dictionnaires `ExcelRibbonControlsBuilder.Initialize*`. Le constructeur de ruban référence 91 des 92 identifiants de bouton. `btnExcelShortcuts` figure dans le dictionnaire d’aide, mais pas dans la méthode `CreateGroupHelp` de cette version. Certains chemins dépendent de la version d’Excel, du contenu chargé ou de l’état de la licence.

| Groupe | Boutons déclarés | Familles observées dans le constructeur |
| --- | ---: | --- |
| Aide | 11 | Centre de commandes, guide, vidéos, blog, compte et réglages. |
| Insertion | 6 | Bibliothèque et cinq points d’entrée, classeurs, icônes, images, graphiques, cartes. |
| Graphiques | 7 | Mekko, Sankey, cascade, Tornado, capture et application des dimensions. |
| Déplacement | 6 | Échange de positions ou de formats, transposition, inversion par lignes ou colonnes. |
| Fusion | 3 | Centrer sur la sélection, gérer les fusions, concaténer. |
| Texte | 9 | Insérer ou retirer des caractères, appliquer cinq modes de casse. |
| Format | 21 | Nombres par échelle, points de pourcentage et de base, devises, dates, bordures, styles de ligne. |
| Cellules | 10 | Sélections par format, visibilité ou fusion, réaffichage, image liée. |
| Nettoyage | 2 | Compacter les cellules vides, nettoyer les formules. |
| Données | 2 | Copier-coller les visibles, dépivoter. |
| Préférences | 3 | Pointeur, ajustement automatique, calendrier. La branche `MenuFocusCell` ajoute un contrôle natif Excel distinct de ces identifiants. |
| Formules | 5 | Ancrages, recalcul, aide des fonctions avancées, édition directe, collage exact. |
| Fichier | 7 | Enregistrer, envoyer, renommer, protéger ou retirer le mot de passe, explorer les classeurs ouverts. |

La boucle du collage visible (`CopyPasteVisibleCellsImplementation`, RVA `0x315F4`) prend pour borne le numéro de ligne absolu de `source.End(xlDown)`. Elle indexe les cellules visibles de la cible sans borner chaque écriture à la plage choisie. Le cas `C7` ci-dessus confirme ce dépassement. Dans la transposition dynamique, le CIL utilise l’adresse de plage sans qualificateur de classeur. Le comportement interclasseurs reste à établir. Le ruban inspecté ne déclare pas de gestionnaire de liens externes dédié. Excel fournit ce gestionnaire dans **Données → Liens de classeur**. [Liens de classeur Excel](https://support.microsoft.com/en-us/excel/manage-workbook-links)

### Fonctions contextuelles Excel hors ruban

Déclarations et chemins statiques du même assemblage Excel. Les interactions suivantes restent à reproduire en démo. Les RVA sont en décimal.

| Fonction | Comportement établi |
| --- | --- |
| Flèches de variation sur les graphiques | Depuis le clic droit d’une série, ajouter ou rafraîchir une flèche. Depuis la flèche, choisir deux catégories et afficher une variation en valeur, en pourcentage ou en CAGR. Le rafraîchissement relit les séries et retire les anciennes formes. `ExcelContextMenusControlsBuilder` aux RVA 146248–146476 et 149088–149164, `ExcelActions` aux RVA 210044–211028. [Historique éditeur](https://www.powerusersoftwares.com/fr/historique-des-versions?comp-m4li6mno_page=2) |
| Sélection fine d’un graphique Power-user | Menus séparés pour cibler séries, données, étiquettes, catégories et légendes selon le type de graphique. Un objet sélectionné peut ainsi être modifié sans sélectionner tout le graphique. [Mekko](https://support.powerusersoftware.com/support/solutions/articles/80001021624-mekko-charts) · [Sankey](https://support.powerusersoftware.com/support/solutions/articles/80001021629-sankey-charts) |
| Cartes liées aux données | Clic droit pour rééditer la carte, rafraîchir sa coloration et sélectionner ses légendes. Le guide précise que la table d’alimentation est ajoutée au classeur lors de l’insertion. `ExcelContextMenusControlsBuilder` aux RVA 148512–148914. [Cartes](https://support.powerusersoftware.com/support/solutions/articles/80001021388-library-maps) |
| Édition contextuelle du Gantt | Fenêtre dédiée, dates de début et de fin, repère Aujourd’hui et mise à jour des catégories multiples dans la feuille de données. `GanttContextualShapeViewModel` aux RVA 45629, 45685 et 45730, `ExcelGanttChart` au RVA 73324. [Gantt](https://support.powerusersoftware.com/support/solutions/articles/80001025374-gantt-charts) |
| Style d’une icône insérée | Le clic droit permet de changer le style d’une icône existante. Le menu couvre les formes, formes libres et groupes. `ExcelContextMenusControlsBuilder` aux RVA 146093, 148978 et 149033. |

Deux actions distinctes côté grille. `cells.swap.format` utilise `xlPasteFormats` pour échanger les formats de deux plages via une zone temporaire (`ExcelActions`, RVA 204208). `objects.swap` passe par des opérations de coupe pour l’échange de plages (RVA 198192). Tester séparément les valeurs, les références de formules et les règles conditionnelles après chaque opération.

## Les 26 articles du guide Excel

Le classement compare les familles du guide au registre Arixcel de Clair. Les comportements précis restent à comparer à l’intérieur des familles communes.

| Nº | Fonction | Comportement documenté dans Excel | Rapport à Arixcel |
| --- | --- | --- | --- |
| 1 | Bibliothèque de classeurs | Modèles de tableaux de bord, feuilles de planning et calculs. Insertion depuis le volet, collections personnelles et partagées par l’organisation, mises à jour. [Bibliothèque](https://support.powerusersoftware.com/support/solutions/articles/80001021415-library-excel-templates) | Complément |
| 2 | Mekko | Graphique depuis matrice avec en-têtes. Aire proportionnelle, options de titre, bordures, totaux, décimales, couleurs et étiquettes par série. Volet à droite et réouverture contextuelle. [Mekko](https://support.powerusersoftware.com/support/solutions/articles/80001021624-mekko-charts) | Complément |
| 3 | Sankey | Matrice origines × destinations, options de graphique/série, couleurs et étiquettes. Le guide illustre un niveau ; les scénarios multiétapes ne sont pas établis par ce relevé. [Sankey](https://support.powerusersoftware.com/support/solutions/articles/80001021629-sankey-charts) | Complément. Flux de valeurs, distinct du graphe de dépendances d’Arixcel |
| 4 | Tornado | Générateur hérité à base de formes. Compare des séries autour d’un axe depuis une plage avec en-têtes. Le guide renvoie vers un modèle plus récent dans la bibliothèque. [Tornado](https://support.powerusersoftware.com/support/solutions/articles/80001021631--legacy-tornado-charts) | Complément historique |
| 5 | Cascade | Générateur hérité à base de formes. Totaux, écarts positifs/négatifs, séries simples ou cumulées, passage sous zéro. Le guide renvoie vers un modèle plus récent dans la bibliothèque. [Cascade](https://support.powerusersoftware.com/support/solutions/articles/80001021633--legacy-waterfall-charts) | Complément historique |
| 6 | Dimensions de graphiques | Capture puis applique dix mesures. Largeur et hauteur du graphique, puis position et taille des zones de tracé externe et interne. [Dimensions](https://support.powerusersoftware.com/support/solutions/articles/80001022841-copy-paste-chart-dimensions) | Complément |
| 7 | Dépivoter | Transforme une matrice avec ses en-têtes en lignes de données, sur une nouvelle feuille. L’exemple demande d’ajouter ensuite les titres des colonnes. [UnPivot](https://support.powerusersoftware.com/support/solutions/articles/80001021824-unpivot-table) | Complément. Préparation de données |
| 8 | Copier/coller les cellules visibles | Deux plages sélectionnées, prise en compte des cellules visibles dans les données filtrées. Le témoin confirme un dépassement de plage. Colonnes masquées, zones multiples et nombres différents restent à tester. [Collage visible](https://support.powerusersoftware.com/support/solutions/articles/80001022168-copy-paste-visible-cells-only) | Voisin d’`PasteSpecial` et `SpecialCells`. Règle de correspondance spécifique |
| 9 | Compacter les vides | Supprime les cellules vides d’une plage et tasse le reste vers le haut ou la gauche. Le guide ne tranche pas les formules produisant une chaîne vide. [Suppression des vides](https://support.powerusersoftware.com/support/solutions/articles/80001022177-delete-empty-cells) | Complément |
| 10 | Nettoyer les formules | Sur la feuille active, retire des références le qualificatif redondant vers la feuille de la formule et garde les références interfeuilles. [Nettoyage](https://support.powerusersoftware.com/support/solutions/articles/80001022178-clean-formulas) | Voisin de `ReplaceFormulas` et `ChangeFormulas`. Règle ciblée distincte |
| 11 | Gérer les cellules fusionnées | Liste de la feuille, navigation à la zone, ignorer, défusionner, défusionner et propager le contenu, actions de masse, alternative Centrer sur la sélection. [Fusions](https://support.powerusersoftware.com/support/solutions/articles/80001022184-merged-cells-manager) | Complément |
| 12 | Fonctions Excel supplémentaires | 28 identifiants publics, détaillés ci-dessous. L’éditeur les inclut dans Free Forever et indique que leur calcul requiert le complément. [Fonctions](https://support.powerusersoftware.com/support/solutions/articles/80001023454-advanced-excel-functions) | Complément. Création de formules au lieu de l’analyse de formules |
| 13 | Absolu / relatif | Modifie les ancrages ligne et colonne des références dans les formules d’une plage. [Ancrages](https://support.powerusersoftware.com/support/solutions/articles/80001024891-absolute-relative) | Voisin de `ChangeFormulas`. Modes d’ancrage à comparer |
| 14 | Coller la formule exacte | Source et cible de même taille. Copie la formule et garde les références relatives telles qu’écrites. [Collage exact](https://support.powerusersoftware.com/support/solutions/articles/80001024892-paste-exact-formula) | Clair propose Dupliquer → Copie A1 exacte ; les critères source/cible et le parcours du presse-papiers restent des frontières distinctes. |
| 15 | Échanger des positions | Échange cellules, plages de même taille, graphiques et formes. Les détails sur l’ajustement des références ne sont pas précisés par le guide. [Échange](https://support.powerusersoftware.com/support/solutions/articles/80001024893-swap-positions) | Complément |
| 16 | Formats numériques | Nombres, pourcentages, unités, devises et dates, formats personnalisés d’organisation. [Formats numériques](https://support.powerusersoftware.com/support/solutions/articles/80001025172-format-numbers) | Voisin de `NumberFormat`, `ConvertNumberFormat` et `DecimalPlaces` |
| 17 | Transposer les plages | Quatre sorties, valeurs, formules ajustées, liens cellule par cellule, formule dynamique TRANSPOSE. Options de format source ou destination. [Transposer](https://support.powerusersoftware.com/support/solutions/articles/80001041850-transpose-ranges) | Voisin du collage. Les quatre modes forment une capacité distincte |
| 18 | Inverser lignes/colonnes | Renverse l’ordre d’un axe sur la sélection. Déplace formules et formats avec les cellules. [Inverser](https://support.powerusersoftware.com/support/solutions/articles/80001041851-flip-range) | Complément |
| 19 | Concaténer des plages | Par ligne, colonne ou plage entière, séparateur défini, option d’ignorer les vides, formule TEXTJOIN ou valeurs. [Concaténer](https://support.powerusersoftware.com/support/solutions/articles/80001041852-concatenate-ranges) | Complément |
| 20 | Changer la casse | Phrase, majuscules, minuscules, initiales de mots, inversion de la casse. [Casse](https://support.powerusersoftware.com/support/solutions/articles/80001043965-change-case) | Voisin de `TextCase` |
| 21 | Insérer du texte | Texte fixe ou plage de valeurs, avant/après un caractère, une position, un mot ou une occurrence de séparateur, en formules ou en valeurs. [Insérer](https://support.powerusersoftware.com/support/solutions/articles/80001044002-add-characters) | Complément |
| 22 | Nettoyer du texte | Retire espaces excédentaires, retours et caractères non imprimables, ou caractères choisis avec option de casse. [Retirer](https://support.powerusersoftware.com/support/solutions/articles/80001043973-remove-characters) | Complément |
| 23 | Pointeur | Surligne les lignes et colonnes de la sélection, couleur réglable, activation/désactivation. L’éditeur signale l’impact sur l’annulation Excel. [Pointeur](https://support.powerusersoftware.com/support/solutions/articles/80001092108-pointer) | Complément. Repère de lecture distinct de la carte de formules |
| 24 | AutoFit automatique | Redimensionne les colonnes après modification d’une cellule, mode à bascule. L’éditeur signale l’impact sur l’annulation Excel. [AutoFit](https://support.powerusersoftware.com/support/solutions/articles/80001092109-autofit) | Voisin d’`AutoFit`. Déclenchement événementiel à comparer |
| 25 | Calendrier | Calendrier de saisie à côté des cellules au format date et accès depuis le clic droit. Navigation année, mois et jour. [Calendrier](https://support.powerusersoftware.com/support/solutions/articles/80001092734-calendar) | Complément |
| 26 | Styles de lignes | Styles différenciés pour textes et nombres, contour de ligne plutôt que de cellule, galerie et jeux personnalisables par fichier. [Styles](https://support.powerusersoftware.com/support/solutions/articles/80001201366-table-styles) | Voisin de `Style`, `Border` et `NumberFormat`. L’application par type et bordure de ligne est spécifique |

Bilan de classement documentaire. **17 compléments**, dont 2 historiques, et **9 voisins**. Ce classement porte sur les catégories ci-dessus. Il ne tranche pas l’égalité du comportement à l’intérieur d’une catégorie.

## Fonctions transversales documentées dans Excel

- **Palettes et raccourcis**. Palette de commandes avec recherche, descriptif, aide et affectation de raccourcis aux commandes de l’extension et d’Excel, disponibilité selon la sélection. Explorateur des classeurs ouverts et actions de fichier. Voisin de la gestion des raccourcis Arixcel et de `NewWindow`. [Command Center](https://support.powerusersoftware.com/support/solutions/articles/80001171221-command-center), [Tab Explorer](https://support.powerusersoftware.com/support/solutions/articles/80001023486-tab-explorer)
- **Bibliothèque multimédia**. Les docs applicables à Excel montrent des modèles de graphiques éditables, du Gantt et des cartes vectorielles, avec possibilité de colorier les zones depuis Excel. L’insertion d’un modèle reste distincte d’un générateur de données. [Graphiques](https://support.powerusersoftware.com/support/solutions/articles/80001031593-library-charts), [Gantt](https://support.powerusersoftware.com/support/solutions/articles/80001025374-gantt-charts), [cartes](https://support.powerusersoftware.com/support/solutions/articles/80001021388-library-maps)
- **Réglages Excel**. Démarrage du volet et des onglets, chargement IntelliSense des fonctions, préférences de raccourci et de bibliothèque. Le catalogue utilise des fichiers et téléchargements distincts. Un contrôle déclaré ne permet pas de déduire la disponibilité du contenu en démo. [Réglages](https://support.powerusersoftware.com/support/solutions/articles/80001023258-settings), [architecture technique](https://support.powerusersoftware.com/support/solutions/articles/80001017162-high-level-technical-documentation), [offres](https://www.powerusersoftwares.com/plans)

Pour la reconstruction Excel, l’explorateur de classeurs peut proposer nom complet, activation, ouverture du dossier, renommage et enregistrement de tous les classeurs. La palette peut rechercher les commandes Clair et Excel, filtrer selon la sélection, proposer l’aide et gérer les conflits de raccourcis. Les cartes peuvent être colorées par gradient ou classes avec noms, valeurs et légende à partir d’une table du classeur. La bibliothèque peut lier des fichiers personnels ou d’équipe avec aperçu de la première feuille ou PNG fourni. Ces détails sont documentés ; leur exécution dans la démo locale reste à tester. [Explorateur](https://support.powerusersoftware.com/support/solutions/articles/80001023486-tab-explorer) · [Palette](https://support.powerusersoftware.com/support/solutions/articles/80001171221-command-center) · [Cartes](https://support.powerusersoftware.com/support/solutions/articles/80001021388-library-maps) · [Classeurs](https://support.powerusersoftware.com/support/solutions/articles/80001021415-library-excel-templates)

## Accès en démo et portabilité

La page des offres consultée le 7 novembre annonce un essai Premium de 90 jours. Free Forever comprend toutes les fonctions de feuille, toutes les images et une partie des autres contenus. La FAQ distingue fin d’essai et fin d’abonnement. Ces annonces ne prouvent ni la date d’expiration ni les droits actuels de l’installation locale. La présente passe ne modifie pas l’état de licence. [Offres](https://www.powerusersoftwares.com/plans)

Le recalcul des UDF Power-user nécessite le complément. Pour Clair, privilégier des mesures visibles dans le volet, ou des formules Excel natives lorsqu’elles suffisent. Si des UDF propres à Clair sont ajoutées, leur dépendance au complément doit être explicite avant insertion. [Fonctions avancées](https://support.powerusersoftware.com/support/solutions/articles/80001023454-advanced-excel-functions)

## Historique et version locale

La lecture de `FileVersion` pour `PowerUser_Excel.dll` installé donne **2609.1.3.454**. Historique éditeur du **21 octobre 2026**, correction du modèle par défaut. Les notes **2609.1.0.445** du **2 octobre 2026** mentionnent la nouvelle interface multi-jeux de styles de tableaux et le calendrier limité aux cellules au format date. Les notes **1.7.87.424** du **10 août 2026** signalent une correction lorsque des lignes ou colonnes entières sont sélectionnées. Les notes **1.7.84.406** du **20 juin 2026** annoncent le renommage des modèles de bibliothèque Excel en classeurs, et **1.7.80.380** du **14 mai 2026** introduit les styles de lignes. Ce sont des changements documentés, pas des observations sur la démo. [Historique](https://www.powerusersoftwares.com/version-history)

## 28 UDF déclarées dans le payload local

Les noms ci-dessous viennent de la propriété `Name` des attributs ExcelFunction localisés. Les paramètres viennent de la méthode managée. Les **28** portent explicitement `IsVolatile=true`, dont **18** avec `IsMacroType=true`. Les essais ci-dessus corroborent le chargement de cinq fonctions. Les 23 autres restent des déclarations dans ce relevé.

| Famille | Noms et paramètres |
| --- | --- |
| Existence | `RangeExists(sheet_name, range_name)`, `ShapeExists(sheet_name, shape_name)`, `SheetExists(sheet_name)`. |
| Limites et contexte de feuille | `LastCell(sheet_name)`, `LastCellAddress(sheet_name)`, `LastColumnNumber(sheet_name)`, `LastRowNumber(sheet_name)`, `SheetName(reference_cell)`, `UsedRange(cell_in_sheet)`. |
| Propriétés de cellule | `HasFormulas(cell)`, `IsMerged(cell)`. |
| Contrôle de plage et de filtre | `SumColor(range, reference_color_cell)`, `CountColor(range, reference_cell)`, `SumFromAllSheets(reference_range, include_active_sheet)`, `CountUnique(range)`, `CountVisible(range)`, `SlicerItems(slicer_name)`. |
| Recherches multi-correspondances | `Vlookup_Max(lookup_value, table_array, column_number)`, `Vlookup_Min(lookup_value, table_array, column_number)`, `Hlookup_Max(lookup_value, table_array, row_number)`, `Hlookup_Min(lookup_value, table_array, row_number)`. |
| Texte | `CountWords(range)`, `Word_n(text, word_position)`, `Concatenate_Range(cell_range, separator)`. |
| Finance et statistiques | `CAGR(ending_value, beginning_value, number_of_years)`, `Weighted_Average(values, weights)`, `Percent_Change(beginning_value, ending_value)`, `RMS(array)`. |

### Nuances que les exemples ne donnent pas

- **Somme et compte par couleur** — `SumColor` à `0x2CD0` et `CountColor` à `0x2FD0` lisent `Interior.ColorIndex`. Le fond conditionnel affiché mérite un témoin à part. Les couleurs de la carte Clair ne doivent pas changer silencieusement la portée d’une mesure.
- **Visibilité** — `CountVisible` à `0x35D4` vérifie l’absence de masquage de la ligne et de la colonne pour chaque cellule énumérée. La présence d’une valeur ne fait pas partie de cette branche.
- **Distincts** — `CountUnique` à `0x33E4` rejette plus de 65 536 cellules. Les cellules non nulles sont converties par `ToString` en clés de dictionnaire. Témoin à prévoir pour `1` numérique et `"1"` texte.
- **Adresse utilisée** — `UsedRange` à `0x36E0` lit une référence de cellule, récupère le nom de sa feuille puis cherche ce nom dans le classeur actif. Le guide public propose un argument texte. La différence doit être arbitrée par essai.
- **Borne de dernière cellule** — `GetLastCell` à `0x2300` fait deux `Find("*")` séparés, par lignes puis colonnes, et prend leur intersection. La cellule d’intersection peut être vide si les deux bornes proviennent de cellules différentes.
- **Toutes les feuilles** — `SumFromAllSheets` à `0x3224` énumère les feuilles de calcul du classeur actif et somme la même adresse sur chacune. Le booléen décide si la feuille active participe. Témoin nécessaire sur feuille masquée ou références vers un autre classeur.
- **Croissance et variation** — `CAGR` à `0x4900` applique la croissance annualisée puis arrondit à quatre décimales. `Percent_Change` passe aussi par un arrondi. Les deux traversent un convertisseur de culture, avec sortie numérique lorsque le parseur réussit. Tester valeurs négatives, zéro, séparateurs et type d’erreur.

Source éditeur. [Fonctions Excel avancées](https://support.powerusersoftware.com/support/solutions/articles/80001023454-advanced-excel-functions).

## Preuves et limites

| Élément | Identité |
| --- | --- |
| `PowerUser_Excel.dll` | SHA-256 `14612b0446bd2a8391b24c28ac87b6ae9438f557ee61f765530d720396c0dd16`. REA `ev_dabec8d850ce9d8c43a5d665239072a713a70e600b8dd01e320d32260305a212`. |
| `PowerUser_Excel_AddIn.dll` | REA `ev_e983e11b7b86ff6907a3149c9fa8cbdab1e30d711ccdf850e2dc72fcaa46ae1a`. |
| Conteneurs XLL | Ressources PE REA x64 `ev_7562e7bfc1d6ef0bf030f6a0c1ed80ace226d24ef8c2c9700bfce8bcdd6dcc1f`, x86 `ev_03a1c4c19946483142d6d29b532af2ddf970d7eb826e01a3a28e7873a0febd68`. Membres de l’assemblage Excel-DNA `ev_18998000b9ca04e42f0eb47e10c3bd0e21679c86e94e452f677ade4879385c52`. |
| Classeur témoin | SHA-256 `B77BA47F001F967419B519A5A10986BAFDFA4FEF1871B7A7E2114662A11EC342` avant et après. |
| Classeur avec filtre | Original `0D974E931A7BF26F32F81C051F2D45A2796A4F0F522DE88B290DAD31C8B80E61`. Copie enregistrée avec le résultat du collage et de la transposition `AA797A8F3854277D1198A5E4DCCC30676D874D6CB7394DAF76404FD7463D7E6D`. |

Les autres commandes et fonctions du catalogue relèvent de la documentation ou de l’inspection statique. Leur disponibilité complète dans la démo, les sélections complexes, l’annulation et les différences entre versions d’Excel restent à établir.

La contre-vérification du 7 novembre n’a lancé aucun nouveau test Excel. Le hash de l’assemblage installé est inchangé. L’inventaire général des membres REA a rencontré la limite de représentation de métadonnées de 64 Mio, preuve `ev_92d52ba055a0c0bf5a6627d04435800ebe9bc342050a23a14b36011fcb8f573c`. Le catalogue des contrôles réutilise la lecture PEReader ciblée conservée localement, dont le digest figure dans le JSON. Les getters hérités non résolus restent inconnus.

## État des questions

- **Identité et périmètre : établis.** Version locale et hash contrôlés, Excel seul, volet natif identifié.
- **Catalogue documenté : établi dans le périmètre consulté.** Les 26 articles, 28 UDF et fonctions transversales ci-dessus sont recensés. Aucun décompte de fonctions uniques n’est déduit des 135 contrôles.
- **Contrats et défauts : partiellement établis.** Cinq commandes et cinq UDF ont des témoins d’exécution datés ; les autres résultats sont statiques ou documentaires. Les défauts de nettoyage des littéraux et de dépassement de collage sont des cas à éviter dans Clair.
- **Reconstruction : à réaliser.** Les lots proposés ne sont pas annoncés comme implémentés. Priorité : fusions, nettoyage syntaxique, collage visible borné, puis dépivotage, transposition et mesures d’audit.
- **Couverture restante : ouverte.** Filtres et masquages croisés, zones disjointes, protection, erreurs, gros volumes, annulation, Office 32 bits, plusieurs fenêtres et accès complet en démo.

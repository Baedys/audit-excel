# Registre de parité Arixcel / Clair

Établi le 15 octobre 2026. Cible : **Arixcel Explorer 9.0.9623.19616**, MSI fourni par l'utilisateur. Objectif : reproduire les fonctions et interactions du produit de référence avec une implémentation originale. **L'identité fonctionnelle et visuelle n'est pas démontrée.**

Ce registre distingue les observations du binaire, les descriptions de l'éditeur et les exigences de validation proposées. Les statuts ci-dessous décrivent la base Clair 0.1 inspectée (`Core.cs`, `MainForm.cs`, `ExcelBridge.cs`, `AddIn.cs`, tests et documentation), avant les travaux de parité. Ils doivent être réévalués lors de chaque livraison, pas remplacés automatiquement par le résultat d'une compilation.

## État courant 0.6

La matrice initiale ci-dessous sert de référence historique. Les évolutions validées sont les suivantes ; aucune ne suffit seule à prouver une ligne entièrement identique au logiciel original.

| Exigences | Évolution | Preuve et reste à vérifier |
| --- | --- | --- |
| I01, I02 | Chargement COM réparé ; ruban avec sept commandes ; explorateurs simultanés et fermeture globale ; marshaling des callbacks vers le thread UI Excel. | Test réel Excel `/x`, `RibbonOnLoad` et callbacks IDispatch ; deux fenêtres vérifiées. En 0.4, affinité native des fenêtres vérifiée depuis des clients STA et MTA après correction d'un blocage inter-threads. Autres versions, 32 bits, clics physiques et tous les groupes du ruban encore non validés. |
| F01 | AST avec offsets source, opérateurs, fonctions, références, matrices, arguments vides. | 682 assertions. R1C1/DDE et comparaison différentielle avec Arixcel restent ouverts. |
| F02, F06, F07 | Calcul par Excel avec contexte cellule/feuille, portée LET, tableaux qualifiés et aperçu matriciel. | Tests Excel sous locale française. Limite Evaluate 255 caractères ; corps LAMBDA non liés refusés. La résolution du graphe de cellules demeure lexicale. |
| F09 | Sélection d'un nœud syntaxique surligne sa portion de formule. | Aperçu et code UI ; synchronisation inverse, cellules et multisélection encore absentes. |
| F11 | Éditeur F2, écriture, historique d'annulation par fenêtre, contrôle des conflits et restauration du format généré. | Tests Excel sur formules, constantes, tableaux, déversements, changement de type. Annulation native Excel et comparaison exhaustive avec Arixcel non établies. |
| C01, C02, C03 | Périmètre/zone centrale, huit catégories, liens directs, blocs rectangulaires et arbre catégorie/feuille configurable ; ordre Excel et émission des blocs, priorité centrale en frontière, visibilité des feuilles, sélection et clavier. | 64 assertions analytiques, 895 de blocs/ordre, 628 d'arbre et 71 de callbacks UI. Capture réelle de Worksheet.Index ; deux préférences indépendantes et persistantes (62 assertions). Preuves statiques dans [FLOW-EVIDENCE.md](FLOW-EVIDENCE.md), sans comparaison runtime. Lots de collections, R1C1 absent, références dynamiques et multi-classeurs restent incomplets. |
| C04 | Fond et police simultanés, activation de tous les descendants d'un parent, d'un bloc ou d'une cellule, états mixtes, dernière activation ciblée prioritaire, retrait par fenêtre ou conservation explicite. Première analyse sans couleurs ; actualisation remappant les sélections existantes. | 110 assertions de composition, 895 de blocs/activation et 73 contrôles de fenêtres Excel ; 71 callbacks d'arbre supplémentaires sur états mixtes, feuilles ciblées, clavier et sélection. Les couleurs utilisent des règles CF propres à Clair ; l'original applique des formats directs avec historique. Permanence, retrait global entre modèles et superpositions complexes restent ouverts. |
| C05 | Palette flux séparée, XML strict, import/export et défaut enregistré explicitement. | 182 assertions de palette flux, plus les 308 de palette carte conservées après mutualisation du stockage. Analyses synchrones ; annulation et progression non implémentées. |
| M01, M02 | Classification R1C1 par voisins haut/gauche, priorité externe/interfeuille/locale, validation sur constantes. Palette initiale relevée statiquement. | 64 assertions moteur ; capture Excel de références mixtes et validation de cellules vides. Preuves CIL dans [MAP-EVIDENCE.md](MAP-EVIDENCE.md). Références non résolues restent partielles ; comparaison runtime avec Arixcel manquante. |
| M03, M04 | Couche de couleurs de carte sur une/toutes les feuilles, retrait indépendant des autres couches, persistant dans le classeur. | Tests Excel sur formats directs, règles existantes, ordre relatif/StopIfTrue, protection, sauvegarde/réouverture/retrait. En 0.4, injection d'échec pendant la deuxième suppression : restauration de la première règle et ordre des couches vérifiés. Numéros de priorité non identiques si Excel laisse des trous. |
| M05 | Fond/police facultatifs, catégories exclues, palette par défaut enregistrée explicitement, import/export XML Clair. | 308 assertions de sérialisation/validation/atomicité. Effet au clic suivant « Colorer » ; options d'analyse uniques/validation du produit original encore absentes. |

Voir [validation détaillée](VALIDATION.md). Toutes les autres exigences restent dans l'état initial ou sans preuve suffisante. L'objectif global reste ouvert.

## Sources et limites

| Repère | Preuve réutilisée | Portée |
| --- | --- | --- |
| E1 | `work/rea-evidence.json`, `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9` | Inspection REA statique de l'assembly, SHA-256 `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`. Version, ressources, dépendances et identité. Aucun code exécuté. |
| E2 | `work/original-ribbon.xml`, ressource `Arixcel.Explorer.Ribbon.xml`, SHA-256 `63b1d16dfd99cd5a683495d69ac3398ac103b6683dc865e8ff159feb25eaea11` | Commandes, regroupements et callbacks déclarés. Ne prouve pas leur fonctionnement. |
| E3 | `work/type-inventory.json` | Types et noms des méthodes lus via métadonnées .NET. Présence d'un chemin probable, pas preuve de sa sémantique. |
| E4 | `work/il-evidence.json` | Appels CIL ciblés de `Ribbon`, `Maps.Analyser`, `Dependents.Collector`, `Comparison.TaskItem` et `CalcFlow.ParsedFormula`. Ne constitue pas une exécution ni une décompilation complète. |
| E5 | `src/`, `tests/`, `README.md`, `docs/VALIDATION.md` | État observable de Clair. Tests existants ciblés ; aucune comparaison différentielle exhaustive avec Arixcel. |
| W1 | [Exploration](https://www.arixcel.com/explorer/functions/explore) | Description publique actuelle de l'éditeur. |
| W2 | [Dépendants multiples](https://www.arixcel.com/explorer/functions/dependents) | Description publique actuelle de l'éditeur. |
| W3 | [Carte des formules](https://www.arixcel.com/explorer/functions/map) | Description publique actuelle de l'éditeur. |
| W4 | [Flux de calcul](https://www.arixcel.com/explorer/functions/flow) | Description publique actuelle de l'éditeur. |
| W5 | [Comparaison](https://www.arixcel.com/explorer/functions/compare) | Description publique actuelle de l'éditeur. |
| W6 | [Utilitaires](https://www.arixcel.com/explorer/functions/shortcuts) | Description publique actuelle de l'éditeur. |

Les pages officielles sont une application JavaScript. Leur contenu textuel a été lu dans les modules publics référencés par `main-3U56D4JK.js` : `chunk-CA3LRAIw.js` (W1), `chunk-DhMRFsvG.js` (W2), `chunk-BoT_RsSg.js` (W3), `chunk-C0OacgCm.js` (W4), `chunk-BYAT5pzZ.js` (W5), `chunk-DWOaslAG.js` (W6). Ce relevé est une lecture de documentation, pas une observation du logiciel en marche. Les descriptions actuelles peuvent différer de la version fournie. Les fichiers extraits et les données REA demeurent locaux ; ils ne sont pas redistribués avec Clair.

L'inventaire REA des membres a échoué à sa limite de représentation de 64 Mio (`ev_e02f1b6fd380f4eb45bf7081600b9d566f771dbf20ee1fa9e62b546530bc1730`). Aucun résultat vide de cette opération ne constitue une absence de fonction. L'XLAM/VBA, le comportement runtime, les dimensions des fenêtres, les valeurs exactes des couleurs et tous les réglages par défaut restent à observer.

**Statuts :** absent = pas d'implémentation dans la base inspectée ; partiel = une partie existe mais les critères manquent ; non vérifié = preuve runtime insuffisante. Aucune ligne ne signifie « identique » sur la seule base d'un nom de méthode.

## 1. Intégration et parcours communs

| ID | Exigence et preuve de référence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| I01 | Complément disponible dans Excel ; ruban à groupes Explorer, Analysis, Utilities, Maintenance (E1, E2). | Non vérifié : chargement COM en échec connu ; un seul bouton déclaré. | Installation propre, redémarrage Excel, découverte effective, callbacks et fermeture sans processus résiduel. Vérifier chaque architecture Excel annoncée. |
| I02 | Plusieurs fenêtres non modales et commande Close All (E2, E3 `ModelessForm`, W1). | Absent : `Explore` ferme la fenêtre précédente. | Ouvrir trois explorations, travailler dans Excel, fermer toutes sans perte de navigation ni effet résiduel. |
| I03 | Historique de lieux/fenêtres et retour au point initial (E3 `History`, W1). | Partiel : pile de cellules du modèle local. | Retour entre feuilles, classeurs et fenêtres ; distinguer fermeture avec conservation de sélection et retour à la source. |
| I04 | Raccourcis principaux personnalisables, réaffectation, import/export paramètres (E2, E3 `ShortcutManager`, `JsonReadStore`). | Absent. | Modifier, enregistrer, relancer, importer/exporter ; détecter conflits ; restaurer les affectations lors de désinstallation. |
| I05 | Réglages persistants des fenêtres, polices, langue/formules locales, couleurs (E3 `ModelessForm`, `ExplorerOptions`). | Absent. | Affichage lisible en DPI usuels ; positions/colonnes conservées ; Français/anglais et A1/R1C1 vérifiés. |
| I06 | Progression, arrêt, annulation ; état incomplet explicite (E3 `ProgressForm`, `CancelToken`, `StopToken`). | Partiel : plafonds annoncés, capture synchrone. | Grand classeur, interruption pendant lecture/analyse, absence de faux résultat complet et Excel toujours réactif. |
| I07 | Actualisation et invalidation après modification (E3 `OnUserFocusIn`, `RefreshModel`, `FormulaChangeAction`). | Partiel : rechargement manuel du classeur. | Modifier une formule, renommer/fermer un classeur, supprimer une feuille : rafraîchir sans réutiliser silencieusement une capture périmée. |
| I08 | Comportement sûr avec protection, masquage et erreurs (E3 `Unhide`, E4 `Collector.Search`). | Partiel : refus de navigation vers feuille masquée. | Respecter réglage d'affichage des cellules masquées ; protection et zone inaccessible signalées ; restaurer état après abandon. |
| I09 | À propos, vérification de mises à jour et informations produit (E2). | Absent. | Équivalents propres à Clair ; aucun serveur, identité ou mécanisme de licence Arixcel réutilisé. Le mécanisme commercial du produit original n'est pas une preuve de parité des fonctions d'audit. |

## 2. Exploration de formules et références

Le produit documente un arbre de composants avec valeurs et emplacements, des choix conditionnels, des cibles de recherches/références dynamiques, la navigation depuis la formule et son édition. Les noms de méthodes corroborent ces familles ; ils ne révèlent pas tous les cas limites. Source : W1 ; E3 `Element`, `Model`, `ExplorerForm`, `FunctionSyntax`.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| F01 | Arbre syntaxique : fonctions, arguments, opérateurs, sous-expressions et références. | Absent : arbre de cellules seulement. | `=IF(A1>0,SUM(B1:B3)*2,C1)` révèle niveaux, texte exact et liens des composants ; imbrication et priorité correctes. |
| F02 | Valeur intermédiaire et emplacement de chaque composant. | Absent. | Comparer valeurs et adresses aux résultats Excel sur nombres, textes, booléens, erreurs, matrices et références multiples. Ne jamais présenter une estimation comme exacte. |
| F03 | Branche sélectionnée des fonctions conditionnelles. | Absent : toutes les références lexicales sont incluses. | Cas vrai/faux/erreur de IF, IFS, CHOOSE, SWITCH ; distinguer références présentes et effectivement utilisées. |
| F04 | Cible effective des recherches et fonctions de référence. | Absent ou incomplet. | INDEX, OFFSET, INDIRECT, VLOOKUP et LOOKUP avec cible locale/interfeuille, résultat vide, erreur et mode approché ; observer les limites exactes de la version cible. |
| F05 | Cellules retenues par critères et filtrage de tableaux. | Absent. | Résultats disjoints, critères multiples, dates, jokers et erreurs ; vérifier SUMIF/SUMIFS, MIN/MAX et filtres logiques avec preuves d'Excel. |
| F06 | Noms définis, LAMBDA et validation d'entrée dans l'arbre. | Partiel : noms absolus seulement. | Noms locaux/classeur, noms relatifs et récursifs, fonctions nommées, formule de validation et source de liste déroulante. |
| F07 | Références structurées, 3D, tableaux, LET et contexte d'évaluation. | Incomplet, avertissements présents. | Tables, colonne courante, déversements, intersection implicite, variables LET, paramètres LAMBDA et plages 3D ; comportement de référence à établir pour chaque famille. E3 expose `Parse3DReference`, `IsTableReference`, `HaveLambda`, `EvaluateArrayFunction`. |
| F08 | Références vers tout classeur ouvert. | Absent : capture d'un classeur, préfixes externes non résolus. | Deux classeurs ouverts avec noms proches ; navigation et valeurs contextualisées, avertissement explicite après fermeture. |
| F09 | Sélection synchronisée entre arbre, portion de formule et cellules ; multisélection. | Absent : texte simple en lecture seule. | Cliquer composant puis texte correspondant ; Ctrl/Shift sélectionnent plusieurs composants sans perdre leurs emplacements. |
| F10 | Surlignage source/antécédents/dépendants configurable. | Absent. | Couleurs distinctes ; nettoyage à fermeture et changement de sélection ; formats antérieurs intacts. |
| F11 | Édition de composant ou formule, validation, annulation, undo. | Absent. | F2, Entrée, Échap, Ctrl+Z ; protection et erreurs gérées ; formule exacte restaurée ; une édition invalide n'écrase pas la cellule. |
| F12 | Navigation clavier et refresh, expansion/réduction, réexploration. | Partiel : TreeView et exploration de cellule. | Tester Ctrl+Q, Ctrl+Backspace, flèches, Ctrl+flèches, Échap/Entrée selon option ; ordre et focus comparés au produit. |
| F13 | Mode d'évaluation précis et signal d'imprécision. | Absent. | E3 `PreciseEvaluation`, `MayBeImprecise` : déterminer différences runtime avant de promettre ce réglage ; couvrir fonctions volatiles/UDF sans effets involontaires. |

## 3. Antécédents et dépendants de plusieurs cellules

W2 décrit des dépendances **directes**, regroupées en blocs, avec sélection multiple et poursuite de l'exploration. OFFSET, INDIRECT et fonctions VBA volatiles y sont explicitement exclus. Ne pas confondre cette limite documentée avec l'exploration détaillée d'une formule F04. E4 montre `ShowPrecedents`, `ShowDependents`, `NavigateArrow`, activation de fenêtres et nettoyage des flèches.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| D01 | Source multizone et multicellule. | Absent : racine unique. | Plages contiguës/disjointes, fusionnées et feuilles multiples ; union sans doublons ; source conservée pour retour. |
| D02 | Résultats regroupés par blocs et feuille, avec nombres fiables. | Partiel : liste/arbre de cellules. | Blocs de formules relatives identiques ; réglage de regroupement par feuille ; compte cohérent avec les cellules réellement trouvées. |
| D03 | Sélection multiple, surbrillance et exploration successive. | Absent. | Sélection de plusieurs résultats puis trace suivante ; navigation clavier et source initiale retrouvée. |
| D04 | Limites du traçage explicites. | Partiel : avertissements lexicaux. | Classeur externe fermé, feuille protégée, opération annulée, référence indirecte : aucune affirmation « aucun dépendant » si recherche incomplète. |

## 4. Carte des formules

W3 décrit une coloration temporaire de la grille selon la cohérence des formules : uniques, copiées horizontalement/verticalement, puis catégories de références pour les formules uniques. Plusieurs feuilles peuvent être traitées. Les couleurs originales peuvent être restaurées même après sauvegarde et réouverture. E3 expose l'analyse de validation, les réglages de couleurs et la conservation/restauration des formats conditionnels ; E4 confirme l'analyse de `FormulaR1C1`, d'erreurs et de références externes/interfeuilles.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| M01 | Classification par structure relative et voisinage. | Absent : types de valeurs seulement. | Distinguer unique, copie horizontale, verticale, dans deux directions ; références mixtes/absolues, trous et frontières de blocs. |
| M02 | Catégories de contenu, erreurs, validation, références. | Partiel : formule/nombre/texte/booléen/erreur. | Fixtures couvrant formule sans référence, interfeuille, externe, cellule vide et validation ; priorité des catégories établie par observation. |
| M03 | Appliquer/retirer sur la grille et sur plusieurs feuilles. | Absent. | Commande de bascule, aucune modification de valeurs/formules, restauration des couleurs et formats conditionnels. |
| M04 | Persistance de la restauration après sauvegarde/réouverture. | Absent. | Appliquer, sauvegarder une copie, fermer/réouvrir, retirer ; comparer propriété par propriété à la copie originale. |
| M05 | Palette, absence de remplissage, options d'analyse, import/export. | Absent. | Paramètres persistants et effet immédiat ; catégorie exclue conserve ses propriétés. |

## 5. Flux de calcul

W4 distingue périmètre analysé et zone centrale. Vue entrées/sorties : cellules ayant seulement des dépendants, cellules ayant les deux relations, cellules ayant seulement des antécédents, formules isolées. Vue flux traversants : antécédents externes à la zone, cellules internes alimentées de l'extérieur, cellules internes alimentant l'extérieur, dépendants externes. Toutes ces relations sont restreintes au périmètre. OFFSET/INDIRECT sont explicitement exclus. E3 expose les modèles Scope, Area, ViewMode et deux modes de couleur, remplissage et police.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| C01 | Périmètre classeur/feuille/plage et zone centrale distincte. | Absent : fermeture transitive d'une cellule. | Choisir chaque périmètre ; interdire zone centrale hors périmètre ; vérifier résultats sur plusieurs feuilles. |
| C02 | Deux vues et catégories décrites ci-dessus. | Absent. | Graphe miniature avec une cellule de chaque rôle, un cycle, une formule isolée, une référence hors périmètre ; comparer ensembles exacts. |
| C03 | Résultats groupés, comptes, exploration et actualisation. | Partiel : grille plate. | Groupes par catégorie/feuille ; basculer vues sans perdre sélection ; relancer après modification. |
| C04 | Surbrillance sélectionnable, couleurs police/remplissage. | Absent. | Activer/désactiver groupes, fermer normalement, fermer en conservant, retirer ultérieurement, rendre permanent ; restauration vérifiée. |
| C05 | Paramètres de palette persistants et analyse annulable. | Absent. | Sauvegarde/chargement palette, progression, annulation et invalidation après fermeture classeur. |

## 6. Comparaison

W5 décrit référence classeur/feuille/plage, plusieurs comparateurs, modes formules/valeurs, option ignorant les entrées, blocs de différences, alignement par insertion puis suppression de lignes/colonnes, fusion directionnelle, édition, undo et rapport. E3 corrobore `AlignmentDiff`, `UndoAlignmentDiff`, `CopyFormats`, `ValueTolerance`, `MaxAlignmentDistance`. E4 montre un contrôle de tailles identiques dans `TaskItem.Compare` ; il ne prouve pas l'absence d'une normalisation préalable.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| P01 | Référence et comparateurs aux trois échelles, plusieurs cibles. | Partiel : deux captures classeur par nom de feuille/adresse. | Feuille renommée, plages décalées, comparateurs multiples et sélection de coin supérieur gauche ; appariement explicite. |
| P02 | Formules/valeurs, tolérance, ignorance des entrées. | Partiel : deux modes et tolérance absolue. | Nombres proches, booléens, erreurs, vide/chaîne vide, entrées changées ; établir sémantique exacte de tolérance. |
| P03 | Détection des lignes/colonnes insérées et alignement. | Absent. | Insertion/suppression en début/milieu/fin, doublons de libellés et formules déplacées ; éviter milliers de faux écarts. |
| P04 | Appliquer puis retirer un alignement, avec undo. | Absent. | Copies jetables : formules, noms, tailles et formats restent cohérents ; annulation restaure toutes les propriétés touchées. |
| P05 | Regrouper différences voisines, mettre en évidence segments. | Absent : lignes par cellule. | Blocs contigus/disjoints, côté actif, navigation des deux côtés et surlignage des parties modifiées. |
| P06 | Fusion directionnelle et copie de formats/feuilles. | Absent ; seulement lecture. | Copier gauche/droite puis re-comparer ; protection, collision et undo. Les détails de copie de formats sont à confirmer runtime. |
| P07 | Édition, exploration, échange de côtés et rapport. | Partiel : rapport texte et navigation vers cellule. | Ctrl+Tab/Ctrl+S, F2, Ctrl+Z, exploration depuis résultat ; rapport reproductible avec périmètre, options, écarts et état incomplet. |

## 7. Utilitaires et paramètres

W6 décrit des utilitaires paramétrables, plusieurs variantes d'une action, raccourcis partagés formant des cycles, visibilité au ruban et sauvegarde/restauration. E3 `Utilities.Actions` fournit l'inventaire ci-dessous. Chaque nom observé nécessite encore un test comportemental et des valeurs de paramètres observées ; il ne suffit pas d'ajouter un bouton.

Tous les groupes suivants sont **absents** de la base inspectée :

| ID | Famille / méthodes observées E3 | Acceptation requise |
| --- | --- | --- |
| U01 | Mise en page : `ColumnWidths`, `AutoFit`, `Zoom`, `ToggleGridlines`, `RemovePageBreaks`, `NewWindow`. | Périmètre de feuilles correct, paramètres persistants, absence d'effet sur classeur non ciblé. |
| U02 | Formats : `CellColor`, `TextColor`, `Style`, `NumberFormat`, `ConvertNumberFormat`, `DecimalPlaces`, `Border`, `TextUnderline`, `TextControl`. | Cas de sélection multiple, formats mixtes, cellule protégée ; modification et undo précis. |
| U03 | Presse-papiers/copie : `CopySum`, `PasteSpecial`, `PasteDuplicate`. | Types ignorés, modes de collage ; copie exacte et préservation des références internes testées séparément. |
| U04 | Transformation : `ReplaceFormulas`, `ChangeFormulas`, `Multiply`, `TextCase`, `SmartFillRight`, `SmartFillDown`. | Formules et constantes distinguées ; regex/paramètres invalides rejetés ; étendue de remplissage et undo vérifiés. |
| U05 | Sélection/navigation : `SpecialCells`, `JumpTo`, `JumpToUniqueFormula`, `SelectA1`, `UnhideSheets`, `ToggleRefStyle`. | Filtrage exact ; masquage restauré ; feuille protégée et zone vide gérées. |
| U06 | Configuration : `Manager`, `SetupForm`, `CommandCycles`, raccourcis, liste active/masquée, suffixes. | Ajouter, dupliquer, paramétrer, ordonner, masquer, exporter/importer, cycler ; conflits détectés et absence de raccourcis orphelins. |

## Priorités et preuve finale

1. **P0 — expérience de base :** I01, F01/F02 et I02. Charger réellement dans Excel et afficher les composants de formule ; un arbre de cellules seul ne reproduit pas l'outil demandé.
2. **P1 — profondeur de l'audit :** F03–F13, D01–D04. Construire corpus de formules, contexte d'évaluation Excel, navigation, sélection et édition avec restauration.
3. **P1 — fonctions d'analyse distinctes :** M01–M05, C01–C05, P01–P07. Un parcours transitif ne remplace pas la classification de flux ; une liste colorée ne remplace pas une carte de grille.
4. **P2 — couverture opérationnelle :** U01–U06, paramètres, clavier, robustesse/performance et installation sur architectures prises en charge.

Les priorités ordonnent le travail ; elles n'excluent aucune exigence du résultat demandé. Avant déclaration de parité : observer Arixcel sur un corpus jetable commun, relever écrans/résultats/actions/restaurations, exécuter le même corpus avec Clair, conserver chaque différence et sa résolution. Tester séparément fonctions, apparence, clavier, persistance, performances et cycle de vie Excel. La compilation, les tests actuels et une revue de code réussie ne prouvent pas cette équivalence.

Restent notamment inconnus : apparence exacte des fenêtres de cette version, seuils de regroupement, algorithmes d'alignement, limites d'évaluation, défauts des options hors carte, rôle de l'XLAM, reprise après incident et compatibilité précise avec chaque version d'Excel. La priorité des catégories et la palette de carte sont désormais établies statiquement, mais leur équivalence runtime reste à vérifier. Ces inconnues maintiennent l'objectif ouvert.

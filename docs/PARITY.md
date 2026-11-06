# Registre de parité Arixcel / Clair

Établi le 15 octobre 2026. Cible : **Arixcel Explorer 9.0.9623.19616**, MSI fourni par l'utilisateur. Objectif : reproduire les fonctions et interactions du produit de référence avec une implémentation originale. **L'identité fonctionnelle et visuelle n'est pas démontrée.**

Ce registre distingue les observations du binaire, les descriptions de l'éditeur et les exigences de validation proposées. Les statuts ci-dessous décrivent la base Clair 0.1 inspectée (`Core.cs`, `MainForm.cs`, `ExcelBridge.cs`, `AddIn.cs`, tests et documentation), avant les travaux de parité. Ils doivent être réévalués lors de chaque livraison, pas remplacés automatiquement par le résultat d'une compilation.

## État courant 0.6

La matrice initiale ci-dessous sert de référence historique. Les évolutions validées sont les suivantes ; aucune ne suffit seule à prouver une ligne entièrement identique au logiciel original.

| Exigences | Évolution | Preuve et reste à vérifier |
| --- | --- | --- |
| I01, I02 | Chargement COM réparé ; ruban avec sept commandes ; explorateurs simultanés et fermeture globale ; marshaling des callbacks vers le thread UI Excel. | Test réel Excel `/x`, `RibbonOnLoad` et callbacks IDispatch ; deux fenêtres vérifiées. En 0.4, affinité native des fenêtres vérifiée depuis des clients STA et MTA après correction d'un blocage inter-threads. Autres versions, 32 bits, clics physiques et tous les groupes du ruban encore non validés. |
| F01 | AST A1 invariant avec offsets source, opérateurs, fonctions, références, matrices, arguments vides. Correspondance avec le texte local et R1C1 fourni par Excel. | Les syntaxes dont Excel réécrit la structure peuvent perdre le surlignage. DDE et comparaison différentielle exhaustive avec Arixcel restent ouverts. |
| F02, F06, F07 | Calcul par Excel avec contexte cellule/feuille, portée LET, tableaux qualifiés et aperçu matriciel. | Tests Excel sous locale française. Limite Evaluate 255 caractères ; corps LAMBDA non liés refusés. La résolution du graphe de cellules demeure lexicale. |
| F09 | Synchronisation dans les deux sens, sélection multiple et surlignage des plages résolues. | Témoins Excel sur Ctrl/Maj, texte disjoint et union par feuille. Volet ancré conservé lors d'une sélection vers un autre classeur. |
| F11, I05, I07 | Saisie directe et F2 dans le volet, en A1 ou R1C1, avec la langue d’Excel ou les fonctions anglaises. Les matrices CSE sont éditées par bloc jusqu’à 255 caractères. Les sous-tracés ont des piles liées. Le XLL transmet le lot d’éditions à `Application.OnUndo` à la sortie du volet. | Témoins Excel sur le contenu typé, les piles sur trois niveaux, la reprise après une action native et le rollback sur conflit. Restent notamment les longues matrices CSE et les éditions déléguées à une comparaison. |
| C01, C02, C03 | Périmètre/zone centrale, huit catégories, liens directs, blocs rectangulaires et arbre catégorie/feuille configurable ; ordre Excel et émission des blocs, priorité centrale en frontière, visibilité des feuilles, sélection et clavier. | 895 assertions de blocs/ordre, 628 d'arbre et 71 de callbacks UI. Capture réelle de Worksheet.Index ; deux préférences indépendantes et persistantes (62 assertions). Les six blocs du témoin 3D concordent dans Excel. Lots de collections, R1C1 absent, références dynamiques et multi-classeurs restent incomplets. [Preuves du flux](FLOW-EVIDENCE.md) |
| C04 | Fond et police simultanés, activation de tous les descendants d'un parent, d'un bloc ou d'une cellule, états mixtes, dernière activation ciblée prioritaire, retrait par fenêtre ou conservation explicite. Première analyse sans couleurs ; actualisation remappant les sélections existantes. | 110 assertions de composition, 895 de blocs/activation et 73 contrôles de fenêtres Excel ; 71 callbacks d'arbre supplémentaires sur états mixtes, feuilles ciblées, clavier et sélection. Les couleurs utilisent des règles CF propres à Clair ; l'original applique des formats directs avec historique. Permanence, retrait global entre modèles et superpositions complexes restent ouverts. |
| C05 | Palette flux séparée, XML strict, import/export et défaut enregistré explicitement. | 182 assertions de palette flux, plus les 308 de palette carte conservées après mutualisation du stockage. Analyses synchrones ; annulation et progression non implémentées. |
| M01, M02 | Classification R1C1 par voisins haut/gauche, priorité externe/interfeuille/locale, validation sur constantes. Palette initiale relevée statiquement. | 64 assertions moteur ; capture Excel de références mixtes et validation de cellules vides. Preuves CIL dans [MAP-EVIDENCE.md](MAP-EVIDENCE.md). Références non résolues restent partielles ; comparaison runtime avec Arixcel manquante. |
| M03, M04 | Couche de couleurs de carte sur une/toutes les feuilles, retrait indépendant des autres couches, persistant dans le classeur. | Tests Excel sur formats directs, règles existantes, ordre relatif/StopIfTrue, protection, sauvegarde/réouverture/retrait. En 0.4, injection d'échec pendant la deuxième suppression : restauration de la première règle et ordre des couches vérifiés. Numéros de priorité non identiques si Excel laisse des trous. |
| M05 | Fond/police facultatifs, catégories exclues, palette par défaut enregistrée explicitement, import/export XML Clair. | 308 assertions de sérialisation/validation/atomicité. Effet au clic suivant « Colorer » ; options d'analyse uniques/validation du produit original encore absentes. |

Voir [validation détaillée](VALIDATION.md). Toutes les autres exigences restent dans l'état initial ou sans preuve suffisante. L'objectif global reste ouvert.

## Branche de développement — couleurs directes et options de carte

Les observations natives d'Arixcel couvrent retrait, permanence, cartes sur bandes B/H et sauvegarde/réouverture d'une copie jetable. La carte originale masque les fonds CF de toute la feuille ; retirer la première bande les restaure même si la seconde reste active. Après sauvegarde avec carte active, l'original observé perd ces fonds. L'utilisateur a explicitement choisi une restauration complète dans Clair, excluant ce défaut.

Le développement raccorde formats directs partagés, références physiques, commandes globales et deux journaux XML propres à Clair. Le journal CF conserve neuf propriétés de fond, distinctes de la police et des conditions ; événements, restauration différée et réparation de métadonnées sont couverts par leur corpus. La version 0.6 empaquetée reste distincte de cette branche en brouillon. Résultats et limites dans [VALIDATION.md](VALIDATION.md) et [DIRECT-COLOR-EVIDENCE.md](DIRECT-COLOR-EVIDENCE.md). Références structurelles complexes, familles CF non couvertes, performances et chargement natif STA/MTA restent à vérifier ; C04, M03 et M04 ne sont pas déclarés entièrement identiques à Arixcel.

M05 reprend les onze fonds de carte, le réglage des formules uniques, les couleurs personnalisées et l'enregistrement immédiat. Les changements recolorent ou réanalysent les zones conservées suivant les cas relevés dans le binaire. Le volet importe les profils JSON Arixcel modernes et plats et exporte le format moderne. Le XML Clair accepte les profils v1/v2 et garde les options propres à Clair. [Preuves de la carte](MAP-EVIDENCE.md)

## Présentation Excel demandée pour Clair

L'utilisateur demande un ruban et un volet latéral qui laisse la grille dégagée. Dans le complément Excel, les vues Formule, Carte, Flux et Comparer s'ouvrent dans un volet Office redimensionnable à droite. Une fenêtre de classeur a son propre volet. La bascule des couleurs reste directement accessible depuis le ruban. La palette de carte occupe le même volet, avec **Terminé** pour retrouver les résultats. Les réglages avancés Clair et les sélecteurs de fichier ou de couleur personnalisée s'ouvrent dans des dialogues ponctuels.

## Branche retenue dans le volet Formule

F03 progresse avec le gras sur le résultat choisi par `IF`, `IFS`, `CHOOSE`, `SWITCH`, `IFERROR` et `IFNA`. Le CIL de `Element.ResolveChosenChild` (RVA `155012`, `IL_01cc` à `IL_052f`) donne les règles de choix. `Element.PostProcess` (RVA `153512`) marque le résultat retenu en gras tout en conservant les arguments visibles. Les conversions scalaires proviennent de `ExcelValue.AsBoolean` et `AsDouble` dans `Arixcel.Common.dll` (SHA-256 `425eb00972ba6d5aaa45ee6e8a4369e1fa40ff1ea2f002da2522893c0711eb01`). Le comparateur `SWITCH` suit l'égalité culturelle et la distinction observée entre blanc et texte vide.

Clair lit les opérandes dans le contexte de la cellule, avec la portée `LET`. Il conserve les offsets du texte et marque uniquement les choix établis. Le résultat décrit une nouvelle évaluation des conditions. Les fonctions volatiles peuvent donc produire une autre branche qu'au dernier calcul de la cellule. Les plages d'une cellule et tableaux d'un élément sont lus comme scalaires pour les six fonctions ci-dessus.

Pour `FILTER`, le code distingue un masque booléen établi d'un tableau brut. Les comparaisons, fonctions à résultat booléen et certaines transformations de masques peuvent établir le repère. Pour `XLOOKUP`, un résultat de type plage déclenche une comparaison d'adresses externes avec le secours. Sinon, un `#N/A` typé de `XMATCH` en contexte scalaire déclenche le repère en conservant les modes de correspondance et de recherche. Ces chemins sont décrits par le CIL de `ResolveChosenChild`, RVA `155012`, `IL_0534` à `IL_06d4`. Les quatre observations natives et la comparaison sur le même classeur sont consignées dans la validation.

Les masques transmis par des noms définis, les comparaisons matricielles avec erreurs, les plages synthétiques issues d'une recherche matricielle, l'intersection implicite complexe et `LAMBDA` restent à comparer. Le graphe des références reste lexical. [Validation](VALIDATION.md)

## Cibles effectives dans le volet Formule

F04 associe une adresse au nœud de fonction et garde son résultat dans l'aperçu. `ResolveRangeOfFunction` (RVA `174880`) et `ResolveDisplayProperties` (RVA `163476`) d'Arixcel séparent la plage physique de la valeur affichée. Pour les formes scalaires de `VLOOKUP`, `HLOOKUP` et `LOOKUP`, le CIL (RVA `178948` et `179532`) passe par un `MATCH` puis un `INDEX`. `MATCH` et `XMATCH` utilisent leur position calculée. `INDIRECT` à référence relative est ancré sur la cellule source avant lecture de la plage. La propagation de la cible vers une somme à argument unique a été observée dans le produit original.

Clair utilise les plages renvoyées par Excel pour `INDEX`, `OFFSET`, `INDIRECT` et `XLOOKUP`, et reconstruit les cibles scalaires des autres recherches. Il recalcule la cible avant navigation. Les unions synthétiques produites par une recherche matricielle, la propagation générale aux agrégats et certains contextes `LET` et `LAMBDA` restent à comparer. Les adresses du volet complètent le graphe lexical. [Validation](VALIDATION.md)

## Cellules retenues par critères dans le volet Formule

F05 affiche l'union retenue pour `SUMIF(S)`, `COUNTIF(S)` et `AVERAGEIF(S)`. La somme et la moyenne utilisent la plage de résultats. Le comptage utilise la première plage de critères. Les groupes **Filtre** affichent les cellules et le compteur de leur propre paire, même si l'intersection est vide. Un groupe unique qui couvre tous les arguments laisse ces arguments au premier niveau. `FILTER` peut montrer séparément les données renvoyées et les cellules sources d'une comparaison directe. **Aller à la cible dans Excel** actualise et sélectionne les zones disjointes.

Les preuves statiques proviennent de `Element.ResolveRangeOfFunction` (RVA `174880`), `ResolveRangeAfterValue` (`171356`), `ResolveRangeFinal` (`170092`), `ChildGroups.CreateForCriteriaArguments` (`357436`), `Element.ProcessChildGroups` (`180820`) et `MatchCriterion` (`185664`, `186600`, `187640`). Les masques multiples sont croisés par position. Pour `SUMIF` et `AVERAGEIF`, les dimensions d'origine de la plage testée déterminent l'étendue de la plage de résultat à partir de son coin supérieur gauche. Les jokers de l'original sont ancrés au début du texte. Cette règle peut produire une adresse alors que le résultat natif d'Excel diffère. Les valeurs des formules affichées dans Clair proviennent d'Excel. [Validation](VALIDATION.md)

`MINIFS` et `MAXIFS` désignent une cellule après filtrage. `MIN`, `MAX`, `MINA`, `MAXA`, `SMALL`, `LARGE`, `SUBTOTAL` en modes 4 et 5 et `AGGREGATE` en modes 4, 5, 14 et 15 peuvent pointer vers la première cellule correspondant au résultat d'Excel. Les autres modes valides de `SUBTOTAL` reprennent la plage entière quand elle est fournie par un seul argument. La propagation générale des plages d'agrégats, les masques de comparaison contenant des erreurs et les formats locaux restent partiels. Clair borne le relevé à 20 000 cellules et 128 zones par adresse. La limite de caractères d'Excel Evaluate reste applicable aux expressions.

Le gras suit `Element.ResolveChosenChildAfterValue` (RVA `157116`) et son prédicat scalaire (RVA `361672`). `MIN`, `MAX`, `MINA`, `MAXA`, `SUBTOTAL` et `AGGREGATE` en modes 4 et 5 retiennent le premier argument dont la valeur correspond au résultat. Pour une référence de plusieurs cellules, `Element.ResolveValue` (RVA `159288`, `IL_06e9` à `IL_08bb`) calcule la fonction sur cet argument seul, en reprenant le mode et les options. Une référence d'une cellule garde sa valeur brute. Les erreurs Excel admissibles se comparent par code. Dans le choix d'argument, un booléen reste distinct d'un nombre, y compris pour `MINA` et `MAXA`. Le ciblage des cellules de ces deux fonctions applique sa propre conversion. `Element.PostProcess` (RVA `153512`, `IL_012a` à `IL_01c8`) retire le gras des branches imbriquées écartées. L'aperçu réduit des plages ordinaires reste à étendre. [Validation](VALIDATION.md)

## Références 3D dans le volet Formule

F07 traite les plages 3D directes. `Element.Parse3DReference` (RVA `146804`) développe chaque feuille entre les deux bornes et conserve les coordonnées. Clair lit la collection des feuilles du classeur à l'actualisation. Il développe les feuilles masquées, écarte les onglets graphiques et conserve les références déjà sélectionnées quand leur feuille reste dans l'intervalle. Les sous-nœuds de feuille surlignent la référence 3D d'origine dans le texte de la formule. [Insertion de feuilles dans une référence 3D, Microsoft](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets)

Pour les agrégats reconnus, le volet évalue la référence seule sous la fonction. Sur une feuille, un enfant d'une cellule conserve sa valeur brute et un enfant de plusieurs cellules affiche la réduction. `ChildRangesUsedByThisFunction` (RVA `171084`) relie les plages des feuilles au ciblage des extrémums. La fonction peut pointer vers une cellule et la feuille retenue apparaît en gras. La navigation de Clair conserve le masquage des feuilles. Les chemins externes sur disque restent ouverts. [Validation](VALIDATION.md)

Dans le classeur capturé, les arbres et le Flux suivent chaque plage physique d'une référence 3D. Le Flux coupe les liens à son périmètre. Les chemins d'Arixcel diffèrent selon la vue. Le Flux lit les enfants physiques de la référence dans `CalcFlow.ParsedFormula` (RVA `302492`). Les listes d'antécédents et de dépendants utilisent les flèches d'audit Excel dans `Collector.Search` (RVA `203188`) et `FollowArrow` (RVA `203648`). Deux listes de dépendants et six blocs du Flux concordent sur le témoin. Les antécédents concordent comme ensemble de cellules. [Validation](VALIDATION.md)

## Dépendances entre classeurs ouverts

F08 couvre les références directes A1 et les noms absolus dans l'Explorateur. Clair capture les classeurs ouverts de la même instance Excel et distingue les cibles par classeur, feuille et adresse. L'arbre et les plages affichent le nom du classeur externe. Le surlignage s'applique à chaque classeur et le volet évalue les formules sélectionnées dans leur contexte. Un clic garde la fenêtre du volet active. **Voir cellule** active la cible d'une relation. Le classeur lié possède son propre volet pour l'édition.

Dans Arixcel, `Element.ResolveRange` demande à Excel de résoudre la référence. `OpenAnyExternalLink` (RVA `194492`) propose une fois par chemin d'ouvrir le fichier lors de la sélection d'une référence ou au passage du curseur dans le texte de la formule. Clair affiche le chemin et les actions **Ouvrir le classeur** et **Ignorer** dans le volet. Après ouverture, il conserve le contexte Excel et reconstruit la trace avec la même racine, l'historique et le nœud sélectionné. L'ouverture programmatique désactive les macros VBA et la mise à jour des liens du fichier ouvert. Les références fermées et les noms relatifs restent signalés tant qu'ils sont non résolus. Le Flux garde le périmètre du classeur analysé. [Validation](VALIDATION.md)

## Sélections multicellules dans le volet

D01 et D02 couvrent les sélections contiguës et disjointes d'une feuille. `SelectionRoot` garde l'adresse de la sélection Excel, y compris ses zones séparées. Le parcours source utilise la zone utile étendue jusqu'à A1. `FollowArrow` (RVA `203648`) retire la source des résultats. `Collector.GetResults` (RVA `204104`) distingue les formules relatives pour les dépendants. `RangeItem` (RVA `367292`) compte chaque bloc pour un. Clair applique ces règles au graphe capturé. Les antécédents incluent les cellules vides de la plage et sont bornés par `Worksheet.UsedRange`. Le découpage exact des rectangles suit les segments de colonnes communs à des lignes adjacentes.

Le volet groupe les blocs par feuille sur demande. Quatre tracés concordent avec Arixcel sur le même classeur. Les feuilles groupées dans Excel demandent d'être dégroupées avant le tracé.

D03 couvre maintenant la sélection de plusieurs résultats, y compris depuis une cellule. Arixcel active le mode multiple dans `ExplorerForm.InitializeComponent` (RVA `198288`) et `DependentsForm.InitializeComponent` (RVA `206136`). `FormulaForm.OpenChildForm` (RVA `133832`) transmet seulement les nœuds associés à une plage et conserve la feuille racine. `MultiRoot` (RVA `366960`) replie les sources et développe les résultats, avec deux compteurs. `Model.AddRanges` (RVA `211560`) place en tête les plages de la feuille racine. Clair reprend cette structure dans le volet. Le retour rétablit les résultats choisis et le regroupement par feuille conserve la sélection.

Un clic ou une sélection au clavier dans le volet sélectionne automatiquement l'union sur la feuille du résultat courant. `FormulaForm.SelectOnSheet` (RVA `134056`) distingue les couleurs et la sélection Excel. Clair suit le résultat dans la fenêtre Excel active. Il conserve le masquage des feuilles. **Voir cellule** permet aussi la navigation explicite.

La surbrillance des dépendances reprend les fonds temporaires rose, bleu et vert de l'original. La source et les résultats sélectionnés peuvent se trouver sur plusieurs feuilles ou classeurs ouverts. Les fonds sont découpés aux chevauchements et partagent la pile Carte et Flux. Les formats conditionnels gardent leurs règles et leur priorité. [Validation](VALIDATION.md)

## Sélection de sous-expressions dans le volet

`SelectFromFormulaBox` (RVA `192800`) descend vers le composant explicite sous le curseur. Ctrl conserve les sélections existantes ou rend courant un ancêtre déjà sélectionné. Dans l'arbre, Ctrl bascule un nœud et Maj étend la sélection visible. `HighlightFormula` (RVA `193772`) souligne les portions choisies en bleu et conserve un curseur sans sélection de caractères.

`FormulaForm.SelectOnSheet` (RVA `134056`) colore la source rose et les plages résolues bleues. `SelectorForm.SelectOnSheet` (RVA `65920`) sélectionne l'union sur la feuille du nœud courant. Clair réutilise la pile de couleurs de trace et garde la fenêtre du volet active lorsque le résultat appartient à un autre classeur. Le menu **Aller à la cible dans Excel** permet d'y aller. [Validation](VALIDATION.md)

## Raccourcis dans le volet

`SelectorForm.MoveActiveCell` (RVA `64992`) part du coin haut gauche de la première zone Excel. Ctrl+flèche rejoint le bord demandé ou passe à la zone suivante avec bouclage. Le coin sur l'autre axe reste choisi si sa dimension le permet. Une seule zone est sélectionnée dans Excel pendant le parcours. Alt+S rétablit la sélection complète avec la zone mémorisée en tête. Clair garde la fenêtre du volet et reprend le parcours depuis son dernier état.

Alt+Z déplie ou replie la branche courante en gardant le premier nœud ouvert au repli. Clair limite le dépliage de l'arbre des dépendances à 5 000 nœuds par commande. Ctrl+Q et Ctrl+Maj+Q explorent les plages sélectionnées dans les deux arbres. Sur la racine seule, dans le même mode, le raccourci bascule l'expansion. [Validation](VALIDATION.md)

`FormulaForm.OnFormClosed` (RVA `134492`) applique le retour sur Échap si l'option est active. Arixcel resélectionne alors le parent ou la plage de lancement avant nettoyage. Entrée suit le chemin OK et garde la sélection. Clair représente le parent dans le même volet ancré. Il garde sa sélection Excel et son coin mémorisés, et rétablit la plage de lancement ainsi que les onglets groupés disponibles au niveau racine. [Validation](VALIDATION.md)

Le raccourci Ctrl+Retour arrière passe par `ModelessForm.KeyDownHandler` (RVA `58168`) depuis un formulaire, et par `History.GoBack` (RVA `64192`) depuis Excel. Clair enregistre une commande XLL avec `OnKey`, conserve les sélections physiques et les vues vivantes de l'instance et actualise le tracé au retour. Le focus suit la fenêtre Excel propriétaire. Quand Arixcel est connecté, Clair lui laisse l'association. [Validation](VALIDATION.md) · [Microsoft OnKey](https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey)

## Dévoilement pendant la navigation

`SelectorForm.Unhide` (RVA `64624`) révèle d'abord la feuille, puis les axes selon `UnhideCells`, activé par défaut. `CellUnhider` réduit les axes entiers à une bande en A ou en ligne 1 et borne les autres cibles par la zone utilisée. Les axes masqués voisins suivent le niveau de plan. À la fermeture, le natif épargne la feuille active et les axes de la sélection (RVA `69764`). Clair suit ces règles dans les arbres de dépendances et de formules. L'option du volet concerne les prochaines navigations.

Clair restaure le masquage avec **Remasquer** et à la fermeture par l'hôte. Avant l'enregistrement du classeur, il tente cette restauration et annule la sauvegarde si Excel la refuse. Il garde les masques des lignes sous filtre actif, isole les écritures des feuilles groupées et restaure les changements d'une navigation refusée. Le chemin d'enregistrement natif inspecté traite le journal de couleurs. [Validation](VALIDATION.md)

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
| I05 | Réglages persistants des fenêtres, polices, langue/formules locales, couleurs (E3 `ModelessForm`, `ExplorerOptions`). | Le volet mémorise le choix de langue, suit la notation Excel au retour du focus et actualise les volets ouverts lors du changement d’option. | Réglages de police, positions, colonnes et DPI, ainsi que les dérogations de style au lancement depuis Comparaison et Flux. |
| I06 | Progression, arrêt, annulation ; état incomplet explicite (E3 `ProgressForm`, `CancelToken`, `StopToken`). | Partiel : plafonds annoncés, capture synchrone. | Grand classeur, interruption pendant lecture/analyse, absence de faux résultat complet et Excel toujours réactif. |
| I07 | Actualisation et invalidation après modification (E3 `OnUserFocusIn`, `RefreshModel`, `FormulaChangeAction`). | Alt+R reconstruit depuis les plages sources Excel. L’édition et l’annulation locale actualisent sans resélectionner la feuille. L’annulation globale sélectionne la cible la plus ancienne et actualise le volet d’origine s’il est ouvert. | Comparer l’actualisation automatique au retour du focus après des écritures externes et la propagation entre fenêtres simultanées. |
| I08 | Comportement sûr avec protection, masquage et erreurs (E3 `Unhide`, E4 `Collector.Search`). | Dévoilement de feuille et d'axes, restauration et protection des filtres actifs dans les arbres. | Feuille très masquée, axes sélectionnés et voisins, groupes, refus de protection et sauvegarde. |
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
| F08 | Références vers tout classeur ouvert. | Références A1 directes et noms absolus dans l'Explorateur, avec navigation, valeurs contextualisées et couleurs. Proposition d'ouverture par chemin dans le volet. | Trois classeurs aux feuilles homonymes, refus d'un chemin, ouverture d'un nom externe, conservation du contexte et de l'historique. |
| F09 | Sélection synchronisée entre arbre, portion de formule et cellules ; multisélection. | Sous-expressions sélectionnables par l'arbre ou le curseur, Ctrl/Maj, union Excel sur la feuille courante et couleurs sur les autres feuilles. | Plages physiques, portions disjointes, restauration des fonds et lien externe. |
| F10 | Surlignage source/antécédents/dépendants configurable. | Absent. | Couleurs distinctes ; nettoyage à fermeture et changement de sélection ; formats antérieurs intacts. |
| F11 | Édition de la formule affichée, validation, annulation, undo. | La saisie utilise le style affiché à l’ouverture du brouillon. F2 conserve le curseur, Entrée valide et Ctrl+Z distingue brouillon et édition. Un enfant sans action remonte à la pile parente. Sa fermeture groupe les éditions liées dans le parent immédiat. À la sortie du volet, les éditions locales forment une action Excel. Le lot restaure les actions de la plus récente à la plus ancienne et compense une annulation partielle en cas de conflit. | Couvrir les longues matrices CSE et les éditions déléguées par Comparaison. Vérifier le clavier et le transfert de focus physiques, la composition IME et les autres configurations Office. |
| F12 | Navigation clavier et refresh, expansion/réduction, réexploration. | Ctrl+flèches, Alt+S, Alt+Z, Ctrl+Q/Ctrl+Maj+Q, Alt+R, Échap et Entrée dans le volet. Ctrl+Retour arrière entre la grille et l'historique via le XLL. | Coins et bouclage, sélection multizone, autre fenêtre, dépliage, nouvelle trace, déplacement et invalidation de source. Réception de la frappe par Excel en mode édition et essai natif 32 bits restent à couvrir. |
| F13 | Mode d'évaluation précis et signal d'imprécision. | Absent. | E3 `PreciseEvaluation`, `MayBeImprecise` : déterminer différences runtime avant de promettre ce réglage ; couvrir fonctions volatiles/UDF sans effets involontaires. |

## 3. Antécédents et dépendants de plusieurs cellules

W2 décrit des dépendances **directes**, regroupées en blocs, avec sélection multiple et poursuite de l'exploration. OFFSET, INDIRECT et fonctions VBA volatiles y sont explicitement exclus. Ne pas confondre cette limite documentée avec l'exploration détaillée d'une formule F04. E4 montre `ShowPrecedents`, `ShowDependents`, `NavigateArrow`, activation de fenêtres et nettoyage des flèches.

| ID | Exigence | État initial | Acceptation requise |
| --- | --- | --- | --- |
| D01 | Source multizone et multicellule. | Absent : racine unique. | Plages contiguës/disjointes, fusionnées et feuilles multiples ; union sans doublons ; source conservée pour retour. |
| D02 | Résultats regroupés par blocs et feuille, avec nombres fiables. | Partiel : liste/arbre de cellules. | Blocs de formules relatives identiques ; réglage de regroupement par feuille ; compte cohérent avec les cellules réellement trouvées. |
| D03 | Sélection multiple, surbrillance et exploration successive. | Absent. | Sélection de plusieurs résultats puis trace suivante ; navigation clavier et source initiale retrouvée. |
| D04 | Limites du traçage explicites. | Résolution partielle signalée pour les liens non résolus et les erreurs de capture, sur une cellule ou une plage. | Classeur externe fermé et référence indirecte. Restent à comparer la feuille protégée et l'interruption d'un tracé. |

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

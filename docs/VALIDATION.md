# Revue et validation

## Références 3D dans le volet Formule — 6 novembre 2026

Arixcel 9.0.9623.19616 et Clair ont ouvert deux copies de `Arixcel3DReferences.xlsx` au même SHA-256. Les valeurs, les repères et les emplacements relevés concordent sur ces cinq formules. `Fev` est masquée dans le classeur témoin.

| Cellule | Formule | Observations |
| --- | --- | --- |
| `B2` | `SUM(Jan:Mar!$B$2)` | Référence 3D `14`. Feuilles `Jan`, `Fev`, `Mar` dans cet ordre, valeurs `8`, `3`, `3`, toutes à `B2`. Les emplacements de la fonction et du groupe restent vides. |
| `B3` | `MIN(Jan:Mar!$B$2,7)` | La fonction indique `Fev!B2`. La référence 3D et `Fev` sont en gras. `Mar` a la même valeur. |
| `B4` | `MAX(Jan:Mar!$B$2:$B$3,5)` | Valeurs par feuille `20`, `19`, `40`, chacune sur `B2:B3`. La référence 3D et `Mar` sont en gras. La fonction indique `Mar!B3`. |
| `B5` | `SUM('North 1:North O''Brien'!$D$2)` | Deux feuilles, valeurs `6` et `11`, somme `17`. |
| `B8` | `SUM(Jan:Mar!B2,4)` | La référence 3D affiche `14` et la fonction `18`. |

Excel a réécrit `MIN(Mar:Jan!$B$2,7)` en `MIN(Jan:Mar!$B$2,7)` à l'enregistrement de la formule. La comparaison avec l'original couvre les références du classeur courant. [Références 3D, Microsoft](https://support.microsoft.com/en-us/excel/create-a-3-d-reference-to-the-same-cell-range-on-multiple-worksheets)

Les **88 contrôles Excel** des références 3D couvrent aussi `COUNT`, un onglet graphique intercalé, une feuille insérée entre les bornes, la sélection conservée dans le volet, la navigation vers la plage d'une feuille visible et la restauration du contexte Excel. Les modules des branches (**106 contrôles**), des critères (**167**) et des cibles (**97**) passent sur le même binaire. Les **68 groupes sans Excel** passent, dont **698 assertions** de syntaxe. Un essai Clair distinct sur deux classeurs ouverts retrouve les deux adresses du classeur source et leur somme `30`. Les cas du graphique et du classeur externe restent à comparer dans Arixcel.

## Argument retenu pour un extrémum — 6 novembre 2026

Le gras d'Arixcel 9.0.9623.19616 et celui calculé par Clair concordent sur dix formules de `ArixcelExtremumChoices.xlsx`. Les copies ouvertes dans Excel ont le même SHA-256. La comparaison identifie chaque argument par sa position, même lorsque deux arguments ont le même texte.

| Cellule | Formule ou cas | Argument en gras |
| --- | --- | --- |
| `B2` | `MIN` avec `5` et `5+0` à égalité | `5` |
| `B3` | `MAX` avec une plage puis sa cellule maximale | `$G$2:$G$4` |
| `B4` | `MINA(TRUE,1,4)` | `1` |
| `B8` | `MAX(NA(),17,NA())` | Premier `NA()` |
| `B10` | `AGGREGATE` mode 4 | `$G$2:$G$4` |
| `B11` | `SUBTOTAL` mode 5 | `$G$2:$G$4` |
| `B12` | `AGGREGATE` mode 14 | Aucun |
| `B13` | `MIN({5},7)` | `{5}` |
| `B14` | `MIN(1,IF(TRUE,2,3),4)` | `1`. Le `IF` reste consultable avec ses arguments en texte normal. |
| `B16` | `MAX` avec une cellule contenant le texte `30` avant le nombre `30` | `$H$3` |

Les **106 contrôles Excel** des branches incluent le gras et l'adresse dans le volet, le contexte `LET`, les modes 4 et 5 et les tableaux de plusieurs éléments. Les modules des cibles (**97 contrôles**) et des critères (**167**) passent séparément sur le même binaire. Les **68 groupes sans Excel** passent, dont **164 assertions** sur les choix et **31** nouvelles sur les extrémums. La comparaison ci-dessus porte sur les repères d'arguments hors 3D. L'aperçu des plages ordinaires conserve ses valeurs brutes.

## Groupes de filtres et cellules d'extrémum — 6 novembre 2026

Arixcel 9.0.9623.19616 et Clair ont ouvert deux copies de `ArixcelCriteriaGroups.xlsx` au même SHA-256. Les sept adresses et les deux emplacements vides ci-dessous concordent. Les adresses sont présentées dans la notation d'Arixcel.

| Cellule | Fonction | Emplacement observé |
| --- | --- | --- |
| `B4` | `SUMIFS` | Résultat `Data!B2;B6`. **Filtre 1** `Data!A2;A4;A6`. La plage résultat précède les deux groupes dans l'arbre. |
| `B14` | `MINA` | `Data!E2`, cellule de texte associée au résultat numérique 0. |
| `B18` | `SUBTOTAL(105,…)` | `Data!G2:G6`, plage entière. |
| `B19` | `AGGREGATE(14,6,…,2)` | `Data!H5`, deuxième plus grande valeur malgré une erreur dans la plage. |
| `B20` | `SUBTOTAL(5,…)` | `Data!G4`, cellule minimum dans une ligne masquée. |
| `B22` | `SUMIFS` sans intersection | Fonction sans emplacement. **Filtre 1** `Data!A2;A4;A6`. **Filtre 2** marqué vide, sans emplacement. |

Les **167 contrôles Excel** couvrent aussi les groupes complets, les dimensions incompatibles, les critères indisponibles, `MAXA`, `SMALL` et `LARGE`. Dans le volet, sélectionner un groupe surligne les deux arguments, affiche leur compteur et garde le groupe ouvert lors de l'évaluation. Entrée relit le critère avant de sélectionner ses zones. Les **68 groupes sans Excel** passent.

## Cellules retenues par critères — 6 novembre 2026

Dans Excel, la colonne **Location** d'Arixcel 9.0.9623.19616 donne les emplacements suivants sur le classeur témoin.

| Cellule | Fonction | Emplacement |
| --- | --- | --- |
| `B2` | `SUMIF` avec joker | `Data!B2;B4;B6` |
| `B3` | `COUNTIFS` avec deux critères | `Data!A2;A6` |
| `B4` | `SUMIFS` avec deux critères | `Data!B2;B6` |
| `B5` | `MINIFS` | `Data!B4` |
| `B8` | `FILTER` dans `SUM` | `Data!B2;B6` pour le filtre et `Data!A2;A6` pour la comparaison. La somme a un emplacement vide. |

Sur B2 et B4, les nœuds de somme d'Arixcel affichent 0 alors que la cellule et la racine indiquent 45. Clair utilise la valeur calculée par Excel et affiche l'adresse indépendamment.

Clair retrouve les **six emplacements** sur les cinq formules d'une copie du même classeur, vérifiée par SHA-256. Les **68 contrôles Excel** des critères couvrent aussi les dates, erreurs, plages de somme courtes, colonnes entières, masques horizontaux, secours physique, absence de cible et navigation par zones disjointes. Les modules des branches (**63 assertions**) et des cibles de recherche (**97**) passent dans des processus séparés avec le même binaire. Les **68 groupes sans Excel** passent, dont **21 assertions** sur la correspondance des critères.

## Cibles effectives des formules — 6 novembre 2026

Les adresses ci-dessous ont été lues dans la colonne **Location** d'Arixcel 9.0.9623.19616 dans Excel, sur un classeur jetable. Le repère appartient à la fonction. La racine de l'explorateur original indique la cellule source.

| Cellule | Fonction | Plage observée |
| --- | --- | --- |
| `B2` | `INDEX` sur deux colonnes | `Data!C3` |
| `B3` | `OFFSET` dans `SUM` | `Data!B3:B4` sur les deux nœuds |
| `B4` | `INDIRECT` interfeuille dans `SUM` | `Data!B3:B4` sur les deux nœuds |
| `B5` | `VLOOKUP` exact | `Data!C3` |
| `B7` | `XLOOKUP` sur une plage de retour | `Data!C3` |

Clair retrouve les **sept emplacements** dans une copie identique du classeur, vérifiée par SHA-256. Les valeurs des cinq cellules concordent aussi. Le contrôle Clair passe par le moteur connecté à Excel.

Les **97 contrôles Excel** couvrent aussi `HLOOKUP`, `LOOKUP`, `MATCH`, `XMATCH`, les erreurs sans cible, les références interclasseurs, le volet et la navigation après modification de la clé. Les **63 contrôles Excel** des branches et les **67 groupes sans Excel** passent sur le même binaire. Les tests distinguent la valeur d'une expression de sa cible. La conversion R1C1 utilise la cellule source comme point d'ancrage. Les cibles matricielles synthétiques restent hors de ce contrôle. [Application.ConvertFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.application.convertformula)

## Palette de carte dans le volet — 6 novembre 2026

- **67 groupes sans Excel réussis**. Les nouveaux groupes couvrent **69 assertions** de profils JSON Arixcel et de persistance XML, ainsi que **19 contrôles** de la palette du volet. Les anciens profils Clair restent lisibles. Les rendus à **320 × 540** et **420 × 650** montrent une liste défilante et les commandes de pied de volet accessibles.
- Le module des options réussit **52 contrôles Excel**, dont l'ouverture de la palette dans la vue Carte, son transfert entre vues, la mise à jour d'une carte active, le maintien du propriétaire et des cellules exclues, le refus d'un profil verrouillé et la conservation des fonds conditionnels déjà restaurés.
- Le module de régression des cartes réussit **235 contrôles Excel** sur le même binaire, dans un processus STA séparé.
- L'interface Arixcel 9.0.9623.19616 a été observée dans un classeur jetable. La disposition et les commandes ont été comparées à celles relevées statiquement. Les tests de fichiers couvrent le schéma observé dans le binaire. L'échange aller-retour dans Arixcel reste à vérifier. [Preuves de la carte](MAP-EVIDENCE.md)

## Valeurs de secours de `FILTER` et `XLOOKUP` — 6 novembre 2026

- **65 groupes sans Excel réussis**. Le module des branches compte désormais **63 contrôles Excel**, dont 26 nouveaux cas. Il couvre les masques de comparaison, les tableaux bruts, les masques arithmétiques, les liaisons `LET`, l'égalité d'adresse et les modes exact, approché, inversé et joker.
- Arixcel 9.0.9623.19616 et Clair ont été testés dans deux instances Excel dédiées sur le même classeur jetable, vérifié par SHA-256. Le gras a été observé dans l'interface Arixcel. La décision de Clair a été vérifiée dans son moteur lié à Excel.

| Cellule | Cas observé | Repère Arixcel et Clair |
| --- | --- | --- |
| `B2` | `FILTER` lit directement une plage de `FALSE` et renvoie `901`. | `901` reste en texte normal. |
| `B16` | `XLOOKUP` trouve `20` dans la même cellule que le secours. | `$G$3` est marqué. |
| `B22` | Le masque de comparaison est vide. | `909` est marqué. |
| `B24` | Le masque de comparaison est vide, mais sa taille diffère des données et Excel affiche `#VALEUR!`. | `910` est marqué dans l'arbre. |

Les cas matriciels avec erreurs, les noms définis et les plages synthétiques d'une recherche matricielle restent ouverts. [Registre de parité](PARITY.md)

## Branches conditionnelles — 6 novembre 2026

- **65 groupes sans Excel réussis**, dont **133 assertions** de choix conditionnel. Les contrôles couvrent les six fonctions, les conversions, les erreurs Excel, les blancs, les tableaux, les arguments omis, les branches imbriquées et les offsets source.
- **37 contrôles** dans un processus Excel dédié. Le volet marque et actualise le gras, y compris après changement de la valeur d'entrée. Il efface le repère si la formule source a changé. Il conserve les arguments inspectables et laisse l'évaluation automatique en attente tant que l'onglet Formule est masqué. Les contrôles couvrent aussi `LET`, les limites de `Evaluate`, une sélection multiple, le groupe de feuilles et le retour des flags d'Excel.
- Le module de régression des expressions passe dans un second processus STA avec le même binaire. Il vérifie la locale, l'édition, l'annulation, les conflits, les références structurées et **550 assertions** de restauration du contexte sur 50 cycles entre deux classeurs.
- Les règles de choix sont issues de la lecture statique d'Arixcel 9.0.9623.19616. Les essais d'exécution couvrent Clair dans Excel. La comparaison native des cas limites avec Arixcel reste ouverte. [Registre de parité](PARITY.md)

## Volet latéral Excel — 6 novembre 2026

- Compilation .NET Framework et **64 groupes sans Excel réussis**. Aperçus WinForms compacts examinés à **420 × 650** et **320 × 540**, avec la formule, les arbres, les commandes de carte et les filtres de flux empilés ou repliés selon la largeur.
- Dans une instance Excel dédiée, `test-addin.ps1` charge le complément enregistré et reçoit le ruban et `CTPFactoryAvailable`. Les quatre vues réutilisent un volet qui réduit le bord droit de la zone de feuille. Aucun explorateur de premier niveau n'est créé. Une seconde fenêtre du même classeur ouvre son propre volet. Fermer cette fenêtre retire son volet. **Tout fermer** redonne sa largeur initiale à la grille. Le test retire l'enregistrement temporaire.
- Deux tentatives initiales ont expiré avant les contrôles du volet, pendant l'acquisition du modèle Excel ou l'ouverture du classeur. Le test garde des délais séparés et ouvre explicitement son fichier jetable si l'instance est prête et vide. Le passage réussi utilise le complément réel avec Office 64 bits. [Documentation Office sur les volets personnalisés](https://learn.microsoft.com/en-us/office/vba/api/office.ictpfactory.createctp)

## Options d'analyse des cartes — 6 novembre 2026

- Compilation .NET Framework et **64 groupes sans Excel réussis**. Le moteur ajoute **67 assertions** sur les options. Le stockage de palette atteint **375 assertions** avec migration v1/v2. L'éditeur couvre **166 contrôles** sur la case d'analyse, l'import, l'export et la remise par défaut.
- Le module Excel des options vérifie **40 assertions** sur les bandes séparées, les propriétaires conservés, la réutilisation des sous-types détaillés, le reclassement de la validation, le déplacement après insertion et l'exclusion des blancs hors capture dans une carte de feuille. Il vérifie aussi que la réanalyse ne remasque pas un fond CF déjà restauré. L'instance et le classeur sont jetables.
- Sur le même binaire, trois modules Excel isolés complètent le contrôle ciblé avec **420 assertions** de bascule native, **235** de cartes WinForms et **62** du port CF. Les quatre modules passent dans des processus STA séparés. Le contrôle natif appelle le callback géré dans Excel. Le clic sur le ruban, le chargement enregistré et la suite globale restent hors de ce passage.
- Comparaison des options avec les chemins CIL du binaire fourni. Le dialogue Clair reste modal et applique sur **Utiliser**. L'original applique et enregistre au changement depuis une fenêtre accessible du ruban. Le format d'échange Arixcel, l'interaction native de cette palette et la comparaison visuelle restent ouverts. [Preuves de la carte](MAP-EVIDENCE.md)

## Fonds conditionnels — 4 novembre 2026

L'utilisateur demande de restaurer complètement les fonds CF après sauvegarde/réouverture, y compris lorsque les couleurs directes sont conservées. Le défaut de perte de fonds observé dans Arixcel est explicitement exclu.

- Compilation application/DLL réussie ; **62 groupes sans Excel, zéro échec**. Nouveaux corpus : **263 assertions** de snapshots/XML CF, **16** de modes de références et **622** de remplacement/réparation coordonnée des journaux. Le reçu de réparation est sérialisable sous .NET Framework et vérifié après désérialisation.
- Port CF : **62 contrôles** sur RGB, thème/nuance, motif et None, avec relecture exacte des neuf propriétés. Masque et restitution préservent règles, conditions, polices, bordures, priorités et StopIfTrue.
- Carte native finale : **420 assertions**, dont masquage CF de toute la feuille pour une bande, retrait précoce de la première carte, carte transparente, groupement de feuilles, exclusion d'une feuille protégée, insertion, palette, grands UsedRanges et sauvegarde/réouverture après déchargement sans explorateur. **16 nouveaux contrôles** vérifient la conservation du groupe, de la fenêtre, de la feuille, de la sélection, de la cellule active et des flags pendant création/retrait avec une feuille protégée.
- Cartes WinForms : **235 assertions** sur le binaire précédant le dernier ajustement d'identité de fenêtre ; le scénario CF vérifie aussi DisplayFormat, permanence, fermeture du créateur et retrait global ultérieur. Le dernier ajustement compare les fenêtres par HWND ; il conserve les objets COM et les identités physiques des feuilles et classeurs.
- Scénarios CF en processus STA séparés : propriété par carte, sauvegarde CF seule, report puis nouvelle sauvegarde, reprise après setters défaillants, événements Open/BeforeSave CF seuls, conservation des couleurs directes avec restitution CF, masque None interrompu et nettoyage natif, réassociation après retour d'un témoin, partie XML recréée, adoption atomique sur deux feuilles et réparation XML vérifiée. Les dix cas hors témoin de formule passent sur le premier binaire figé ; le témoin de formule et les parcours Save/fault/events repassent sur le binaire avec contexte. Les ajustements finaux d'identité de fenêtre sont validés par le contrôle natif groupé et le port ci-dessus. Aucun passage global unique de tous ces cas sur le dernier binaire n'est revendiqué.
- La première suite complète de ce travail a produit **85 groupes réussis, 6 échecs COM**. Les six modules ont échoué en acquisition/configuration d'instances Excel fraîches, avant leurs appels au bridge/CF. Les modules directs restants passent, notamment port58, session79, flux23, persistance95, événements21 et contexte550. Les reprises isolées ont ensuite révélé deux défauts de fixture et un vrai défaut de groupement : watcher demandé avant son acquisition ; `Modify` changeant format numérique et priorités ; setter CF atteignant une autre feuille protégée du groupe. Ces cas sont corrigés et repris séparément. **Aucune suite unique à 91/0 n'est annoncée.**
- Sonde de fenêtre détenue : 21 comparaisons avant/après Activate renvoient le même HWND et des IUnknown différents, avec une seule fenêtre de classeur. Le guard initial d'identité de fenêtre était donc invalide ; Window.hWnd le remplace. La sonde termine proprement avec son processus Excel fermé. Ce résultat n'explique pas les arrêts natifs et déconnexions COM précédemment observés.
- Revue indépendante en lecture seule : aucun constat restant sur DTO, coordination XML, suivi CF et contexte. La revue n'exécute pas Office. Les essais réels et la sonde d'identité relèvent de processus détenus, sans nouvelle preuve de chargement enregistré du complément en STA/MTA.

Restent ouverts : groupes dans plusieurs fenêtres, sélection de forme dans un groupe, références CF structurelles complexes, remplacement de règles indiscernables, familles sans Interior/dégradés, performances et identité visuelle exhaustive. Les archives 0.6 ne sont pas remplacées par cette branche en brouillon.

## Contrôles du 3 novembre 2026

- Bascule native par zones : compilation application/DLL réussie, puis **59 groupes sans Excel réussis, zéro échec**. La classification possède désormais **80 assertions**, dont les frontières haute/gauche d'une bande, les blancs synthétisés et la validation ; les **5 326 assertions** du planificateur restent réussies. Le callback natif appelle la bascule sans ouvrir d'explorateur. La capture des seules créations réutilise le lecteur commun ; les retraits utilisent le descripteur physique sans lire le contenu des cellules.
- Suite Excel globale de ce raccordement : **84 groupes réussis, 2 échecs** dans la préparation des nouveaux tests de sélection et de bascule. La première fixture tentait de sélectionner une plage sur une feuille non active du groupe ; la seconde protégeait une feuille encore groupée. Fixtures corrigées en sélectionnant la plage de la feuille active et en protégeant/déprotégeant chaque feuille séparément avant regroupement. Aucun correctif produit n'est déduit de ces échecs de préparation. Les autres modules passent, dont **110 assertions des cartes partagées**, **38 des références de zones**, **79 de session**, **95 de persistance**, **58 du port**, **23 de flux direct**, **21 d'événements**, **73 des fenêtres de flux**, **95 des anciennes couches CF** et **550 de contexte/restauration**.
- Reprise ciblée finale des deux modules corrigés, sur les mêmes binaires produit : **2 modules réussis, zéro échec**, code de sortie 0. **43 contrôles de sélection** couvrent notamment feuilles groupées/protégées, forme refusée, bornes et seuil de réduction ; **147 assertions de bascule** couvrent bandes B/H, état groupé mixte, carte transparente, palette, insertion, grands UsedRanges dans le pipeline complet et déchargement sans explorateur. Après OnDisconnection, SaveAs/réouverture conserve l'historique et la restauration retrouve les bases directes. Les appels ShowMap sont exécutés avec une véritable application Excel depuis un client C# STA ; ils ne constituent ni un clic physique ni une nouvelle preuve de chargement du complément par Excel. La suite globale n'a pas été relancée après les seules corrections de fixtures ; le résultat 84/2 et la reprise 2/0 sont conservés séparément.
- Original installé revérifié : Arixcel Explorer 9.0.9623.19616, SHA-256 `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`, identique à la cible des observations conservées.
- Régression confirmée dans une instance Excel détenue : après modification d'un caractère, `Font.ColorIndex` et `Font.Color` renvoient `DBNull`. Le retrait échoue et le remplissage temporaire reste appliqué. Correction de réconciliation par propriété, adoption et sauvegarde ; tests réels supplémentaires pour retrait multi-cellules, historique rechargé et historique reporté.
- Suite complète `test.ps1 -Excel` terminée avec code 0 : **82 groupes réussis, zéro échec**. Dont **79 assertions de session directe**, **95 de persistance**, **58 du port Excel**, **59 de carte**, **23 de flux**, **21 d'événements**, **73 des fenêtres existantes** et **95 des anciennes couches CF**. Les 550 assertions de contexte/restauration réussissent également. Cette réussite ne démontre pas l'absence définitive des intermittences observées auparavant. Revue indépendante du correctif : aucun constat.
- Diagnostic natif : un contrôle sans Clair a rencontré `RPC_E_CALL_REJECTED` pendant le démarrage ; un autre s'est terminé avant d'exposer l'objet modèle. Une sonde avec Clair et des budgets distincts par phase a ouvert son classeur après 60,9 secondes, avec `Connect=True`, `Status=Clair 0.6 connected` et `IsRibbonLoaded=True`. Dialogue Excel de récupération après démarrage raté observé et démarrage normal choisi dans le seul processus détenu. Cela confirme le chargement et le ruban de cette sonde, sans valider les deux explorateurs ni la fermeture native.
- `test-addin.ps1` donne maintenant 45 secondes à chacune des phases fenêtre native, objet Application et classeur ; le dernier HRESULT est conservé en cas d'échec. L'ancien délai commun pouvait expirer après la découverte tardive de la fenêtre. Les alertes ne sont désactivées qu'après ouverture de la fixture. Les contre-essais complets restent en échec intermittent avant assertions : délai d'ouverture avec `0x8002000B`, puis absence d'objet modèle sous 45 secondes. Les résultats ne prouvent donc pas une correction du démarrage ; la PR reste en brouillon. Aucun réglage Office ni complément tiers modifié.
- Dernier contrôle natif du build corrigé, client PowerShell STA : fenêtre native obtenue en 3,3 secondes, mais fixture toujours absente après le budget d'ouverture ; dernier HRESULT `0x8002000B`. Aucun dialogue Excel/Clair exposé par l'inventaire de fenêtres au contrôle final. Les assertions de ruban et deux explorateurs n'ont pas été atteintes. Processus détenu terminé et enregistrement temporaire retiré. Les contrôles natifs complets STA et MTA restent ouverts.
- Diagnostic complémentaire : alertes Office explicites proposant de désactiver `solver.xlam`, puis Power-user, observées dans des processus détenus. Aucun complément tiers n'a été désactivé volontairement. Ces alertes et les échecs également reproduits sans Clair empêchent d'attribuer tous les délais au complément. Une sonde visible sans Clair ouvre la fixture en 52,7 secondes ; elle sert ensuite au contrôle différentiel de carte, sans valider le smoke complet Clair.
- Carte Arixcel testée réellement : bascule/apparence CF/retrait, deux bandes disjointes et retrait de la première, puis Save/fermeture/réouverture du classeur et réponse Oui. Fonds CF masqués sur toute la feuille, rétablis au retrait du premier modèle même si un autre subsiste ; après Save/reopen, couleurs directes restaurées mais fonds CF perdus. Formules, priorités, StopIfTrue et polices des règles restent identiques dans ce corpus. Valeurs et limites détaillées dans [DIRECT-COLOR-EVIDENCE.md](DIRECT-COLOR-EVIDENCE.md). Cela ne prouve ni tous les types CF ni le redémarrage complet d'Excel.
- Partage des cartes et palettes entre fenêtres du même Excel : suite pure finale **59 groupes réussis, zéro échec**, dont **5 326 assertions** du planificateur de bascule par rectangles physiques. Ce planificateur reste indépendant de la commande native et du choix de sélection Excel.
- Contrôle ciblé final des cartes : **110 assertions Excel réussies**, code de sortie 0. Même propriétaire entre vues, palette globale, conservation après fermeture, isolation entre applications Excel, purge après fermeture/réouverture homonyme ou suppression de feuille, déplacement après suppression partielle et protection des nouvelles cellules à l'ancienne adresse. Une modification de palette filtre les bindings morts ; une demande explicite sur une référence périmée reste refusée. Revue indépendante du partage, du planificateur et du dernier correctif : aucun constat.
- Avant ce dernier filtrage limité à la palette, le contrôle ciblé commun a également réussi les **73 contrôles des fenêtres de flux**, **23 de flux direct** et **21 d'événements**, code de sortie 0. Ces modules sont inchangés depuis ce passage. La suite complète de 82 groupes ci-dessus précède le nouveau partage et le groupe pur ajouté ; aucune nouvelle réussite globale de 83 groupes n'est revendiquée. Le masquage CF original et son défaut de sauvegarde ne sont pas activés dans Clair.

## Travaux locaux non livrés — moteur de couleurs directes

Ces résultats ne remplacent pas encore la livraison 0.6. Le raccordement local des cartes, flux, commandes globales et persistance utilise désormais les formats directs ; sa suite complète est en cours de validation. Preuves originales et limites dans [DIRECT-COLOR-EVIDENCE.md](DIRECT-COLOR-EVIDENCE.md).

- Raccordement local du 18 octobre : 159 assertions XML pures ; 46 contrôles initiaux de carte, 70 contrôles initiaux de persistance. Des régressions supplémentaires ont été ajoutées après ces premiers essais ; leurs résultats doivent être confirmés par la suite finale.
- Le test des événements autonomes a exposé un callback BeforeSave hors thread avec ComEventsHelper. Sink AppEvents sous StandardOleMarshalObject : **21 contrôles Excel réussis** après correction, couvrant sauvegarde, réouverture, Oui/Non/Annuler, refus de confirmation et restauration reportée.
- Première suite globale du raccordement : échec préexistant dans l'évaluation SUM(A1:A2), active cell attendue Other!E10, restaurée Other!D9. Les échecs antérieurs de COLUMN()/Activate ne sont pas déclarés corrigés ; diagnostic séparé en cours. Aucun résultat global vert n'est affirmé ici.
- Résultat terminal de cette suite : **80 groupes réussis, 2 échecs**. Carte : **59 contrôles** ; persistance : **82 contrôles** ; événements : **21 contrôles** ; fenêtres de flux existantes : **73 contrôles**. Le second échec vient du nouveau test qui lançait Analyser avant le callback Shown/chargement Excel : ordre corrigé, fermeture de la fixture garantie en cas d'échec ; contre-essai encore requis. Ce formulaire abandonné a ensuite présenté une erreur de RCW séparé pendant le pompage d'événements : message acquitté uniquement dans le processus de test. Après terminaison, seule l'instance Excel utilisateur antérieure restait ouverte.
- Dernière compilation sans avertissement et suite pure : **58 groupes réussis, 0 échec**. Aperçu de flux contrôlé après regroupement des actions : commandes de conservation désormais entièrement visibles à 1160 × 820. La taille minimale et la reproduction de l'interface originale restent ouvertes.
- Contre-essai borné de l'évaluation, sans observations internes : une séquence capture/opérations complète réussie, dont **550 assertions de restauration**, `COLUMN()=2` et cellule E10 conservée. L'intermittence n'est pas reproduite et sa cause reste inconnue ; aucun correctif spéculatif.
- Le contre-test de flux sous protection a identifié un second défaut réel : `SpecialCells` lève 0x800A03EC sur la feuille protégée et la capture classait silencieusement une formule comme constante. Sonde minimale dans une instance détenue : `UsedRange.HasFormula=null`, erreur SpecialCells, `Capture(B1).IsFormula=false` malgré `=A1*2`. Fallback ajouté pour confirmer par `Range.HasFormula` les candidats lorsque la présence globale reste incertaine ; le test inclut du texte formaté `@` commençant par `=`. Revue indépendante sans constat sur ce correctif ; validation runtime finale encore en cours.
- Contre-essai final ciblé : **23 contrôles de flux et 21 d'événements réussis**, processus terminé avec code 0. Couverture : édition manuelle, palette incrémentale, insertion, capture protégée formule/texte, rebind suivi d'un refus puis reprise, renommage/SaveAs, cellule et racine supprimées, activation d'un groupe mixte, permanence globale et déconnexion host sous protection. Le sink final avec interface COM importée est inclus. Seule l'instance utilisateur antérieure restait ouverte après terminaison.
- Contrôle natif du complément après ces modifications : enregistrement temporaire effectué puis supprimé ; échec à l'ouverture du seul classeur jetable, `Owned test workbook did not finish opening.` Les contrôles Status/ruban/fenêtres n'ont donc pas été atteints. Cause non établie ; ce résultat ne confirme pas le chargement natif de la nouvelle intégration. Processus détenu arrêté et clés propres au test retirées ; instance utilisateur antérieure préservée. La livraison suivante reste en brouillon.

- Original observé dans une instance jetable : retrait et permanence décoche l'arbre en gardant la fenêtre ouverte ; permanence conservée après fermeture puis nouvelle analyse/retrait. Lecture COM finale : Z!B4, couleur directe 15652797, zéro règle CF, formule et valeur intactes. Pas de sauvegarde/réouverture du classeur original. Processus détenu nettoyé, fichier jetable inchangé.
- Moteur pur : **1 938 assertions** sur piles, ordre, transparence, propriétés indépendantes, permanence, abandon et snapshots. Transactions : **226 assertions**, dont écriture modifiant avant de lever, rollback partiel, réparation/reprise, libération, abandon vérifié d'une référence morte et reçu conservé malgré une erreur de libération.
- Port Excel : **58 contrôles réels**, dont automatic/none, thème/tint/motif pour rollback, texte enrichi homogène, refus avant mutation des cas non couverts, CF utilisateur préservée, SaveAs/renommage et références déplacées/supprimées. La suppression de la cellule référencée a produit HRESULT 0x800A01A8.
- Session Excel : **74 contrôles réels**, deux bridges et plusieurs classeurs, superpositions, permanence masquée et par propriété, modèles détachés, retrait ciblé, édition externe préservée, activation après déplacement et permanence, pré-liaison d'une cellule jamais activée, suppression et réouverture homonyme. Échec injecté après un vrai setter COM puis rollback refusé ; fermeture du classeur concerné, reçu d'abandon et reprise dans un autre classeur vérifiés.
- Revue indépendante : corrections de cible déplacée, réactivation indésirable d'une édition manuelle et réparation définitivement bloquée après disparition de la référence. Rerevue des corrections sans constat restant ; première liaison des cellules exige encore le raccordement explicite du modèle frais à BindDirectHighlightTargets.
- Première exécution complète `test.ps1 -Excel` : **76 groupes réussis, 1 échec**, COLUMN() retournant 5 au lieu de 2 dans le module existant d'évaluation. Les deux nouveaux modules COM passent. Diagnostic borné : 800 sondes ciblées sans reproduction, une séquence capture/opérations réussie puis une autre interrompue par Range.Activate dans la préparation du test, avant Evaluate. Cause du résultat COLUMN initial non établie ; aucune modification spéculative du code d'évaluation.
- Tests de contexte renforcés : cellule et propriétaire journalisés avant/après les six premières évaluations ; 50 contrôles ROW()*100+COLUMN()==502 ajoutés aux changements de classeur/sélection, soit **550 contrôles de contexte/restauration**. Exécution finale `test.ps1 -Excel` : **77 groupes réussis, 0 échec**, dont les 58 et 74 contrôles des nouveaux modules COM. Cette réussite ne démontre pas l'absence de l'intermittence précédente.

## Livraison 0.6 — 15 octobre 2026

- Ordre capturé par Worksheet.Index, y compris indices non contigus ; snapshot immuable transmis aux flux. Fallback explicite pour les fixtures sans index. Les blocs restent dans l'ordre d'émission au lieu d'être retriés par adresse. **895 assertions de blocs/ordre**, dont un grand rectangle démarrant plus haut mais émis après un petit bloc plus bas.
- **628 assertions d'arbre** : deux vues et deux regroupements, niveaux catégorie/feuille/bloc, feuilles cachées, priorité centrale, ordre des catégories et des feuilles, doubles rôles, clés stables, listes immuables et rejet de blocs étrangers ou répétés.
- **62 assertions de préférences** : quatre combinaisons indépendantes, remplacement, création du dossier parent, XML strict borné à 4096 octets, DTD et contenu étranger refusés ; destination verrouillée conservée exactement et temporaires nettoyés avant nouvelle tentative.
- **71 contrôles WinForms sans Excel** : état mixte des parents, ciblage d'une seule feuille, synchronisation arbre/grille, détail B2 conservé sans saut vers B1, curseur retrouvé après regroupement et changement de vue, Espace/Entrée, réinitialisation et retour. Le test de budget a été précisé après la suite complète : feuille cachée contenant directement plusieurs blocs, pour reproduire exactement le défaut corrigé.
- Suite `test.ps1 -Excel` : **73 groupes réussis, 0 échec**, dont 73 contrôles de fenêtres Excel, capture de feuilles dans un ordre contraire à l'alphabet, 95 contrôles de couches et 500 assertions de restauration. Compilation finale application/DLL 0.6.0.0 sans avertissement et **55 groupes sans Excel réussis**, incluant le test de budget renforcé.
- Complément final depuis PowerShell 7 MTA : chargement, callback du ruban, deux explorateurs sur le thread UI Excel, retour et fermeture globale réussis. Enregistrement temporaire supprimé. Aucun nouveau contrôle du client STA en 0.6.
- Interface native contrôlée à **1160 × 820 et 940 × 640** : activation d'une catégorie, détail B2 conservé après décochage, états mixtes arbre/grille, Espace/Entrée, flèches et molette ; regroupements des deux vues indépendants et curseur retrouvé au retour. Au minimum, l'arbre ne montre que 1 à 2 lignes selon la barre horizontale et les longs textes sont tronqués ; défilement et commandes restent accessibles. Processus de démonstration fermé, aucun Excel ni préférence réelle modifié ; réglages de contrôle isolés dans `work/`.
- Revue indépendante : deux défauts corrigés, signalement de troncature sous une feuille cachée et parcours quadratique du focus pour les libellés de frontière. Calcul du contexte central une seule fois par reconstruction ; revue des corrections sans constat restant. Les états visuels utilisent StateImageIndex avec CheckBoxes désactivé, conformément au [contrat WinForms documenté par Microsoft](https://learn.microsoft.com/dotnet/api/system.windows.forms.treenode.stateimageindex).

Les vérifications propres à Clair ne prouvent pas l'identité runtime avec Arixcel. Les lots internes des collections originales, les libellés R1C1, le périmètre multi-classeurs, les couches de formats directs et les commandes globales/permanentes restent ouverts, avec les autres exigences du [registre de parité](PARITY.md).

## Livraison 0.5 — 15 octobre 2026

- Scanner de rectangles original identifié statiquement dans Arixcel.Common.dll. Segments horizontaux ancrés sur leur premier point ; prolongement vertical seulement à largeur/X identiques, avec comparaison à l'origine du rectangle. R1C1 comparé uniquement pour deux formules dans la vue rôles. Preuves, cas non transitifs et limites dans [FLOW-EVIDENCE.md](FLOW-EVIDENCE.md).
- Nouveau corpus indépendant : **753 assertions** sur rectangles, L/T/escaliers, trous, largeurs différentes, feuilles distinctes, limites de grille, copies R1C1, casse, données manquantes, ancrages non transitifs, doubles rôles et activation ciblée d'un bloc. Un invariant de couverture vérifie chaque cellule par rôle exactement une fois et l'absence de trous, sans recopier l'algorithme.
- Suite de régression `test.ps1 -Excel` : **69 groupes réussis, 0 échec**, dont 52 sans Excel, 95 contrôles de couches et 500 assertions de restauration. Cette exécution précède le dernier ajustement d'état initial/reprise ; la session de fenêtres concernée a ensuite été rejouée séparément avec les cas supplémentaires ci-dessous.
- **71 contrôles de fenêtres Excel** sur le code ajusté : première analyse sans couleur, activation de tout un bloc, détail individuel, retour à un bloc mixte, nouvelles cellules non sélectionnées après actualisation, extension puis réduction d'un bloc avec remappage de l'historique, retrait puis réactivation. Les contrôles précédents de doubles rôles, fenêtres indépendantes, SaveAs et carte préservée passent aussi.
- Revue indépendante en lecture seule du diff, puis du remappage et de l'état initial : aucun constat actionnable restant. Aucun test runtime du logiciel original n'a été exécuté.
- Interface native contrôlée à **1160 × 820 et 940 × 640** : démonstration 7 rôles/0 actif/5 blocs, groupe Entrées activé puis détail individuel, état mixte au retour aux blocs et activation de tout le bloc. Double-clic sur son texte : exploration de sa première cellule ; ancien parcours aval et double-clic de sa première colonne également vérifiés. Au minimum, commandes Flux accessibles, listes défilantes ; une seule ligne de groupe visible simultanément. Parcours amont non contrôlé dans cette séance. Processus de contrôle fermé, aucun Excel ni réglage persistant modifié.
- Compilation finale application/DLL 0.5.0.0 sans avertissement, puis **52 groupes sans Excel réussis**. Test final du complément depuis PowerShell 7 MTA : chargement, ruban, deux explorateurs sur le thread UI Excel, retour et fermeture globale réussis ; enregistrement temporaire supprimé. Le client STA avait été vérifié en 0.4, sans nouvelle vérification de ce client pour 0.5.

Le regroupement par catégorie/feuille/adresse de Clair conserve un ordre stable propre au projet ; il ne reproduit pas l'ordre d'émission du scanner original. L'arbre hiérarchique par feuille, les commandes globales de retrait/permanence, la progression et le graphe complet demeurent ouverts. La version ne revendique pas l'identité exhaustive avec Arixcel.

## Livraison 0.4 — 15 octobre 2026

- Compilation application et DLL : réussite sans avertissement. Les tests purs comprennent 51 groupes, dont 182 assertions de palette flux et 110 de composition, en plus des corpus 0.3 conservés.
- Suite finale `test.ps1 -Excel` : **68 groupes réussis, 0 échec**, dont **110 assertions de composition** (les 67 initiales puis activation ciblée sans réordonner les autres cellules).
- **95 contrôles de couleurs dans Excel 16.0** : inventaire strict des couches, découverte après SaveAs/renommage, retrait de plusieurs couches avec prévalidation globale, maintien des règles utilisateur et des propriétés effectives. Échec injecté lors de la deuxième suppression : la première règle supprimée est recréée dans l'ordre relatif correct, sans renuméroter les règles utilisateur.
- Plages disjointes : 50 zones espacées sont construites via des objets Range et `Application.Union`. L'échec injecté pendant leur retrait vérifie aussi la restauration de cette règle ; l'adresse de sauvegarde ne rajoute pas de `$` superflus.
- Palette carte et flux : stockage strict et atomique mutualisé, schéma carte v1 conservé. Éditeur partagé ; import et défaut remplacent les styles dans le dictionnaire référencé par la sauvegarde, sans changer la palette active si le dialogue est annulé.
- **36 contrôles de fenêtres de flux** : activation par cellule/groupe et état mixte, fond/police sur une cellule à deux rôles, fenêtres indépendantes, fermeture après SaveAs, connexion périmée suivie d'une connexion encore valide, conservation puis retrait ultérieur, carte préservée. Une feuille protégée de notes sans rôle ne bloque pas les couleurs d'une autre feuille.
- Restauration de sélection : chemin plage via `Application.Goto`, chemin forme conservé. 50 cycles entre deux classeurs et sélections multizones, soit **500 assertions** dans la suite finale ; deux séquences isolées supplémentaires ont été exécutées en STA. Fenêtre, feuille, sélection, cellule active et flags sont vérifiés.
- Observations CIL dans [FLOW-EVIDENCE.md](FLOW-EVIDENCE.md) : deux propriétés de couleur simultanées, deux rôles possibles par cellule, activation par blocs/parents, ordre des couches, distinction entre fermeture conservant et commande rendant permanent. Ces observations ne constituent pas une exécution du produit original.
- Interface native 0.4 contrôlée à **1160 × 820 et 940 × 640** : analyse de démonstration 7 rôles/7 actifs, case individuelle décochée puis 6 actifs et groupe partiel 2/3. Palette de 8 catégories, fond/police et 8 boutons visibles. À la taille minimale, commandes accessibles et listes défilantes. Processus de contrôle fermé ; aucun classeur ou réglage persistant modifié. Les écritures de couleurs Excel sont couvertes par les tests de callbacks ci-dessus, sans clic physique revendiqué sur un classeur réel.
- Complément : correction du marshaling par `StandardOleMarshalObject`. Avant correction, le test automation créait une fenêtre Clair sur un thread différent du thread UI Excel ; une attente SendMessage bloquait Excel. Rapport local ClrMD/WCT conservé dans `work/diagnostics/hang-0.4.md`. La correction suit le [contrat Microsoft des objets de complément exposés hors processus](https://learn.microsoft.com/en-us/visualstudio/vsto/calling-code-in-vsto-add-ins-from-other-office-solutions?view=visualstudio). Le test contrôle désormais aussi les threads natifs des deux fenêtres.
- Validation finale du complément depuis **PowerShell 7 en MTA** puis **Windows PowerShell 5.1 en STA** : activation, véritable callback du ruban, deux explorateurs sur le thread UI Excel, retour et fermeture globale réussis. Enregistrement temporaire supprimé à la fin des deux exécutions.

Pendant la mise au point, des échecs intermittents ont concerné `Range.Select`, les flags d'application et un contrôle de format numérique après réouverture. La restauration des plages utilise désormais Goto ; les diagnostics par étape et les assertions ont été conservés. Les deux séquences STA et la suite finale passent, sans cause externe établie pour les écarts de flags/format : ces résultats ne prouvent pas leur absence sur tous les hôtes ou timings.

Le regroupement exact en rectangles, la commande Arixcel qui promeut une couche en nouvelle base, la progression et l'annulation d'analyse restent ouverts. Clair active les catégories et cellules et utilise ses propres règles conditionnelles ; aucune identité exhaustive des formats, interactions et écrans avec Arixcel n'est revendiquée.

## Livraison 0.3 — 15 octobre 2026

- Compilation de l'application et de la DLL .NET Framework : réussite sans avertissement.
- `test.ps1 -Excel` : **65 groupes réussis, 0 échec**. Les 49 groupes sans Excel comprennent 682 assertions syntaxiques, 64 de flux, 64 de carte et 308 de palette ; les compteurs de groupes incluent les sessions englobantes.
- Capture réelle de `FormulaR1C1` : références relatives, mixtes et absolues ; copie vers le bas classée correctement. Cellules vides avec validation capturées ; validation d'une formule ne remplace pas sa catégorie structurelle.
- **52 contrôles de couleurs dans Excel 16.0** : remplissage et police effectifs, formats directs, valeurs/formules, règles conditionnelles préexistantes, `StopIfTrue`, ordre relatif, deux couches indépendantes, prévalidation de plusieurs feuilles protégées, remplacement par une couche vide, regroupement de 300 cellules, sauvegarde/réouverture/retrait.
- Excel peut conserver des trous dans ses numéros de priorité après suppression. La version livrée ne renumérote pas les règles utilisateur : elle préserve leur ordre relatif et leur effet. Un essai de renumérotation sur processus jetable a provoqué un arrêt natif Excel ; cette approche a été retirée.
- Palette XML : couleurs par défaut issues des observations statiques, clone profond, catégories exactes, couleurs bornées, DTD interdit, remplacement atomique du fichier. Tests d'import invalide et de remplacement refusé sans perte du fichier précédent.
- Revue indépendante en lecture seule : correction des faux ajouts/suppressions dus aux cellules vides validées, et création du dossier de réglages lors du premier enregistrement. Revue finale de la couche de couleurs : aucun constat restant.
- Interface native 0.3 contrôlée à 1160 × 820 (échelle Windows 200 %) : commandes de carte visibles, 12 catégories et 8 boutons de palette accessibles, sélecteur Windows de couleur fonctionnel. Processus de contrôle fermé. Taille minimale et clics d'application/retrait via MainForm non vérifiés ; les opérations Excel correspondantes sont couvertes par les tests d'API ci-dessus.
- Complément 0.3 : `test-addin.ps1 -SkipBuild` réussi dans Excel ; chargement COM, callback du ruban, deux explorateurs, retour et fermeture globale. Enregistrement temporaire supprimé et absence de la clé Clair vérifiée après le test.

La restauration après un échec survenant **après suppression partielle** des anciennes règles n'a pas été injectée en test. Le code tente la restauration et signale explicitement toute restauration incomplète. Les tests ne prouvent pas une identité binaire du classeur ni une fidélité exhaustive avec Arixcel. Options d'analyse de carte, couverture des références complexes, surbrillance de flux, comparaison avancée et utilitaires restent ouverts dans [PARITY.md](PARITY.md).

## Livraison 0.2 — 15 octobre 2026

Les résultats de cette section remplacent le diagnostic COM de la base 0.1 conservé plus bas.

- Compilation C# .NET Framework, application et DLL : réussite.
- 46 groupes de tests sans Excel, comprenant **682 assertions syntaxiques** et **64 assertions de flux** : réussite.
- 59 groupes avec Excel 16.0, dont deux sessions isolées : réussite. Les compteurs incluent les groupes englobants, pas uniquement des assertions unitaires.
- Évaluation : langue anglaise explicite sous Windows français, contexte de feuille et cellule (`ROW`, `COLUMN`), plages, matrices, erreurs, références structurées ordinaires et en-tête contenant `!`.
- Édition/annulation : contrôle des formules périmées et changements de types, constantes texte commençant par `=`, valeurs, formules déversées, absence de propagation dans une colonne de tableau, format numérique et respect d'un changement de format ultérieur.
- Contexte UI Excel : sélection multizone de cellules et cellule active restaurées ; sélection d'une forme restaurée. Les tests ne modifient que leurs classeurs jetables.
- Complément : activation réelle dans un processus Excel `/x` identifié, `OnConnection`, objet d'automatisation, véritable callback `RibbonOnLoad`, XML du ruban, deux fenêtres simultanées, retour et fermeture globale. Les actions sont invoquées via IDispatch sur les callbacks ; aucun clic physique n'est revendiqué.
- Test du complément exécuté avec Windows PowerShell 5.1 et PowerShell 7.6.5 ; registre temporaire et processus propres au test nettoyés.
- Interface réellement contrôlée dans sa fenêtre : sélections des trois menus de flux visibles, 7 rôles dans le modèle de démonstration ; avec zone centrale limitée à la feuille Résultat, 3 antécédents extérieurs et 2 entrées intérieures. La [capture de fenêtre](flow-preview.png) provient de ce contrôle. Le processus lancé a été fermé.

Le défaut COM initial provenait d'une signature `ref Array` qui ne spécifiait pas le contrat natif `SAFEARRAY(VARIANT)* [In]` des callbacks `IDTExtensibility2`. Le fait que l'activation COM directe fonctionnait ne prouvait pas le chargement dans Excel.

Revue ciblée indépendante et corrections : qualification des références structurées avant `Worksheet.Evaluate`, conservation des en-têtes contenant `!`, restauration des sélections de formes, détection de conflits nombre/texte, restauration du format généré automatiquement par `DATE`. Le nouvel arbre conserve les positions du texte ; les variables LET sont encapsulées dans leur portée pour évaluation. Le graphe de dépendances lexical reste distinct.

L'équivalence Arixcel reste **non démontrée** : voir [PARITY.md](PARITY.md), notamment branches effectives, recherches détaillées, paramètres LAMBDA, carte sur grille, alignement de comparaisons et utilitaires. Les versions/architectures Excel non testées ne sont pas couvertes par les résultats locaux.

## Historique : résultats locaux 0.1

Windows 11, .NET Framework, Excel de bureau version 16.0. Compilation sans dépendances NuGet.

- 44 tests du moteur : réussite.
- 56 vérifications avec session Excel isolée : réussite. Le compteur inclut le contrôle englobant de session.
- Interface de démonstration rendue et inspectée visuellement : texte, arbre et détails lisibles à 1160 × 820.
- Activation COM directe de la classe depuis Windows PowerShell : réussite, état `Disconnected` attendu hors hôte.
- Découverte via `Excel.COMAddIns` après enregistrement temporaire : échec, `DISP_E_BADINDEX`. Le complément n'apparaît pas dans la collection. Les clés temporaires ont été supprimées après chaque essai.
- Aucun test de fidélité comportementale complète contre Arixcel.

## Revue review-agent

Relecture indépendante en lecture seule de tous les fichiers C# initiaux et tests, suivie d'une revue des corrections et du complément/installateur.

Défauts relevés et corrigés :

1. Plages avec qualification de feuille répétée : conservation des cellules intérieures.
2. Intersections directes : calcul de l'intersection au lieu de l'union.
3. Noms relatifs : détection et avertissement explicite au lieu d'une résolution figée sur la sélection Excel.
4. Noms externes : aucun repli vers un nom local homonyme.
5. Changement de référence de comparaison : invalidation des résultats précédents.
6. Références relatives de lignes/colonnes entières : détection étendue.
7. Feuilles nommées `RC` : exclusion du nom de feuille lors de la détection R1C1.

Dernière revue ciblée : aucun constat restant dans les corrections examinées. Cette revue ne remplace pas la validation runtime du complément COM, qui demeure en échec.

Le test d'intégration a également détecté une erreur de classification des valeurs `VT_ERROR` d'Excel. La lecture distingue désormais les erreurs des nombres et affiche notamment `#DIV/0!`.

## Reproduction

Exécuter `test.ps1` pour les tests sans Excel, `test.ps1 -Excel` pour l'intégration et `test-addin.ps1` pour le chargement du complément. La CI GitHub n'exécute que les tests sans Excel ; les tests Office sont locaux.

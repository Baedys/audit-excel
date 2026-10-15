# Revue et validation

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

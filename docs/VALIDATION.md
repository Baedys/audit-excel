# Revue et validation

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

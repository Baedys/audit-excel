# Analyse du logiciel de référence

Date : 15 octobre 2026. Cible fournie : `Arixcel_Explorer_v9.0.9623 (1).msi`.

## Méthode et niveau de preuve

Le MSI a été lu avec l'API Windows Installer, puis extrait avec 7-Zip sans lancer son installation. 23 fichiers embarqués ont été matérialisés dans un dossier de travail local. Ils ne font pas partie de ce dépôt.

REA a inspecté les métadonnées de l'assembly principal sans exécuter son code. L'inventaire détaillé des membres REA a atteint sa limite de représentation de 64 Mio. Le relais a été assuré par `System.Reflection.Metadata` / `PEReader`, avec lecture du XML de ruban et décodage ciblé des instructions CIL. Aucun chargement CLR du logiciel original n'a été nécessaire pour cette analyse.

Les composants extraits, les ressources et les relevés CIL restent locaux. Ce document résume les observations ; il ne distribue pas d'implémentation issue du logiciel original.

## Identité observée

- Produit MSI : Arixcel Explorer 9.0, version `9.0.9623.19616`.
- Assembly principal : `ArixcelExplorer.dll`, 1 524 800 octets.
- SHA-256 : `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`.
- Framework cible : `.NETFramework,Version=v4.8`.
- Implémentation : CIL, architecture managée AnyCPU.
- Métadonnées : 670 types, 3 748 méthodes, 32 ressources embarquées. Ces nombres incluent les types générés et interop.
- Preuve REA : `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9`.
- Inventaire membres REA incomplet : `ev_e02f1b6fd380f4eb45bf7081600b9d566f771dbf20ee1fa9e62b546530bc1730`.

## Composition

Le paquet contient le complément VSTO, une extension XLAM avec projet VBA, la DLL principale, deux bibliothèques Arixcel, une bibliothèque de contrôles d'arbre, les utilitaires VSTO, log4net et plusieurs bibliothèques de compatibilité .NET. Les références WinForms et Drawing confirment l'emploi d'une interface native Windows.

L'XLAM contient un projet VBA et ses signatures. Son code VBA n'a pas été analysé. Son rôle exact reste donc inconnu.

## Fonctions et chemins observés

| Fonction | Éléments observés | Ce que cela établit |
| --- | --- | --- |
| Antécédents | Commande du ruban, modèle d'exploration, méthodes de résolution et de valeur | Présence d'un explorateur structuré, allant au-delà d'une simple liste de cellules |
| Dépendants | Collecteur appelant `ShowPrecedents`, `ShowDependents` et `NavigateArrow` ; activation de feuilles et fenêtres | Utilisation des mécanismes de traçage d'Excel dans le chemin étudié |
| Édition de formule | Méthodes d'édition, acceptation, annulation et historique dans le formulaire d'exploration | Fonction d'édition prévue ; comportement complet non exécuté ni vérifié |
| Carte des formules | Analyse du contenu, détection d'erreurs, parsing de références externes/interfeuilles ; sauvegarde/restauration de formats conditionnels | Classification plus détaillée que les seuls types de valeurs |
| Flux de calcul | Parsing des formules et collecte de références ; catégories et regroupement par feuille | Graphe d'entrées/flux/sorties avec fonctions de surbrillance |
| Comparaison | Sélection de plage/feuille/classeur, contrôle de tailles égales dans un chemin, options de tolérance, rapport et alignement | Comparaison multi-échelle et traitement d'alignement, non reproduits intégralement |
| Utilitaires | Actions de format, navigation, collage, remplissage, transformation et raccourcis | Ensemble d'outils supplémentaires distinct du moteur d'audit |

Les appels ciblés ci-dessus sont observés dans les corps CIL. Les capacités déduites seulement de noms de types/méthodes sont indiquées comme telles. L'analyse ne prouve pas une équivalence de comportement de Clair avec le produit original.

## Choix de reconstruction

Clair utilise un moteur original d'extraction de références et des captures en mémoire. Le prototype ne reproduit pas l'algorithme du collecteur original. Il évite notamment de modifier les flèches de traçage et les formats du classeur.

La lecture complète des dépendances ne peut pas reposer uniquement sur `Range.DirectPrecedents` : Microsoft indique que cette propriété se limite à la feuille active et ne suit pas les références distantes. Voir [documentation officielle](https://learn.microsoft.com/en-us/office/vba/api/excel.range.directprecedents).

Le point d'entrée COM est fondé sur les interfaces Office documentées : [IDTExtensibility2](https://learn.microsoft.com/en-us/dotnet/api/extensibility.idtextensibility2?view=visualstudiosdk-2022), [IRibbonExtensibility](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.iribbonextensibility?view=office-pia) et [COMAddIns.Update](https://learn.microsoft.com/en-us/office/vba/api/office.comaddins.update). Le test local de découverte du complément reste en échec ; aucune réussite de chargement n'est revendiquée.

## Questions encore ouvertes

- Évaluation exacte des sous-expressions, recherches, tableaux dynamiques, noms locaux LET/LAMBDA.
- Algorithmes de comparaison et d'alignement de blocs.
- Comportement runtime et rôle de l'XLAM.
- Couverture des versions et architectures d'Excel.
- Fidélité de l'interface, des raccourcis et du cycle d'annulation.

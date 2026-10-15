# Carte de formules : observations et critères

Relevé statique du 15 octobre 2026, Arixcel Explorer `9.0.9623.19616`. Assembly SHA-256 : `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`. Identité établie par REA, preuve `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9`. Les corps ciblés ont été lus avec `PEReader`/`System.Reflection.Metadata`, sans charger ou exécuter l'assembly original. Les scripts et le binaire restent locaux dans `work/`.

La [documentation officielle](https://www.arixcel.com/explorer/functions/map) décrit une carte temporaire distinguant formules uniques et copies, une analyse optionnelle des références des formules uniques et une restauration des couleurs même après sauvegarde/réouverture. Les détails ci-dessous proviennent du binaire fourni, et non d'une déduction à partir des couleurs de captures d'écran.

## Priorités effectivement observées

`Arixcel.Explorer.Features.Maps.Analyser.GetCellTypeMain`, RVA `113232` :

| Emplacement CIL | Observation |
| --- | --- |
| `IL_0000`–`IL_0075` | Distingue absence de contenu, texte et constantes. Doubles, décimaux, booléens et erreurs constantes relèvent de la catégorie Number. |
| `IL_0076`–`IL_00b6` | Lecture du voisin immédiat **au-dessus** ; aucun voisin si la cellule est sur la borne haute de la zone lue. |
| `IL_00b7`–`IL_00f7` | Lecture du voisin immédiat **à gauche** ; aucun voisin si la cellule est sur la borne gauche. |
| `IL_00f9`–`IL_0126` | Égalité exacte des chaînes FormulaR1C1 : deux voisins égaux → copie dans deux directions ; gauche seul → copie horizontale ; haut seul → copie verticale. Aucun test droite/bas. |
| `IL_0127`–`IL_0143` | Si l'analyse des formules uniques est désactivée : catégorie UniqueFormula. Sinon, valeur initiale UniqueFormulaNoRefs avant inspection des éléments. |
| `IL_0165`–`IL_0197` | Référence externe → résultat immédiat UniqueFormulaExternalRefs. Référence interfeuille → UniqueFormulaSheetRefs. Une référence locale remplace seulement UniqueFormulaNoRefs par UniqueFormula. Priorité externe > interfeuille > locale > aucune. |

Ainsi, dans un bloc recopié, la cellule en haut à gauche reste « unique » ; les autres révèlent la direction de copie. Une comparaison avec les quatre voisins modifierait ce comportement. Les trous interrompent les chaînes. Une formule qui calcule une erreur conserve son rôle structurel : il n'existe pas de catégorie spécifique « formule en erreur » dans cet enum.

`Analyser.GetCellType`, `IL_000e`–`IL_0023` : la validation de données remplace seulement les catégories Empty, Text et Number, si l'analyse de validation est active. Elle ne remplace aucune catégorie de formule. `MapColors.get_AnalyseDataValidation` compare la couleur de DataValidation à `Color.Empty` : l'analyse de validation dépend donc de l'activation de cette couleur.

Ces éléments sont des observations de branches CIL, pas une preuve d'exécution ni une garantie sur les versions ultérieures.

## Palette par défaut observée

Le constructeur statique du type imbriqué nommé `FormulaMap` charge les couleurs suivantes par `ColorTranslator.FromHtml`. `MapColors..cctor` les associe aux catégories. Les couleurs personnalisées peuvent remplacer ces valeurs au chargement des paramètres.

| Catégorie | Valeur par défaut | Affectation CIL du champ |
| --- | --- | --- |
| UniqueFormula | `#FFC000` | `IL_000a` |
| UniqueFormulaNoRefs | `#00FFFF` | `IL_0019` |
| UniqueFormulaSheetRefs | `#F09100` | `IL_0028` |
| UniqueFormulaExternalRefs | `#00FF80` | `IL_0037` |
| FormulaCopiedAcross | `#B8CCE4` | `IL_0046` |
| FormulaCopiedDown | `#C3D69B` | `IL_0055` |
| FormulaCopiedDownAndAcross | `#C4BD97` | `IL_0064` |
| DataValidation | `#DBEEF3` | `IL_0073` |
| Text | `Color.Empty` | `IL_007d` |
| Number | `#FFFF71` | `IL_008c` |
| Empty | `Color.Empty` | `IL_0096` |

`Color.Empty` signifie ici absence de couleur de carte pour la catégorie, pas effacement du format existant. `MapColors..cctor`, `IL_0177`–`IL_0183`, charge l'option AnalyseUFs avec une valeur par défaut vraie. Le stockage précis de la restauration, les limites de la zone effectivement lue et les interactions avec les formats conditionnels restent à vérifier en exécution.

## Implémentation originale de Clair

`FormulaMap.Analyze(Book)` fournit des `MapCell` avec cellule, catégorie, raison et avertissements, et un résultat global `Cells`, `Warnings`, `Incomplete`. L'implémentation reproduit les priorités et directions ci-dessus ; elle utilise les chaînes R1C1 réellement capturées dans Excel. Les comparaisons sont ordinales exactes : références mixtes/absolues, espaces et casse des chaînes littérales ne sont pas normalisés arbitrairement.

Onze catégories correspondent à l'enum observé. La douzième, `Unknown`, est propre à Clair : R1C1 absente/invalide pour la cellule ou un voisin pertinent, syntaxe non reconnue, noms ou références insuffisamment résolus. Cette catégorie empêche d'affirmer « formule unique sans référence » à partir de données manquantes.

Les copies prouvées par R1C1 peuvent conserver une catégorie de copie tout en signalant une analyse syntaxique incomplète. Une référence externe explicite permet la catégorie externe, de priorité maximale, même si sa cible ne peut être résolue ; l'avertissement de résolution demeure. Les chaînes littérales et les noms de colonnes structurées contenant `!` ne sont pas traités comme des liens externes. La recherche de liens externes par noms définis respecte les variables locales LET/LAMBDA : une variable masquant un nom du classeur ne crée pas de faux lien externe. Ces formules restent signalées incomplètes par le moteur de références actuel.

Le moteur classe les cellules présentes dans `Book.Cells`. Il ne fabrique ni les cellules vides absentes de la capture, ni leur validation. L'appelant doit capturer les cellules vides avec validation nécessaires au périmètre. Toute capture plafonnée doit fournir un avertissement ; celui-ci est propagé dans `Incomplete`.

Pour l'instant, le moteur applique l'analyse des formules uniques et de validation ; la désactivation de ces options n'est pas exposée par cette API. Les détails des noms complexes, références 3D, tables, fonctions dynamiques et qualification explicite de la feuille courante ne sont pas considérés comme reproduits intégralement. Le parseur de Clair est indépendant de celui du produit original et ses limitations restent visibles.

## Validation

`FormulaMapTests.Run()` couvre directions de copie et cellule source, trous, séparation des feuilles, limites A1/XFD1048576, R1C1 relatives et absolues, casse des chaînes, références locales/interfeuilles/externes, chaînes littérales trompeuses, tableaux structurés non résolus, syntaxe invalide, données R1C1 absentes, priorités validation/constantes/formules, erreurs calculées et chaînes de noms.

Compilation .NET Framework et **64 assertions réussies** lors de cette livraison. Ces tests vérifient le moteur pur. Ils ne prouvent ni l'application/restauration des formats dans Excel, ni la persistance après réouverture, ni une comparaison différentielle complète avec Arixcel. Ces validations restent nécessaires avant de déclarer la carte identique.

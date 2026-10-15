# Flux de calcul : preuves de comportement

Relevé du 15 octobre 2026, pour **Arixcel Explorer 9.0.9623.19616**. Binaire local : `work/arixcel-extracted/ArixcelExplorer.dll`, SHA-256 vérifié `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`. Identité REA : `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9`.

Méthode : réutilisation de l'inventaire existant puis lecture ciblée des corps CIL et constantes avec `PEReader` / `System.Reflection.Metadata`. Le lecteur de `work/inspect-map-palette.ps1` a été adapté **en mémoire**, sans lancer, installer ou charger dans le CLR le logiciel original. Aucun inventaire REA géant des membres n'a été répété. Les noms de méthodes seuls ne sont pas traités comme preuve de comportement.

Les observations sont statiques, limitées aux chemins ci-dessous. Elles ne prouvent pas le résultat visuel dans toutes les versions d'Excel. La [documentation officielle du flux](https://www.arixcel.com/explorer/functions/flow), lue dans son module public `chunk-C0OacgCm.js`, décrit le choix d'un périmètre et d'une zone centrale, deux vues, des groupes à colorer/décolorer, la fermeture normale avec retrait, la fermeture conservant les couleurs, leur retrait ultérieur et leur conservation comme format du classeur. Elle exclut explicitement les dépendances via OFFSET/INDIRECT. Ces descriptions actuelles corroborent certaines branches du binaire, sans remplacer leur vérification runtime.

## 1. Catégories et palette par défaut

Enum `Arixcel.Explorer.Features.CalcFlow.CellType` : Empty=0, Inputs=1, Calculations=2, Outputs=3, OrphanFormulas=4, OutsidePrecedents=5, InsideInflows=6, InsideOutflows=7, OutsideDependents=8. Enum ViewMode : InputsOutputs=0, InflowsOutflows=1.

`FlowColors..cctor`, `IL_0000`–`IL_00af`, construit deux dictionnaires distincts :

| Catégorie | Remplissage | Police |
| --- | --- | --- |
| Inputs | `Color.Empty` | `#0000FF` |
| Calculations | `#F2F2F2` | aucune entrée par défaut |
| Outputs | `#CCFF99` | aucune entrée par défaut |
| OrphanFormulas | `#FF9999` | aucune entrée par défaut |
| OutsidePrecedents | `#DEEBF6` | aucune entrée par défaut |
| InsideInflows | `Color.Empty` | `#188A18` |
| InsideOutflows | `#BDD7EE` | aucune entrée par défaut |
| OutsideDependents | `#E2EFD9` | aucune entrée par défaut |

`FlowColors.GetColor(type, property)` sélectionne Colors pour property=0, FontColors pour property=1, et lève une exception pour une autre valeur (`IL_0000`–`IL_0028`). Une clé absente laisse la couleur à sa valeur par défaut via TryGetValue. Ce relevé ne montre pas un booléen Enabled dans la palette : la sélection des blocs et leur état de surbrillance sont gérés séparément.

**Remplissage et police sont combinables**, pas deux modes exclusifs d'analyse. `Area.HighlightBlock`, RVA **301424**, lit successivement GetColor(type,0) (`IL_0039`–`IL_003f`) et GetColor(type,1) (`IL_009c`–`IL_00a2`). Chaque propriété non transparente est appliquée séparément ; une propriété transparente est sautée. L'entrée Inputs produit donc normalement une police bleue sans changer le remplissage, et InsideInflows une police verte sans changer le remplissage. `Area.UpdateColor`, RVA 301644, peut modifier, ajouter ou retirer une propriété seule selon la transition transparent/non transparent.

Les valeurs ci-dessus sont des défauts statiques. Les paramètres utilisateur peuvent les remplacer. L'absence de couleur ne signifie pas effacer le format déjà présent.

## 2. Doubles rôles : lire les appels, pas seulement le helper

`CellTypeHelpers.AsCellType`, RVA **302348**, projette les flags reçus vers un seul CellType. En vue InputsOutputs, il vérifie successivement Calculations, Outputs, Inputs, OrphanFormulas, puis Empty. En vue InflowsOutflows, l'ordre interne est OutsidePrecedents, InsideInflows, InsideOutflows, OutsideDependents, puis Empty.

**Cet ordre ne constitue pas une priorité globale qui élimine un double rôle.** L'appelant `Area.BuildBlocks`, RVA **300580**, prépare en vue InflowsOutflows **deux passes** (`IL_0133`–`IL_0142`) :

| Passe | Masque | Flags conservés | Catégories produites |
| --- | --- | --- | --- |
| Première | 80 = 16 + 64 | InsideInflows, OutsideInflows | InsideInflows, OutsideDependents |
| Seconde | 40 = 8 + 32 | OutsideOutflows, InsideOutflows | OutsidePrecedents, InsideOutflows |

Les noms internes OutsideInflows/OutsideOutflows ne sont pas les libellés visibles. La correspondance ci-dessus vient des branches de AsCellType. Le lambda `<BuildBlocks>b__2`, RVA **384424**, construit le bloc avec **les flags complets**, mais calcule son CellType à partir de **flags & mask** (`IL_0048`–`IL_0061`).

Conséquence prouvée par ce chemin : une cellule centrale qui reçoit de l'extérieur et fournit à l'extérieur peut appartenir à deux blocs, InsideInflows et InsideOutflows. Une cellule externe peut également relever de OutsidePrecedents et OutsideDependents. Ne pas supprimer ces doublons de rôle au nom de l'ordre du helper. Avec les défauts, le double rôle interne combine naturellement police verte et remplissage bleu. Cela ne prouve pas encore l'ordre visuel de deux couleurs personnalisées visant la même propriété.

La vue frontières n'est pas construite lorsque la zone centrale n'est pas plus petite que le périmètre : `Area.BuildBlocks`, `IL_00d0`–`IL_0100`, saute ce mode ; `ResultsForm.ResetTree`, RVA 292648, rejette sa sélection avec un message explicite (`IL_0006`–`IL_0043`).

## 3. Blocs, groupes et activation

Le regroupement précède l'activation de la surbrillance. `Area.BuildBlocks` appelle un processeur de rectangles avec trois callbacks. Les callbacks observés établissent :

- `<BuildBlocks>b__0`, RVA **383988** : admission selon le masque de la passe.
- `<BuildBlocks>b__1`, RVA **384100** : compatibilité si flags masqués identiques ou même catégorie projetée ; si les deux valeurs masquées conservent le flag Formula, comparaison supplémentaire de FormulaR1C1 (`IL_0078`–`IL_012d`).
- `<BuildBlocks>b__2`, RVA **384424** : création d'un Block avec plage rectangulaire, flags complets et catégorie calculée pour la passe.

Le masque de la vue InputsOutputs est -1, donc conserve Formula ; les masques frontières 80/40 ne le conservent pas. **L'exigence d'égalité R1C1 du callback ne s'applique donc pas aux blocs de la vue frontières.** Une constante et une formule de même rôle peuvent être compatibles : la comparaison R1C1 s'applique seulement lorsque les deux points sont des formules.

### Processeur de rectangles, relevé complémentaire 0.5

`Arixcel.Common.dll`, SHA-256 **425eb00972ba6d5aaa45ee6e8a4369e1fa40ff1ea2f002da2522893c0711eb01**, contient `AreaScanner.Process`, RVA **75764**. Lecture CIL ciblée, sans chargement ni exécution de l'assembly original :

- Balayage des lignes puis des colonnes croissantes. Un segment horizontal compare chaque candidate à son **premier point**, pas à sa voisine précédente.
- Un rectangle actif se prolonge seulement si le segment de la ligne suivante a exactement les mêmes X et largeur (`IL_0273`–`IL_029d`). La compatibilité verticale compare uniquement l'origine du rectangle au premier point de ce segment (`IL_02b8`–`IL_02ce`). Les autres colonnes ne sont pas recontrôlées verticalement.
- Une rupture réelle de ligne ou de colonne rompt un bloc de rôle visible. Les chemins de points neutres du scanner générique ne prolongent pas un rôle non vide de Flow, car son callback de compatibilité refuse Empty.
- Les blocs terminés sont émis lorsqu'ils cessent de croître ; les survivants le sont plus tard. Dictionary/HashSet fournissent l'ordre interne, sans tri explicite par adresse. Un rectangle plus haut peut donc apparaître après un petit rectangle plus bas.

Ce n'est **pas** une recherche de rectangles maximaux. Des lignes de largeur 2 puis 3 restent deux blocs, sans extraction d'un sous-rectangle commun de largeur 2. La compatibilité n'est pas transitive : une constante à l'origine peut admettre deux formules R1C1 différentes ; une formule à l'origine n'est pas « reliée » à une autre formule différente par une constante intermédiaire. Verticalement, une constante en première colonne peut également permettre la fusion malgré une autre formule dans les colonnes suivantes.

`FlowBlocks.Build` reconstruit cette géométrie pour les rôles visibles de la capture Clair. Il garde chaque cellule de chaque rôle, y compris les doubles rôles de frontière. Les tests couvrent aussi les ancrages non transitifs et la comparaison verticale limitée à la première colonne. Deux différences restent explicites : Clair trie les résultats par catégorie/feuille/adresse, et sépare deux formules si leur R1C1 manque. Aucun test runtime différentiel n'établit l'ordre de navigation original.

`Model.BuildTree`, RVA **305640**, et son lambda RVA **385220** établissent également des niveaux catégorie, classeur, feuille et bloc. En vue rôles, le réglage GroupBySheet peut placer la feuille avant la catégorie ; en vue frontières, la catégorie reste avant la feuille et le réglage change surtout visibilité et libellés. Les aires suivent Worksheet.Index, avec feuille centrale prioritaire en frontière (lambda RVA **386016**). L'arbre hiérarchique et ces réglages ne sont pas encore reconstruits par la grille Clair.

État initial borné : `Area..ctor`, RVA **300320**, initialise un ensemble de blocs colorés vide ; `BlockItem..ctor`, RVA **315865**, commence décoché sauf si son aire connaît déjà une couleur. Le chemin de modèle frais inspecté ne colorie pas tous les blocs à l'ouverture. Une actualisation peut retrouver des couleurs existantes ; cela ne prouve pas tous les comportements de reprise. Clair 0.5 démarre également sans couleurs et remappe ses choix lors d'une nouvelle analyse, avec tests propres.

`TreeItem.get_DisplayHighlight` / `set_DisplayHighlight`, RVA **314344** / **314352**, utilisent les états 0, 1, 2 ; les branches indiquent retrait, activation et état mixte. Le lambda `<set_DisplayHighlight>b__0`, RVA **388020**, parcourt les descendants ; pour chaque BlockItem, il appelle Area.HighlightBlock si l'état demandé vaut 1, sinon Area.RemoveBlockHighlight (`IL_0025`–`IL_0068`). Les états des parents sont ensuite agrégés, avec passage à l'état mixte si leurs enfants diffèrent. Le callback de progression compte les blocs à changer.

Ce chemin permet donc d'agir sur un bloc ou un parent regroupant plusieurs blocs. Il ne modifie pas les flags de calcul et ne refait pas AsCellType. Désactiver InsideInflows ne transforme pas ce rôle en InsideOutflows : cela retire la couche des blocs concernés, tandis qu'un éventuel bloc InsideOutflows activé reste distinct. La disponibilité de regroupements par feuille dans les deux vues est attestée par les deux réglages séparés lus dans ResultsForm.ResetTree (`IL_0044`–`IL_0077`).

## 4. Ordre des couleurs et fermeture

L'application originale observée repose sur des **couches par propriété**. `Highlighter.Highlight`, RVA 12000, appelle Sheet.AddHighlight. Le bloc de surbrillance, distinct du bloc d'analyse, conserve une liste de couches :

| Méthode | Preuve observée |
| --- | --- |
| `Block.AddLayer`, RVA **321820** | Applique une couleur non transparente si nécessaire puis ajoute la couche à la fin de la liste (`IL_0004`–`IL_0056`). |
| `Block.TopColor`, RVA **321737**, et `FirstNonTransparentBelow`, RVA **321644** | Recherche de la dernière couche non transparente, avec retour à la base. |
| `Block.MoveLayerToTop`, RVA **322608** | Recherche par propriétaire/origine, application éventuelle, retrait puis ajout en fin de liste (`IL_00c2`–`IL_00d5`). |
| `Block.TryToApplyColor`, RVA **323116** | Property=0 : Range.Interior puis ExcelColor.Apply ; property=1 : Range.Font puis Apply. Aucun appel FormatConditions dans ce corps. |

Ainsi, une collision de couleurs personnalisées sur la même propriété dépend des couches et actions d'activation. L'ordre des passes 80 puis 40 n'est pas une preuve suffisante d'une priorité visuelle fixe. Les interactions avec d'autres outils de surbrillance, les formats conditionnels existants et les modifications manuelles demandent des tests runtime.

Fermeture normale : `Manager.FormClosed` délègue au lambda `<FormClosed>b__0`, RVA **387733**. Celui-ci appelle Model.RemoveHighlights seulement si Model.Form n'est pas nul (`IL_0000`–`IL_0022`), puis retire le modèle de la liste si le retrait réussit. Le modèle suit les aires, puis retire les deux propriétés de chaque bloc.

Fermeture conservant les couleurs : `ResultsForm.CloseKeepingHighlights`, RVA **294544**, affecte **Model.Form=null avant Close** (`IL_0000`–`IL_000d`). Le handler de fermeture saute alors la suppression. Ce chemin conserve le modèle et ses surbrillances pour un retrait ultérieur ; il n'en fait pas immédiatement des formats permanents. La [documentation officielle](https://www.arixcel.com/explorer/functions/flow) distingue également ce parcours de la commande Keep Colors.

## 5. Retirer et rendre permanent ne sont pas la même action

`Manager.RemoveAllHighlights(forget)`, RVA **307000**, propage le booléen forget au modèle puis à l'aire. `Area.RemoveAllHighlights`, RVA **302008**, transmet ce booléen aux retraits du remplissage et de la police. `Highlighter.RemoveHighlight`, RVA **12716**, aboutit à `Sheet.RemoveHighlight`, RVA **325044** :

- forget=false appelle **Block.RemoveLayer** (`IL_0020`–`IL_0025`). Le corps, RVA **321968**, détermine la couleur visible après retrait, appelle TryToApplyColor si elle change, puis retire la couche.
- forget=true appelle **Block.ForgetLayer** (`IL_002c`–`IL_002f`). Le corps, RVA **322424**, cherche la couche par propriétaire et origine. Si elle est transparente, il la supprime. Sinon, il la remplace par une couche sans propriétaire/origine gardant sa couleur (`IL_0064`–`IL_0075`), puis retire toutes les couches placées **en dessous** (`RemoveRange(0,index)`, `IL_007a`–`IL_0082`).

ForgetLayer ne contient aucun appel TryToApplyColor : la couleur a déjà été appliquée. Il transforme la couche ciblée en nouvelle base, en abandonnant l'historique sous-jacent. Les couches au-dessus restent présentes. Si la couche rendue permanente n'était pas au sommet, elle peut rester masquée jusqu'au retrait d'une couche supérieure ; elle sera alors la base révélée. Cette sémantique provient du corps, pas du seul nom ForgetLayer.

Cette observation ne justifie **pas** d'affirmer qu'une conversion de règles de mise en forme conditionnelle en formats directs dans Clair est identique. Le chemin original étudié applique directement Interior/Font et gère son historique propre. Une reconstruction utilisant des règles conditionnelles doit valider indépendamment le résultat, le retrait des seules règles qu'elle possède, la préservation des règles préexistantes, les thèmes/patterns/tints et les cas de superposition.

## Incidences pour Clair et validations restantes

Relevé complémentaire de la permanence : `App.Run`, RVA **23132**, transmet true à `Manager.RemoveAllHighlights` après confirmation (`IL_01a8`–`IL_01a9`). Ce Manager traite **tous ses modèles de flux**, y compris ceux dont la fenêtre a été fermée en conservant les couleurs ; il ne filtre pas sur le classeur actif (lambda RVA **387796**). Les fenêtres encore ouvertes restent ouvertes et synchronisent leurs états, via `Model.RemoveHighlights`, RVA **305744**. La commande Clair de retrait reste limitée au classeur chargé : sa portée n'est donc pas identique.

Par propriété, ForgetLayer abandonne aussi les couches inférieures appartenant à d'autres outils ou fenêtres. Les couches supérieures subsistent et peuvent masquer la nouvelle base. `ExcelColor.TryToDerive`, RVA **21760**, conserve explicitement ColorIndex automatic/none, sinon la couleur RGB ; `Apply`, RVA **22544**, réécrit cet index spécial ou cette couleur. ThemeColor, TintAndShade et le motif complet ne sont pas mémorisés dans ce chemin. Cela ne prouve pas les effets exacts des setters Excel sur ces métadonnées. Une commande de permanence ne peut donc pas être remplacée par le maintien indéfini d'une règle conditionnelle Clair.

Pour une API de flux fidèle : conserver les liens et les deux rôles éventuels, les deux propriétés de palette, la distinction entre bloc d'analyse et couche de surbrillance, et les trois intentions « fermer en retirant », « fermer en conservant temporairement », « rendre permanent ». Une sélection de catégories est une première approximation de l'activation par blocs, pas sa preuve d'équivalence. Une ligne unique par cellule peut être un choix de présentation propre à Clair si les deux rôles restent visibles ; elle ne doit pas être décrite comme la structure originale établie.

À vérifier : comparaison runtime des rectangles et des cases initiales, ordre de navigation, fidélité de l'arbre et portée des commandes du ruban, superposition entre vues/outils, persistance de l'historique après réouverture, effets Excel sur thèmes et motifs, erreurs/protection/annulation pendant retrait, et résultat de Keep Colors en présence de plusieurs couches. Aucun test différentiel Arixcel/Clair n'a été exécuté dans ce relevé. L'objectif de parité reste ouvert.

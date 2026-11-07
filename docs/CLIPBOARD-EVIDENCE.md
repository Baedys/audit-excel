# U03 — copie et presse-papiers Excel

Développement du 7 novembre 2026, distinct de la livraison 0.6. Cible : Arixcel Explorer 9.0.9623.19616. Le volet **Utilitaires → Copie et collage** contient la somme copiée et cinq collages spéciaux. La duplication conserve pour le moment un moteur de réécriture indépendant, sans commande de collage raccordée. **U03 reste partiel.**

## Identités et frontières de preuve

Assemblages fournis par l’utilisateur, inspectés par PEReader sans intégrer leurs fichiers au dépôt :

- Explorer : SHA-256 `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`, identité REA `ev_d2f19675e5608e3dce2bcd898fcc1e1121d907a3815b566b5cd009312573fe47`.
- Common : SHA-256 `425eb00972ba6d5aaa45ee6e8a4369e1fa40ff1ea2f002da2522893c0711eb01`, identité REA `ev_e84bff74838a8aa2889c553a322e35ee438f8af31b1096ba86366558aa2fe51d`.

Les points CIL établissent des branches et contrats, sans prouver leur résultat dans Excel. Les oracles privés chargent ces deux identités contrôlées dans une instance Excel 16.0 x64 masquée appartenant au témoin. Ils initialisent le contexte Excel et n’altèrent aucun champ de licence. Les workers choisis évitent les wrappers qui écrivent le presse-papiers global.

## Somme copiée

Explorer `Actions.CopySum` RVA 70692 et son worker 337652 : sélection courante uniquement, sans multiplication par le nombre de feuilles groupées. Les rectangles sont réduits à `Worksheet.UsedRange`, avec repli sur le coin initial si l’intersection est vide. Les doublons de rectangles exacts sont éliminés ; les chevauchements partiels ne sont pas convertis en union sans doublons.

Le worker capture hauteurs et largeurs, puis accepte les cellules dont ces deux dimensions sont strictement positives. Lignes masquées manuellement, lignes filtrées et colonnes masquées sont donc écartées. Il lit `Range.Value`, qui conserve les types Date et Currency. Les Double sont ajoutés directement, les Decimal sont convertis en Double et les textes sont acceptés par `Double.TryParse(NumberStyles.Any, UserLocale.Culture)`. Dates typées, booléens, erreurs et textes non numériques contribuent zéro. Un échec du getter typé peut déclencher le repli `Value2` du natif ; cette frontière ne permet pas de garantir l’exclusion des dates dans tous les cas d’erreur.

`UserLocale.Culture` vient de `CultureInfo.CurrentCulture`, séparément de la langue UI Office et des séparateurs personnalisés Excel. Addition Double sans arrondi explicite, puis `Double.ToString(culture)`. La précision Currency peut déjà avoir été réduite par le getter COM. Les blocs de lecture ont au plus 1 000 lignes et 1 000 colonnes ; blocs de lignes, blocs de colonnes, puis lignes et colonnes internes déterminent l’ordre d’addition.

Le lookup de visibilité Common `BlockMap<double>` utilise des bandes d’une ligne, triées par origine, et un cache de prédécesseur. Il ne recherche pas une autre bande chevauchante si la bande retenue ne contient pas le point. Le relevé réel confirme le cas suivant : `B1=10`, `B2=20`, `B3=30`, `C2=100`, toutes visibles ; sélection `Union(B1:B3,B2:C2)` ; Arixcel donne **80**. Clair reproduit ce parcours, y compris l’exclusion de C2 et le second passage sur B2. Ce résultat ne représente ni la somme d’une union sans doublons, ni une addition complète des deux rectangles. Le volet indique que les chevauchements peuvent doubler ou exclure des cellules.

Dans Clair, le résultat reste lisible dans le volet et l’adaptateur l’écrit en `UnicodeText`, comme la valeur 1 du `TextDataFormat` observé. L’interruption ne publie aucun résultat. Le calcul ne crée aucun reçu d’annulation et ne change pas le contenu du classeur. Progression et arrêt restent dans le volet, conformément à la présentation demandée.

## Collages spéciaux

Explorer `Actions.PasteSpecial` RVA 77344 : modes 0 Formules, 1 Valeurs, 2 Formats, 3 Validation et 4 Formats numériques. L’entrée utilisateur exige le mode Excel **Copier**, une source identifiable et une sélection de cellules. Les modes 0 à 3 appellent `Range.PasteSpecial` avec `xlPasteFormulas`, `xlPasteValues`, `xlPasteFormats` ou `xlPasteValidation`, opération `xlNone`, sans imposer SkipBlanks ou Transpose.

Le snapshot d’annulation suit les familles natives : contenu A1 invariant, format numérique local et police pour Formules/Valeurs ; sept familles de formats pour Formats ; validation seule pour Validation ; format numérique seul pour Formats numériques. Les lignes ou colonnes entières ajoutent dimensions et niveau de plan. Les zones sont projetées sur les feuilles groupées. L’extension minimale à la taille de la source varie selon le mode et la première zone sélectionnée, comme le contrat relevé.

Formats numériques utilise un moteur de répétition de propriétés, avec découpe aux bornes de la cible. Les autres modes restent confiés au collage Excel natif. Les reçus partagent la pile chronologique U01/U02, vérifient les propriétés suivies avant annulation et gardent leur reçu en cas de conflit. Après annulation, la sélection et la source copiée sont rétablies. Les échecs d’écriture tentent une restauration explicite ; un échec de restauration est rapporté.

Les snapshots de contenu réutilisent le moteur du comparateur : plafond de 250 000 cellules, matrices et fusions entières, restrictions sur les origines de déversement. Les dégradés et la traduction des références de validation après déplacement structurel restent ouverts. Ces restrictions ne prouvent pas une parité complète avec le produit de référence.

## Adaptateur OLE

Common `Clipboard.GetRange` RVA 14332/14424 utilise `Link Source`, un `IDataObject`, un `IStream` et des monikers. L’adaptateur Clair décode le composite et une adresse absolue R1C1 locale. Il cherche la source parmi les classeurs déjà ouverts dans la même instance Excel. La liaison native susceptible d’ouvrir un fichier ou de retrouver une autre instance n’est pas reproduite dans cette passe.

Les essais synthétiques construisent de vrais monikers OLE et un vrai stream en mémoire, puis passent ce stream à l’adaptateur par un `IDataObject` local. Ils ne remplacent pas le presse-papiers Windows. Un composite de deux monikers est accepté ; un ou trois sont refusés. Le stream est repositionné avant décodage. Les références COM et le STGMEDIUM sont libérés.

L’interface injectable sert aussi aux essais du volet et des reçus. **Le véritable presse-papiers d’une session utilisateur et les appels natifs des quatre premiers modes ne sont pas exécutés par ces tests.** Leur présence dans le code et le dispatch vérifié ne constituent pas une validation de bout en bout de ce chemin.

## Duplication : fondation encore non raccordée

Explorer `PasteDuplicate` RVA 80724 distingue `PreserveInnerRefs=0` et `ExactCopy=1`. Le moteur pur Clair accepte une formule A1 et un contexte explicite : rectangle source, destination, classeur, feuille et éventuel tableau. Les références internes suivent le bloc, même absolues ou mixtes ; celles qui traversent partiellement la frontière source sont refusées. Les références externes au bloc restent attachées aux cellules d’origine. Les littéraux sont préservés.

La branche native à source d’une cellule conserve le texte A1 dans les deux modes. Le worker relevé traite aussi les constantes non textuelles d’une manière encore non validée dans Excel ; le helper Clair ne prend que des formules, pas ces valeurs. Références 3D, bornes calculées ou nommées, résolution générale des noms et qualification explicite de la feuille source demandent encore des oracles. Une commande générique de duplication avec formats, copie et annulation reste à raccorder et tester.

Les preuves privées détaillées sont conservées dans les ledgers `u03-clipboard-contracts.md`, `u03-duplicate-contracts.md` et `u03-paste-special-ledger.md` du dossier de travail. Les résultats effectivement exécutés figurent dans [VALIDATION.md](VALIDATION.md).

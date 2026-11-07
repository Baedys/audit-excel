# Revue et validation

## Copie et presse-papiers — développement U03

Le volet Utilitaires ajoute Copier la somme, cinq collages spéciaux et la duplication. Les [preuves U03](CLIPBOARD-EVIDENCE.md) séparent contrats statiques, commandes raccordées, workers originaux exécutés et frontières encore ouvertes. Les validations propres à la duplication sont détaillées dans la section suivante.

La suite hors Excel passe 86 groupes. Elle comprend 69 contrôles de l’adaptateur de presse-papiers, 72 du helper de formules, 185 de géométrie et 49 de notation. L’adaptateur décode un vrai composite OLE depuis un stream et un IDataObject en mémoire : aucun presse-papiers global n’est remplacé. La lecture commence à un offset volontairement non nul ; le repositionnement et le rejet des composites d’une ou trois parties sont vérifiés. Les classeurs absents ou ambigus sont refusés sans ouverture de fichier.

Le témoin du volet et de la somme passe 103 contrôles dans Excel 16.0 x64 privé et masqué, dont 51 comparaisons au worker original. Le corpus couvre valeurs typées, dates, Currency, textes numériques, formules, erreurs, filtres, lignes et colonnes masquées, fusions, feuilles groupées, zones disjointes et chevauchantes, limites de UsedRange, blocs de plus de 1 000 lignes ou colonnes, cultures fr-FR/en-US/de-DE contrôlées et séparateurs Excel personnalisés. Le volet utilise un writer injecté ; interruption et échec de ce writer conservent la pile d’annulation. Le collage numérique raccordé au volet et son reçu partagent cette pile. Ces essais de cultures ne représentent pas trois installations Office différentes.

Un autre oracle exécute le worker Arixcel de Formats numériques sur douze scénarios : répétition, blocs finaux partiels, extension minimale, cible d’une ligne ou colonne, source d’une cellule, groupe de feuilles, deux ordres de zones, source uniforme, chevauchement et source multi-zone. Il compare cinq propriétés de 144 cellules sur trois feuilles, avant et après collage, puis après annulation Clair : **77 760 comparaisons de propriétés**, zéro divergence, plus 122 contrôles de contexte et de copie injectée. Il n’exécute ni le wrapper d’identification du presse-papiers, ni la recopie globale de la source, ni l’annulation native Arixcel.

Un oracle d’axes distinct passe 43 contrôles. Les neuf lectures dynamique, late binding et PIA concordent. Sur la hauteur initiale 14,3, la propriété Arixcel mémorise bien 14,3, puis son setter donne 14,4 ; le reçu Clair donne aussi 14,4. La largeur initiale 10,07 reste 10,07. La quantification observée ne permet pas de promettre un retour au dixième affiché pour toutes les tailles. Le test de reçu général utilise séparément une hauteur contrôlée de 12 points et conserve une vérification exacte de restauration.

Le témoin des collages passe 99 contrôles dans le gate final : cinq modes, extension et répétition, source multi-zone conservée, familles de propriétés restaurées, validation mixte, dimensions et plans, rollback, conflit, échec de recopie et reprise. La revue en lecture seule a corrigé une source multi-zone tronquée : le reçu conserve maintenant l’union complète pour la recopie après annulation, avec un rectangle distinct pour la géométrie native. Le groupe est revérifié après restauration de la sélection, même s’il correspondait avant cette sélection. Les constantes de dispatch sont testées par callbacks injectés ; ces callbacks ne prouvent pas le fonctionnement réel des quatre premiers collages avec le presse-papiers Windows. Le plafond de contenu de 250 000 cellules, certaines matrices/fusions, les dégradés, les références de validation déplacées, les autres configurations Office et le coût des grandes sélections restent ouverts. U03–U06 et la parité complète ne sont pas déclarés achevés.

## Duplication de cellules — développement U03

La commande **Dupliquer** est raccordée au volet : PreserveInnerRefs=0, ExactCopy=1, option de formats, progression, arrêt et reçu d’annulation partagé. Les tests produit utilisent une instance Excel privée et masquée, avec callbacks injectés pour le collage initial, la reconstruction par liens et la recopie de source. Les écritures de contenus et les restaurations sont réelles ; ces callbacks évitent le presse-papiers global.

Le gate externe final passe **534 contrôles de duplication**, **99 de collages spéciaux** et **67 du panneau**, dans trois sessions Excel privées distinctes : PID82280, PID87132 et PID70600. Il charge Clair.exe SHA-256 `ECD85AF382DA0FC3BDAE28E90D4898896E15C98F4AF5CCAE9858F6C16D6716CF`. Le journal `work/clipboard-final-f6e8e04d5c094b448331473fa3288b27/managed.log` (SHA-256 `877a4628bd95f9cf8aeb1938ca0cbf7bc72419f4b279732dd6139d4b2468b402`) termine par `CLIPBOARD_REGRESSION_RESULT success`. Les trois sessions n’ont pas chargé l’oracle Single original : le compteur de comparaisons Single originales vaut explicitement **0**.

Les 534 contrôles couvrent les quatre combinaisons mode/formats, références et constantes, tuiles et zones disjointes, groupes avec source par feuille ou cache d’une autre feuille, lecture source après chevauchement et irrégulier filtré. Ils comprennent matrices CSE complètes, source CSE d’une cellule, membre d’une matrice source, répétition en plusieurs tuiles, déversement pris en charge, fusions et contenus restaurés, protection avec cible déverrouillée, puis notation des formules, valeurs typées et littéraux sous General/Texte. Setters et assertions CSE utilisent la culture en-US. Les échecs contrôlés comprennent mutation initiale partielle, sélection post-paste absente, annulation avant collage et entre tuiles, recopie impossible puis reprise, échec du collage de liens et changement de zones.

Les 67 contrôles du panneau couvrent les deux modes, la transmission de l’option de formats désactivée, la pile commune du volet, rollback, annulation, fermeture, sortie du volet pendant une opération et transfert différé de plusieurs reçus de duplication vers une vraie session d’annulation avec enregistrement injecté. Une demande de fermeture pendant la duplication annule celle-ci et ferme le volet après restauration ; l’historique précédent reste utilisable. Ces contrôles ne prouvent pas un rappel déclenché depuis le véritable menu Annuler pour cette nouvelle commande.

Les helpers courants passent **72 contrôles de réécriture, 185 de géométrie et 49 de notation**. Les deux contrôles de réécriture ajoutés distinguent les noms qualifiés `S!A:S!A` des références de colonnes. La notation couvre ancrages, axes, quotes, noms/tableaux, classeur cible, Unicode, noms quotés du getter et une expression pure de 5 000 références. Le digest fonctionnel du classificateur complet des 65 536 unités UTF-16 vaut `e63632e99f5b1c85ebeadcb9a86fbfe751bd6b7c187a8655a8efea0e31e9bf31`. Le test le recalcule sans charger de ressource Arixcel. Il ne prouve pas qu’Excel accepte comme saisie une formule de 5 000 références ni toute graphie quotée.

L’oracle original de `FastRangeStore.MergeRanges`, des Values et de ToOrderedList passe séparément **96 cas, 770 assertions, zéro écart**, sur rectangles virtuels sans cellules ni presse-papiers. Le corpus contient 80 graines et au plus dix rectangles par cas. Pack est comparé à la transcription CIL du placement, avec cumul sans réinitialisation de coordonnée ; ce résultat ne constitue pas une exécution du collage Excel irrégulier.

La capture des constantes utilise maintenant le getter Formula2 brut, pas Value2 ni une conversion G17. Toutes les données d’un bloc PreserveInnerRefs sont appliquées par writer bulk R1C1 ; les sources d’une cellule et ExactCopy utilisent bulk A1. Un convertisseur pur traduit les vraies formules réécrites vers R1C1 au point cible et reproduit les règles de quotes natives, Unicode et de noms de tableaux du classeur cible. Les noms quotés issus des getters sont reconnus sans être traités comme des cellules : une source `='A1'` peut être obtenue par seed R1C1 `=A1`. Les setters A1 directs de certaines graphies quotées ont été refusés ; l’acceptation du getter par le parseur ne garantit pas leur acceptation en saisie A1.

Le gate natif des matrices utilise la **même DLL privée `4AFFE93D47CF6C6091633C22C5EDF6C45156BC32ECBEE3E92623CD13CD1DC3D0`**, antérieure au focus qualifié ci-dessous, dans deux nouveaux Excel masqués, PID75416 et PID81356, quittés naturellement. La véritable API Clair est comparée aux applications des props/stores natifs aux mêmes coordonnées cibles :

- **20 cas formules, 384 assertions, zéro écart** : références internes/externes, source d’une cellule ou bloc, deux modes, General/Texte, nom A1 quoté issu de R1C1 et destination dans un autre classeur. Sous Texte, PreserveInnerRefs produit notamment `=RC[1]`, `=S!R[-3]C[-1]` ou `=[Feuil1]S!R[-3]C[-1]`, comme le natif.
- **48 cas constantes/littéraux, 912 assertions, zéro écart** : 24 littéraux `=A1`, `=R1C4`, `=R[-1]C` avec/sans apostrophe, deux modes et General/Texte ; 24 cas nombres/booléens/erreurs, source d’une cellule ou bloc, deux modes et General/Texte. Les écarts antérieurs sont corrigés : le natif et Clair donnent des valeurs typées sous General et des chaînes `12.5`, `TRUE`, `#DIV/0!` sous Texte. Il ne faut pas transformer le texte brut en Value2 avant le writer.

Journaux privés immuables : `work/u03-duplicate-writer-final-2c0e9597890e4b9fbcb636315d845470/formula-format.log` (SHA-256 `79eb027638adb129b19c77bd998193a7716be9941dd0160ac0551cb47fd05d84`) et `coercion-matrix.log` du même dossier (SHA-256 `e2384e43b3df05dba4d8774e3c035a82b32014148011e8a66828fd40c3075e8b`). Collage initial typé, lecture de sélection et recopie sont injectés ; aucun presse-papiers global. Les trois phases publiques FormulaTargetFormats, ConstantTargetFormats et LiteralNotation, auparavant passées séparément avec **280 contrôles**, font désormais partie du gate complet à 534. Les matrices natives de 20 et 48 cas n’ont pas été rejouées sur la DLL suivante.

Un focus réel de la véritable API Clair sur Clair.AddIn.dll SHA-256 `CA82698ED543E7D4AA29C6A12E6419D49D4FAB2F2BD6174768CF675F2A6FA69E` passe **2 cas, 71 assertions, zéro écart**. Cette DLL et l’exécutable ECD85 du gate final sont deux cibles compilées séparément par build.ps1 depuis le même état des sources, sans identité binaire commune. La source `=SUM(S!A:S!A)` conserve deux noms qualifiés, et non une plage de colonnes. Le natif et Clair concordent sous General et Texte pour Formula2, Value2, PrefixCharacter, NumberFormatLocal et HasFormula. L’annulation Clair restaure exactement ces cinq champs sur les deux cellules cibles, préserve la source et rétablit ScreenUpdating/EnableEvents. L’Excel privé PID88872 est quitté naturellement ; aucun presse-papiers ni Undo.Run natif. Journal : `work/u03-qualified-names-ac2bc1d8841e41e39023760f1385cf7a/managed.log`, SHA-256 `2b74ca06e14538c90402fb093db4ef1e7b2a1bdb0781520344455a2c57e1ebfa`.

Un oracle privé antérieur de formules recense **75 cas, 127 assertions, zéro divergence parmi les comparaisons réalisées, zéro défaut de préparation**. Le corpus couvre références mixtes et qualifiées, autres classeurs, 3D, bornes calculées, noms, références de tableaux implicites et explicites, contenus typés, sources d’une cellule, CSE, déversement et textes sous General ou Texte. Son compteur brut conserve deux `applicationGaps` : les cas ExactCopy « implicit table outside » (`=[@Amount]+$A$2`) et « table normalizes to implicit » (source canonicalisée `=SUM([Amount])+A2`), pour lesquels le bulk Apply natif lève `OutOfMemoryException`. Le journal privé `u03-duplicate-external-db83dcce57d14f359837f305c8a2fe71/managed.log` porte le SHA-256 `79d14cf0c72c5e8e930c1dfeb5f48d81bb9e347dbfb793c69ddd3d1d08341746`. Il ne représente pas 75 cas de bout en bout de l’API Clair et n’a pas été rejoué sur la DLL finale du writer.

Un témoin produit séparé a fermé cette incertitude pour ces deux entrées sur une version antérieure : véritable `RunPasteDuplicateUtility`, ExactCopy sans formats, nouvelle session Excel privée, **2 cas, 46 contrôles, zéro écart**. Le bulk write Clair rejette les deux opérations avec la même `OutOfMemoryException` (`0x8007000E`). Le marqueur de rollback est vrai ; les six valeurs sentinelles 91 sont restaurées, les neuf cellules sources restent identiques sur quatre propriétés, ScreenUpdating et EnableEvents sont rétablis, et la sélection post-paste E5:G6 est conservée. Aucun callback de recopie n’est demandé et aucun presse-papiers global n’est utilisé. Le journal est `work/u03-duplicate-product-gaps-2ea3bb2db7134067bc78823549981d47/managed.log`. Il s’agit de **rejets partagés suivis d’un rollback validé**, sans preuve de duplication réussie sur ces deux formules et sans conclusion générale sur les tableaux. Ces deux entrées n’ont pas été rejouées sur la DLL finale du writer.

Les résultats ciblés ne valident pas encore le vrai presse-papiers Windows, les appels forward natifs Paste/Copy sans destination de bout en bout, les autres instances ou sources fermées, le plafond de 250 000 cellules, tous les dégradés/règles CF et validations, ni U04–U06. Le code et les contrats statiques sont distingués des chemins réellement exécutés ; U03 et la parité complète restent partiels.

## Utilitaires de formats — développement U02

Le relevé des signatures et du CIL d’Explorer et Common établit huit commandes de formats : `CellColor` 70632, `TextColor` 79408, `Style` 78952, `NumberFormat` 79624, `DecimalPlaces` 77016, `Border` 81856, `TextUnderline` 72568 et `TextControl` 72836. Le helper `ConvertNumberFormat` 79836 convertit un code invariant en code d’affichage ou l’inverse. Les commandes projettent les zones de sélection sur les feuilles de calcul sélectionnées, sans les réduire à UsedRange.

Le format numérique est reconstruit par un lexer indépendant, les propriétés `Application.International` et les huit mots réservés de couleur dans les 29 langues du relevé. La date standard invariante `m/d/yyyy` adopte l’ordre, les zéros initiaux et les séparateurs du standard local. Les autres formats de date gardent leur ordre. Les devises et littéraux suivent des branches distinctes. Le traducteur ne modifie pas les conditions numériques ou `ColorN`. L’adaptateur lit la langue UI Office, tandis que les règles de casse utilisent la culture du processus comme le traducteur observé.

Les décimales utilisent uniquement le format de la première cellule de la sélection pour produire un code commun à toutes les cibles. Le helper Common 34176/35984 traite les trois premières sections numériques, préserve les littéraux, les échappements, les crochets et l’exposant. La diminution d’une cellule General conserve General ; l’augmentation produit `0.0`, avec le point invariant observé.

`CellColor` utilise les couches du journal partagé : la fermeture du reçu conserve sa couleur comme nouvelle base et retire l’historique situé dessous, conformément au nettoyage du propriétaire natif. Un fond vide signifie absence de fond ; un texte vide conserve la couleur de police. Le noir correspond à la couleur automatique pour ces deux propriétés. Sans filtre actif, les lignes masquées participent. Dès qu’un filtre de feuille ou de tableau est actif, les fragments dont la hauteur de ligne vaut zéro sont exclus, y compris les lignes masquées manuellement hors du tableau. Les colonnes masquées restent incluses. La couleur de texte seule, distincte du journal de fond, conserve la police complète pour son annulation.

Les seize variantes de bordures suivent les masques d’arêtes et les valeurs natives. Les variantes nommées Dotted dans Arixcel emploient un trait continu Hairline. Les modes Thick concernés emploient Medium. Une couleur vide conserve la couleur de chaque bord ; le noir est un RVB nul. Les cinq modes de soulignement correspondent à aucun, simple, simple comptable, double et double comptable.

La disposition du texte comprend renvoi à la ligne, retrait du renvoi, fusion, fusion centrée, dissociation, alignements gauche/droite/centré/centré sur sélection/général et variation de retrait. Le retrait est contrôlé sur tous les fragments avant écriture, entre 0 et 250. Tous les fragments gardent l’alignement à droite uniquement s’ils le possédaient tous ; sinon ils passent tous à gauche. Fusionner capture contenus et validations. Fusionner et centrer suit le reçu natif d’alignement seul : son annulation ne recrée pas les contenus supprimés hors du coin haut gauche. Le volet indique cette différence avant exécution.

Un style absent peut être créé explicitement depuis la cellule active, avec un choix dans le volet à la place du dialogue natif. La nouvelle définition de style reste dans le classeur après annulation. Les états de police enrichie, couleurs de thème et nuances sont conservés pour l’annulation prise en charge. Les lectures qui doivent neutraliser momentanément une nuance RVB sont isolées de la feuille groupée. Les fusions restaurées utilisent un déplacement relatif à la plage COM retenue, afin de suivre une insertion avant cette plage.

Le build final passe 82 groupes hors Excel. Le témoin sur le thread UI d’une instance Excel 16.0 x64 privée et masquée passe 843 contrôles de formats : groupes et multizones, polices enrichies, cinq soulignements, seize bordures avec diagonales individuelles, douze dispositions, contenus et validations après fusion, création et restauration de styles, formats locaux, filtres et masquages, rollback sur feuille protégée, conflit et reprise, fusion déplacée après insertion et rappel géré d’annulation mixte. Il conserve aussi les 46 contrôles U01 et les 193 contrôles de fenêtres, gel et scission compris.

Un oracle privé charge les assemblages d’origine dont les SHA-256 sont contrôlés, puis initialise leur contexte avec l’application Excel appartenant au témoin. Il ne modifie aucun champ de licence. Il passe 1 810 comparaisons de décimales, 174 de conversion pure et 87 de l’adaptateur Excel. S’y ajoutent 4 176 comparaisons contrôlées de mots de couleur : 29 langues, huit couleurs, trois casses et six enveloppes, avec des locales synthétiques isolées. Ces dernières ne constituent pas des essais de 29 configurations Office. Total : 6 247 comparaisons, zéro divergence. Le témoin du volet vérifie séparément 24 conversions contre le helper Arixcel, puis la date effectivement appliquée après les autres commandes.

Le même hôte, avec la DLL exacte de production et Arixcel connecté, rejoue les 24 contrôles natifs antérieurs de comparaison, formats et formules. Un lot supplémentaire de trois reçus applique texte rouge, double soulignement comptable et zoom 125 depuis 95. Le pilote place un effacement natif au-dessus, puis exécute deux vraies commandes `CommandBars.ExecuteMso("Undo")`. La première restaure l’effacement ; la seconde appelle le XLL et restaure ensemble couleur, soulignement et zoom. Ce passage compte 14 contrôles du pilote et 21 du capteur dans Excel. La frappe physique n’est pas simulée.

Les essais ont corrigé un problème de restauration de fusion : le moteur partagé emploie désormais un chemin explicite de capture et d’écriture `NumberFormatLocal` pour ce lot. Le comparateur conserve le chemin invariant, choisi par défaut. Le code local `0.0`, avec un séparateur décimal virgule, est traité comme le natif et devient `0,0.0` à l’augmentation ; il ne doit pas être confondu avec le code local `0,0`. Les bordures sont capturées depuis leur collection uniforme ou partitionnées ; une cellule conserve six arêtes, tandis qu’une capture agrégée couvre les contours et les deux directions internes, sans attribuer une diagonale absente à partir d’un getter agrégé trompeur.

Les limites héritées restent explicites : dégradés dans les captures complètes de style ou fusion, plafond de 250 000 cellules du moteur de contenu réutilisé pour Fusionner, traduction des formules de validation après déplacement structurel ou changement de ReferenceStyle, redimensionnement interne d’une fusion partiellement sélectionnée et coût des très grandes sélections. Les variantes et cycles U06 restent distincts des réglages de présentation du volet. Office 32 bits, les autres configurations et la parité visuelle complète restent à vérifier.

## Utilitaires de mise en page

Le relevé statique des actions Arixcel porte sur `ColumnWidths` 69868, `NewWindow` 70876, `ToggleGridlines` 73712, `RemovePageBreaks` 74216, `Zoom` 74384 et `AutoFit` 80380. Les scopes 0 et 1 désignent les feuilles sélectionnées et toutes les feuilles. L’axe 0 est la colonne et l’axe 1 la ligne. Les six actions n’ont aucun paramètre CLR optionnel. Le gestionnaire commence avec une liste active vide en l’absence de configuration. Les préréglages du volet Clair sont des choix de présentation.

`ColumnWidths` lit la feuille de référence du comparateur, exclut celle-ci des cibles et capture toutes les colonnes. Common 98272 tente la dérivation avant le découpage. 88844 accepte une largeur de type Double et 89088 l’applique aux colonnes entières. Un témoin dédié a exécuté ces deux méthodes Common dans Excel masqué. Sur une référence où A vaut 7 et B vaut 18, la bande complète donne un Double égal à 7. La dérivation conserve 7, puis l’application ramène une cible B de 31 à 7. La référence B reste à 18. Ce résultat établit l’uniformisation native sur ce poste.

L’utilisateur a ensuite explicitement choisi de préserver chaque largeur. Clair lit désormais les 16 384 colonnes individuellement, puis applique leurs suites de largeurs aux autres feuilles sélectionnées. Son annulation utilise aussi les valeurs individuelles d’origine des cibles, en distinguant cette commande du comportement natif antérieur. La référence n’est jamais réaffichée temporairement : une colonne masquée fournit la valeur réellement exposée par `ColumnWidth`. Le masquage n’est pas capturé ni copié comme une propriété séparée.

Le témoin Excel x64 masqué passe 57 contrôles de mise en page : référence A=7, B=18, dernière colonne XFD=21, deux cibles groupées, conservation de la référence masquée et restauration des largeurs hétérogènes de chaque cible. La colonne source masquée expose ici zéro ; écrire cette largeur peut masquer la colonne cible. La largeur latente d’une cible déjà masquée après son réaffichage n’est pas garantie par ce reçu. Le balayage individuel complet est synchrone et son coût augmente avec le nombre de feuilles, surtout via COM entre processus.

`AutoFit` réplique les zones de sélection sur les feuilles groupées et utilise l’intersection avec UsedRange, avec repli sur le coin haut gauche si elle est vide. L’ajustement vise l’axe entier. Son reçu restaure la sélection d’origine. Le zoom traite les feuilles visibles, y compris les graphiques, et conserve leur ancienne valeur individuelle. Le quadrillage s’applique directement aux WorksheetViews de la première fenêtre du classeur. Son reçu inverse uniformément la nouvelle valeur. Le masquage des pointillés modifie `DisplayPageBreaks` et garde les sauts d’impression.

`NewWindow` capture les vues des feuilles visibles, exécute `WindowNew` et réapplique ces vues à la nouvelle fenêtre. Les paramètres comprennent affichage, zoom, couleur du quadrillage, gel ou scission, défilement des volets et sélection. Le groupe et la feuille active sont restaurés en dernier. Les feuilles masquées et les classeurs non ciblés restent hors de cette capture. Si une réapplication échoue après création, la nouvelle fenêtre reste ouverte et le volet le signale. La commande ne tente pas de fermer le classeur par une fermeture de fenêtre en compensation.

Les contrôles de configuration couvrent l’aller-retour XML, les paramètres indépendants et le rejet d’un zoom localisé ou d’une version inconnue, sans remplacer le fichier valide. La copie de largeurs, l’ajustement, le zoom et le quadrillage partagent la pile d’utilitaires du volet et le rappel d’annulation Excel. Les modifications concurrentes sont contrôlées avant restauration. Les branches d’échec restaurent les dimensions observables déjà modifiées. La lecture précise de toutes les dimensions est plus coûteuse dans le pilote COM externe que dans l’hôte Excel et doit encore être mesurée dans le complément.

Le build final passe 79 groupes hors Excel, dont les cinq contrôles de préférences. Le témoin chargé avec la DLL exacte du complément sur le thread UI d’Excel x64 masqué passe 46 contrôles d’utilitaires et 193 contrôles de fenêtres. Il vérifie les 16 384 colonnes, les feuilles groupées, les deux axes d’ajustement, les zooms distincts, le quadrillage inverse, la conservation des sauts d’impression, le rappel d’annulation et sa reprise après conflit. Les trois vues de feuilles, la sélection multizone, la cellule active, le gel, la scission, les défilements, la source et un autre classeur sont vérifiés. Le contrôle de protection des fenêtres est signalé indisponible lorsque cette version d’Excel ignore ce réglage. Les 104 contrôles du tracé passent aussi sur le build précédent l’initialisation du cadre masqué.

Dans Excel masqué, une fenêtre créée par `WindowNew` peut exposer quatre volets sans avoir calculé leurs positions. La commande initialise alors son cadre avec l’état et les dimensions de la fenêtre source avant d’appliquer les vues. Ce chemin de compatibilité n’active pas `Visible`. Le même témoin conserve Arixcel connecté et passe ensuite 24 contrôles d’annulation native mixte des formules et formats.

## Coordonnées des résultats après une modification Excel

Dans Arixcel, 272836 matérialise les objets `Range` des tâches, blocs et écarts. 273040 relit leurs rectangles au retour. Il adopte les déplacements et redimensionnements et écarte le côté inaccessible. Les données du relevé restent en mémoire. 372428 exige des tailles égales pour la copie et 231944 utilise le côté actif pour F2. La validité du périmètre de plage contrôlée par 264688 et Common 54176 reste distincte.

Clair conserve les coordonnées capturées et les plages physiques sur deux couches. Le retour au volet rafraîchit les adresses de navigation et de rapport. Le témoin Excel masqué couvre insertion avant et dans un bloc, renommage, suppression d’une cellule, remplacement par une feuille homonyme, tracé à l’adresse déplacée, F2 et annulation sur un bloc redimensionné. La suite de comparaison vérifie aussi la copie et son annulation sur une cellule déplacée. Lorsqu’un ancien périmètre de plage ne permet plus de relancer le relevé après une écriture, le volet conserve la capture précédente jusqu’à une nouvelle comparaison. Le tracé lit les cellules courantes. Si sa formule reprend le détail comparé, son édition reste rattachée au bloc physique après déplacement ou renommage.

## Progression et interruptions du comparateur

Arixcel déclenche son moniteur après 2 000 ms via Explorer 252212 et Common 73848. 245603 ignore l’unité active. 273620 remplace cette unité pour chaque tâche et chaque bloc. 245464 annule pendant la préparation ou après une demande d’arrêt. Durant l’alignement et le scan, le premier arrêt conserve les résultats produits. Common 75764 vide les rectangles ouverts avant de rendre un résultat partiel. 266068 exporte les tâches interrompues et les bornes inférieures.

Clair place ces commandes dans le volet. Les points de contrôle traitent l’annulation entre les lectures Excel et aux frontières des lignes et blocs. Un appel COM individuel peut retarder la réponse. Annuler laisse le résultat précédent et une comparaison différée peut être reprise. Les tâches et blocs interrompus restent distincts et l’export indique une borne inférieure. La copie ou l’édition d’un bloc incomplet entier est désactivée. Ses sous-écarts déjà lus restent utilisables.

Le témoin Excel masqué vérifie 48 résultats, dont l’omission d’une tâche pendant sa lecture, d’un bloc encore vide, l’arrêt durant l’alignement, les lignes rendues après un arrêt et le retour au résultat précédent après annulation. Il active les boutons du volet de manière programmée et vérifie le rapport Excel partiel.

## Retour à la grille et actualisation du comparateur

Dans Arixcel, `ModelessForm.GoBackToExcel` 58303 rend le focus à Excel. `History.GoBack` 64192 réactive une vue disponible. Le retour redessine la comparaison et synchronise les coordonnées de ses plages par 233522 et 273040. Alt+R suit un autre chemin. `CompareForm.Recompare` 223744 vérifie le périmètre, reconstruit le modèle avec ses options, conserve les côtés et tente de retrouver les nœuds dans la même phase. Un changement de phase vide l’annulation locale. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair inscrit le comparateur dans l’historique partagé, y compris quand aucune formule n’a été explorée dans le volet. Le retour conserve le modèle et la sélection de la grille. Alt+R repart des classeurs et plages capturés avec les mêmes options. Le volet retrouve les lignes, coches et groupes ouverts encore présents et conserve l’annulation dans la même phase. Les nœuds gris sont retirés lors du nouveau relevé. Un changement de phase réinitialise la sélection et l’annulation locale. L’essai de frappe physique reste ouvert.

Le témoin Excel masqué vérifie 41 résultats. Il passe par les gestionnaires clavier du volet et par l’historique du complément. Il couvre deux fenêtres Excel, la sélection de grille, le retour au contrôle de détail, la recapture des formules et la conservation des options, de la sélection et de l’annulation. Il couvre aussi un brouillon actif et une référence remplacée.

## Rapport Excel du comparateur

`CompareForm.Export` 231112 appelle `Report.Generate` 266068 après la comparaison. Le rapport Arixcel ouvre un classeur à une feuille. Les en-têtes sont à la ligne 5. 375368 choisit cinq colonnes pour un résultat unique de feuille ou de plage et sept lorsque la tâche ou la feuille est indiquée. Les colonnes de contenu gardent leurs côtés physiques après échange visuel. Les chaînes de formule sont protégées par 375551 et colorées par 375592. 250940 produit le total et le préfixe indiquant un résultat partiel. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair exporte le résultat capturé dans la feuille `Comparaison`. Le filtre et les lignes figées facilitent la lecture du tableau. Les formules et les textes d’entrée restent littéraux. Les écarts résolus peuvent rester dans le détail d’une tâche ouverte. Un onglet hors comparaison, un onglet supplémentaire et une tâche non terminée portent des états distincts. La capture conserve les feuilles lisibles lorsqu’une autre dépasse la limite ou ne peut pas être lue. Les commandes interactives Skip et Stop du scan Arixcel restent à porter.

Le témoin Excel masqué vérifie 44 résultats. Il couvre les rapports à cinq et sept colonnes, les côtés inversés dans le volet, la coloration de caractères, une sous-différence résolue au milieu d’une tâche ouverte, les onglets supplémentaires, un graphique ignoré et une capture partielle à cause de la limite. Il vérifie aussi les nombres, les booléens, les erreurs, les textes qui ressemblent à une formule et l’export sans écart.

Dans le témoin x64 embarqué, Excel masqué refuse le gel même après retour à sa boucle d’interface. Le rapport filtrable et ses formules littérales restent disponibles. Le volet signale le gel refusé. Le témoin COM externe vérifie séparément le gel de cinq lignes quand Excel l’applique.

## Tracé depuis la comparaison

`CompareForm.OpenFormulaExplorerForm` 230148 choisit les plages du côté actif. Pour une formule, il prend le coin haut gauche du bloc. Common 56608 déduplique les rectangles. L’explorateur 193548 reprend le curseur sur la même formule et 195592 délègue l’écriture au comparateur. Le comparateur 231892 contrôle le rattachement au modèle et 231944 écrit sur le bloc d’origine. Le retour groupe les éditions du tracé dans la pile parente. Ce relevé vient de l’inspection statique d’Arixcel. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair garde le tracé et la comparaison dans le même volet. Le témoin couvre le côté actif après échange visuel, la sélection multiple, les antécédents et dépendants hors des plages comparées, les différences dans le texte, le transfert du curseur et le retour au résultat. Il vérifie l’édition du bloc, l’édition d’un antécédent hors bloc, l’annulation dans le tracé et le lot d’annulation dans le comparateur. Si le modèle change après l’ouverture d’un brouillon rattaché, Clair demande de rouvrir l’édition.

Les 104 contrôles du tracé couvrent aussi le rattachement à un bloc déplacé avant son ouverture, son redimensionnement et le renommage de sa feuille pendant le tracé. Les éditions successives restent possibles après renommage de la feuille opposée. L’édition et l’annulation visent le rectangle physique entier, y compris une cellule nouvellement insérée. Un déplacement ultérieur de la racine demande une actualisation du tracé. Une plage supprimée reste inaccessible. Un échec d’actualisation après écriture restaure le bloc et conserve le brouillon sans ajouter de reçu d’annulation. Ces cas reprennent les conditions de rattachement de 230148 et 231892, la cible de 231944 et la restauration après échec de 195592.

## Édition depuis la comparaison

Dans Arixcel, `KeyDownHandler` 229608 et `StartEditing` 231400 ouvrent l’édition du côté actif en mode Formules. `TreeItem.FormulaEditable` 272168 autorise un sous-écart vide et exclut les lignes de tâche. `FormulaChangeAction` 231944 écrit une formule A1 dans le coin du bloc, la relit en R1C1 puis applique cette formule relative à toute la plage. Une saisie R1C1 ou littérale s’applique directement. Common 40180 et 40888 traitent la matrice historique. 232500 valide et 232832 abandonne le brouillon. L’annulation restaure la plage et réactualise les sous-écarts. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Le volet Clair fige la plage physique et le style de saisie à l’ouverture. Avant l’écriture, il vérifie le périmètre de comparaison, les identités et le contenu de la plage. Une erreur laisse le brouillon disponible et s’affiche dans le volet. Le reçu préserve le contenu et les formats numériques et partage la pile de copie et d’alignement. Un écart égalisé est grisé. Le témoin Excel masqué couvre les deux côtés, les blocs A1 et R1C1, la saisie directe, les cellules vides, les données hétérogènes, les refus sur une plage modifiée ou déplacée et l’annulation d’une copie suivie d’une édition.

Les matrices CSE doivent être sélectionnées en entier et restent limitées à 255 caractères. Les fusions et feuilles protégées sont écartées avant l’écriture.

## Segments du détail de comparaison

`CompareForm.UpdateSides` 224148 appelle `StringMatch.GetUnmatchedParts` 79260 avec une distance maximale de 2 000. `FormulaString` 244396 et 244640 considère deux caractères de mot égaux à la même position d’une suite complète et identique de lettres, chiffres ou signes dollar. La ponctuation se compare caractère par caractère. La casse compte. Les indices sont UTF-16. 273948 utilise le style d’affichage, séparé du style de comparaison. 224148 colore en rouge les segments et décale leur début lorsqu’il ajoute des guillemets autour d’une entrée littérale égale au libellé des données hétérogènes. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Le témoin isolé appelle les DLL livrées, sans Excel, sur 3 029 paires. Les plages des deux côtés concordent, y compris quand plusieurs alignements sont possibles et aux distances 2 000 et 2 001. Les contrôles WinForms couvrent les retours ligne, le texte inchangé, les libellés, le marqueur des données hétérogènes, l’inversion des côtés et les lignes résolues. Excel masqué vérifie le passage A1/R1C1 avec une comparaison R1C1 déjà ouverte.

Le volet utilise le texte de la capture. Pour une formule locale, il vérifie la formule actuelle avant de demander le texte à Excel. Si la présentation locale échoue, les deux formules utilisent leur texte invariant. Le rapport garde le texte capturé pour le mode de comparaison. Les valeurs typées comme les dates, devises, booléens et erreurs continuent de suivre la capture `Value2` de Clair.

## Affichage des écarts résolus

Le CIL de `DiffItem.UpdateFormula` 271184 conserve le nœud, relit le côté écrit et reporte la variation du compteur aux ancêtres. `TreeView_DrawControl` 234668 grise les nœuds à zéro. 373684 écarte ces nœuds de la copie, tandis que 225176 autorise leur sélection et leur surbrillance. 374268 fixe à zéro les sous-écarts d’un bloc copié. 277840 force à un les sous-écarts après annulation. `SheetItem.UpdateLabels` 215484 donne zéro aux onglets désormais appariés. `Recompare` 223744 construit un nouvel arbre. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair vérifie l’état résolu avec les cellules de la paire physique et les règles actives de comparaison. Un bloc mixte conserve ses sous-écarts et sa plage de copie entière, y compris ses cellules intermédiaires. Après annulation, un sous-écart qui était déjà égal reste résolu. Un onglet est marqué copié lorsque les deux identités physiques attendues sont présentes. Le rapport sépare les écarts ouverts et résolus. Le témoin Excel masqué couvre les couleurs du volet, les boutons, la navigation, la copie de bloc, le renouvellement de la comparaison, un conflit d’annulation et les onglets.

## Copie des formats directs — 7 novembre 2026

Dans Explorer, 247276 active la copie des formats par défaut. 257996 désactive l’option en mode valeurs. 253168 et 374616 appliquent les formats après le contenu et le format numérique. Common 86492 et 87044 traitent fusions et alignement. 89588 et 89724 traitent la police. 81172 découpe les caractères dont le format est hétérogène. 87816 traite les bordures, 92100 et 92800 la couleur et le motif du fond, et 88284 la protection. 21760 et 22544 transmettent le RVB effectif ou les index Automatique et Aucun. [Polices Excel](https://learn.microsoft.com/en-us/office/vba/api/Excel.Font%28object%29) · [Bordures Excel](https://learn.microsoft.com/en-us/office/vba/api/excel.borders)

Clair écrit les propriétés directes de la plage. Les règles conditionnelles et la validation restent dans la destination. L’annulation rétablit la liaison au thème, la nuance et les formats des caractères originaux. Elle protège les modifications ultérieures. Une fusion doit être entièrement comprise dans la plage. Les dégradés donnent lieu à une copie du contenu avec un message dans le volet. Si Excel refuse un format après écriture, Clair rétablit l’état d’origine avant de réessayer le contenu seul.

## Copier un onglet manquant — 7 novembre 2026

Le CIL `JobItem` 247420 examine les onglets Excel par nom et ne compare les cellules que si les deux objets correspondants sont des feuilles de calcul. Common 46828 accepte aussi les onglets graphiques parmi les onglets manquants. `CopyExtraSheets` 253616 copie avant le prochain onglet commun, après le précédent ou en fin de classeur. Les nouveaux onglets servent de repères aux suivants. `UndoCopySheet` 215696 supprime l’onglet physique copié. [Worksheet.Copy](https://learn.microsoft.com/en-us/office/vba/api/Excel.Worksheet.Copy) · [Chart.Copy](https://learn.microsoft.com/en-us/office/vba/api/excel.chart.copy)

Arixcel fait un remplacement littéral du qualificateur source sur chaque feuille. Clair réécrit seulement les références de cellules résolues après l’arrivée du lot. Les chaînes littérales, les noms définis, les références 3D et les liens encore externes conservent leur texte. Les noms et liens au niveau du classeur qui résultent de la copie native restent gérés par Excel. Le reçu d’annulation suit l’identité physique, même après renommage. Un homonyme de remplacement est écarté.

Les bases des fonds conditionnels sont appliquées à la source pendant la copie puis le masque source est rétabli. Le journal CF peut être chargé lors de la copie si le classeur vient d’être rouvert. Le reçu est enregistré dès la création de l’onglet afin de couvrir aussi une erreur lors de la restauration du contexte. Le volet annule d’abord les cellules d’un lot mixte, puis les onglets.

Le témoin Excel masqué couvre la copie dans les deux sens, la feuille masquée, l’onglet graphique, le groupe de feuilles actif, le lien vers un onglet copié plus tard, le maintien d’un lien externe et le fond conditionnel après réouverture. Il vérifie aussi l’annulation après renommage, le refus sur un homonyme de remplacement et la reprise après protection de la structure.

## Copie directionnelle — 7 novembre 2026

Le CIL de `CompareForm.CopyData` 229328 limite la copie au mode formules. `KeyDownHandler` 229608 traduit les côtés affichés pour Ctrl+L et Ctrl+R. `Model.CopyData` 253168 trie la sélection par position et écarte les descendants déjà couverts. 255092 et 374268 copient le rectangle entier du bloc, dont les cellules vides, en appliquant le format numérique et le contenu A1 ou R1C1. Common 90940 et 91544 traitent les préfixes et les matrices CSE. 277840 restaure les propriétés stockées lors de l’annulation. Les formats supplémentaires et les feuilles manquantes suivent des chemins distincts. [Guide Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair applique le contenu typé et les formats numériques. Le bouton copie depuis le côté actif et les raccourcis ciblent le côté affiché. Le recalcul garde les blocs comparateurs initiaux même si la sélection Excel change. Si une insertion déplace un bloc comparateur, Clair demande une nouvelle comparaison. Une destination répétée est écartée. Une annulation protège les modifications ultérieures du contenu ou des formats numériques et rétablit les copies déjà annulées si un conflit interrompt le lot de copie. Après une annulation mixte partielle, le volet garde les opérations restantes.

## Alignement groupé — 7 novembre 2026

Le CIL de `Model.Align` 252772 développe la sélection aux descendants, déduplique les propositions et les trie par position dans l’arbre. La boucle 372908 garde les réussites malgré une erreur COM ou métier et crée une seule entrée d’annulation. `ApplyDelta` 214412 décale les propositions du même axe dans la tâche modifiée. Lorsque plusieurs tâches partagent une feuille de référence, une insertion de ce côté invalide le modèle. Le callback 241452 traite cette invalidation à la fin du lot.

Clair traite les propositions cochées ou affichées dans l’ordre du volet. Il décale les positions physiques après chaque insertion et recalcule les propositions lorsque la référence est partagée. Une proposition déjà résolue est écartée. Une nouvelle proposition reste visible pour le prochain lot. L’annulation dans le volet couvre toutes les insertions réussies d’une commande. Excel refuse de configurer un nouveau style d’alignement si une feuille du classeur protège son contenu. Clair vérifie cette condition avant de créer le style.

Le témoin Excel masqué couvre le sous-ensemble coché, deux décalages successifs sur le même axe, le partage d’une référence, une feuille protégée, les insertions réussies après le refus et l’annulation du lot. Dans Excel x64, la commande native annule d’abord une modification ordinaire, puis les deux axes Clair en un seul geste. Les contrôles appellent la commande Excel via COM. La validation du clavier physique et d’Excel x86 reste ouverte.

## Appliquer et retirer un alignement — 7 novembre 2026

Le CIL 213848 insère les axes entiers dans la feuille du côté le plus court. 214836 utilise un style nommé dont les six propriétés d’inclusion sont désactivées. La hauteur ou la largeur est lue sur le côté excédentaire et reprise lors de l’insertion. 218136 et 277676 distinguent le retrait d’axes marqués et l’annulation des plages retenues. [Comparaison Arixcel](https://www.arixcel.com/explorer/functions/compare)

Clair applique la proposition sélectionnée après l’avoir recalculée sur les feuilles physiques retenues. Un repère masqué par axe complète le style `Clair Alignment`. Un nom couvrant plusieurs axes peut s’étendre quand un utilisateur insère un axe à l’intérieur. Le retrait exige une géométrie entière, le style attendu et une vacuité contrôlée par Excel sur tout l’axe. Le retrait conserve les axes occupés et ceux qui traversent le corps d’un tableau, d’un filtre ou d’un tableau croisé.

Le témoin Excel masqué vérifie l’insertion, les références ajustées par Excel, les dimensions, l’annulation et le retrait après sauvegarde, réouverture et renommage. La comparaison de hauteur tient compte de l’arrondi de l’affectation Excel. Une formule renvoyant du texte vide en XFD bloque la suppression de la ligne. Un axe non marqué portant le style Clair reste intact.

## Détection d’alignement — 7 novembre 2026

Le CIL de `RowColumnString` 219896, 220436 et 220992 utilise jusqu’à trois textes utiles dans les cinquante premières cellules perpendiculaires à l’axe. La lecture s’arrête aux zones exclues. Le CIL 221308 normalise les références locales dans les formules. Le Common 78972 et 79736 cherche une plus longue sous-séquence avec une distance maximale de 2 000. Le CIL 376920 traite les colonnes avant les lignes, déplace certains repères sur des blancs ou des doublons et peut annuler deux décalages opposés. Un écart ajouté après la dernière ancre non vide reste dans la comparaison ordinaire. [Comparaison Arixcel](https://www.arixcel.com/explorer/functions/compare)

Le comparateur Common et Clair donnent les mêmes correspondances et le même dépassement de seuil sur 20 000 paires synthétiques. Les témoins Clair vérifient les décalages initiaux et intermédiaires, les ajouts finaux, les doublons, les formules et les zones exclues. Dans Excel masqué, le volet présente les propositions avant les écarts et sélectionne la plage excédentaire. Le passage aux écarts utilise la même capture. Le témoin différentiel couvre la recherche des correspondances. La construction complète du modèle Arixcel reste à couvrir.

## Blocs d’écarts et navigation — 6 novembre 2026

Le CIL de `TaskItem` 378176 classe les cellules en différentes, égales ou neutres. Le scanner Common 75764 étend les rectangles en conservant leurs limites et peut traverser une zone neutre. La seconde passe de `BlockItem` 380076 subdivise chaque bloc selon la paire de formules R1C1 ou les valeurs des deux cellules d’ancrage avec tolérance. Les constantes variées d’un sous-écart affichent un marqueur. Chaque sous-rectangle compte pour un écart. Un bloc avec un seul sous-écart se replie sur cet écart. [Comparaison Arixcel](https://www.arixcel.com/explorer/functions/compare)

Le scanner livré a été exécuté seul dans un processus .NET Framework sans Excel. Sur 40 000 grilles synthétiques avec quatre politiques de cases absentes et d’extension, ses rectangles correspondent à ceux de Clair. Un témoin pur couvre aussi un périmètre de la taille d’une feuille avec deux lignes occupées. Dans Excel masqué, les témoins Clair vérifient les rectangles translatés, les signatures, la tolérance par rapport à l’ancre, la navigation physique et la restauration des fonds. Le témoin différentiel couvre le scanner. La construction du modèle complet Arixcel, les frappes physiques et les segments de formule restent à comparer.

## Comparaison de classeurs, de feuilles et de plages — 6 novembre 2026

Le CIL de `RangeJobSpec` 264776 étend un coin cible à la taille de la référence et découpe les sélections plus grandes en blocs complets, avec saut des lignes ou colonnes séparatrices vides. `TaskItem` 378176 compare les plages par décalage. `JobItem` 247420 et 249272 apparie les feuilles de classeur par nom ou prend les onglets cibles sélectionnés, puis compare aux mêmes coordonnées. `Options` 247276 et les comparateurs 378940 et 379328 établissent R1C1 par défaut, A1 au choix, l’ignorance des entrées sans formule et la tolérance absolue inclusive de 0,000001 sur les valeurs. [Comparaison Arixcel](https://www.arixcel.com/explorer/functions/compare)

Le témoin Clair utilise deux volets et deux classeurs dans Excel masqué. Il modifie la référence après l’avoir définie, compare une plage ancrée puis plusieurs blocs, vérifie la sélection physique des deux côtés, change la notation et le mode de comparaison, groupe deux onglets cibles et renomme la feuille de référence. Les feuilles présentes d’un seul côté sont listées séparément. L’alignement et les opérations d’écriture restent à couvrir. Clair utilise ici les valeurs typées de `Value2` et la formule invariante d’Excel. Les conversions de types propres à `Value`, la présentation locale d’Arixcel et les grandes plages au-delà de la limite de capture restent ouvertes.

## Annulation dans les sous-tracés — 6 novembre 2026

Le CIL d’Arixcel donne une pile à chaque tracé et la relie à celle du parent. L’ouverture interne suspend les événements de focus. Ctrl+Z prend la dernière action dans la première pile liée qui contient une entrée. À la fermeture d’un enfant, les piles liées sont vidées du plus récent au plus ancien, puis leur contenu forme une seule action dans le parent immédiat. Un retour à Excel transfère la chaîne dans l’annulation globale et détache ses liens. Le chemin relève de `ModelessForm` aux RVA 57684, 58327, 58340 et 59720, et de `UndoStack` aux RVA 67272 et 67604 dans l’assemblage commun.

Les témoins Clair dans Excel masqué exercent l’ouverture, la fermeture et l’annulation sur trois tracés liés. Ils vérifient la priorité de l’édition enfant, le repli sur le parent, le regroupement à la fermeture d’un enfant vide et l’unique publication du lot à la sortie du volet. Un nouveau lancement transmet aussi l’ancien lot. Quand Alt+R retire un tracé dont la feuille source a disparu, Clair garde son édition dans la chaîne si sa cible physique est toujours accessible. La garde de conflit et la compensation du lot restent actives. La référence Arixcel provient du CIL. Le focus physique et la délégation d’édition à une comparaison restent à vérifier.

## Annulation des éditions depuis Excel — 6 novembre 2026

Le CIL d’Arixcel transfère l’historique local lors de la perte de focus. L’application inscrit une seule action `OnUndo`. Son lot est exécuté du changement le plus récent au plus ancien. Une sélection sûre est tentée après l’écriture. Clair enregistre une commande XLL dédiée, sélectionne la cible la plus ancienne après succès et conserve ses contrôles de conflit. Si un élément du lot échoue, les écritures déjà annulées sont compensées. [Application.OnUndo](https://learn.microsoft.com/en-us/office/vba/api/excel.application.onundo)

Le témoin dans le processus Excel x64 masqué utilise le bouton natif via `ExecuteMso("Undo")`. Il empile d’abord un effacement natif par-dessus deux éditions Clair. Le premier Undo restaure l’effacement et le suivant restaure les deux formules, sélectionne la cible la plus ancienne et actualise le volet. Arixcel connecté garde son raccourci Ctrl+Retour arrière. Le test du volet couvre aussi la fermeture, une saisie intervenue avant la publication, le changement de classeur et la compensation d’un conflit. Le focus du volet est simulé. La frappe physique et le chemin Excel x86 restent à vérifier. [Commande Office](https://learn.microsoft.com/en-us/office/vba/api/office.commandbars.executemso)

## Langue et notation des formules — 6 novembre 2026

Le CIL d’Arixcel distingue la langue, activée par défaut, de la notation Excel relue au retour du focus. Le changement redessine les composants et leurs positions dans le même arbre. Le témoin Clair compare le texte du champ aux quatre membres Excel `Formula2*`, conserve le nœud courant et la sélection de la grille, et vérifie la préférence entre deux volets. La sélection de texte utilise les positions recalculées, tandis que l’évaluation utilise toujours le texte A1 invariant.

Le témoin Excel français couvre les références relatives `LC(-1)`, les fonctions, décimales et séparateurs matriciels locaux, et un changement de notation pendant un brouillon. Un changement des séparateurs locaux pendant le brouillon suspend son application. La conversion CSE utilise l’ancre du bloc et conserve la sélection Excel. Les tests sans Excel couvrent les chaînes échappées, les axes compressés, les retours à la ligne et les correspondances contradictoires. Certaines réécritures de références structurées ou de qualifiants 3D peuvent laisser le texte affiché sans surlignage. Les variantes anciennes de l’API et les autres installations linguistiques restent à valider. [Application.ConvertFormula](https://learn.microsoft.com/en-us/office/vba/api/excel.application.convertformula)

## Saisie directe, constantes et matrices CSE — 6 novembre 2026

Le CIL distingue le texte affiché avant l’édition et le premier brouillon après une mutation. Les témoins WinForms exercent le remplacement d’une sélection, une frappe, la suppression, Alt+Entrée et Ctrl+Z. Excel masqué confirme l’interprétation des nombres, booléens, erreurs et textes, la différence entre un blanc et un texte vide, et la restauration du préfixe apostrophe. Le contrôle de conflit distingue un nombre d’un texte identique à l’affichage.

Les témoins CSE partent d’une cellule intérieure ou d’une sélection du bloc. Ils vérifient la formule canonique, l’étendue, les formats mixtes, les constantes répétées, l’effacement, la recréation après annulation et le rollback après recapture impossible. Une modification externe de la géométrie ou du contenu bloque l’opération. Les feuilles protégées et les tables de simulation refusent l’édition. Clair utilise la limite vérifiée du setter `FormulaArray` sans pilotage clavier. Ce lot utilisait A1 anglais. Les formules CSE plus longues, la frappe physique et la composition IME restent à vérifier ou à prendre en charge. [Range.FormulaArray](https://learn.microsoft.com/en-us/office/vba/api/excel.range.formulaarray)

## Éditer une formule dans le volet — 6 novembre 2026

L’analyse CIL d’Arixcel suit `StartEditing`, `AcceptEditing` et le callback d’annulation. F2 reprend le curseur de la formule entière, y compris depuis une sous-expression. Entrée valide et reconstruit depuis la source du tracé. Un échec de reconstruction déclenche une restauration. Ctrl+Z porte d’abord sur la saisie, puis sur une écriture au niveau de l’explorateur.

Le témoin Clair utilise un volet WinForms et une instance Excel masquée. Il édite un antécédent situé sur une autre feuille, vérifie les commandes clavier, la sélection du texte, la racine, les branches et l’historique. La limite de capture provoque un échec après une écriture qui aurait généré un format de date. La formule et le format initial sont restaurés, avec le brouillon toujours ouvert. Un autre témoin vérifie qu’une cellule au format Texte reçoit une vraie formule et retrouve son contenu et son format après annulation. Il déplace aussi la cellule éditée par insertion de ligne et renomme sa feuille. Ctrl+Z retrouve la cellule physique, même si l’adresse précédente contient une formule identique. Ces témoins portent sur l’annulation locale dans le volet.

## Reprendre un tracé depuis Excel — 6 novembre 2026

Le modèle d'historique couvre la limite de 100 entrées, les dédoublonnages, la reprise du parcours après restauration d'une sélection et l'éviction des vues fermées. Les deux XLL compilent. Le témoin natif appelle le délégué géré puis vérifie son retrait. Dans Excel isolé, les témoins de volet utilisent une cible de focus simulée pour vérifier le choix du contrôle, l'activation de la fenêtre réelle, le rafraîchissement du tracé et le repli sur sa sélection de lancement après fermeture. Le contrat Arixcel vient du CIL. La frappe physique dans Excel reste à vérifier.

## Sortir d'un tracé Excel — 6 novembre 2026

Les témoins Excel couvrent Échap et Entrée au niveau racine et enfant. Ils vérifient la sélection courante, la restitution optionnelle des plages disjointes après insertion, le coin actif du parent et le groupe d'onglets. La fermeture racine retire les couleurs du tracé et, quand Échap rétablit la source, remasque l'ancienne feuille active. Si la feuille source est supprimée, la fermeture conserve la sélection courante, même si une feuille homonyme a été recréée. Le volet en arrière-plan laisse la sélection de l'autre fenêtre intacte. L'appel de masquage du volet est observé après la fermeture dans le témoin WinForms. Le contrat Arixcel provient du CIL.

## Actualiser un tracé Excel — 6 novembre 2026

Alt+R reconstruit depuis la cellule ou les plages de lancement. Le témoin Excel modifie la formule pendant que la sélection Excel pointe ailleurs, restaure deux nœuds et une branche ouverte, puis exerce le même raccourci depuis le texte de formule. Le nœud courant donne la première zone resélectionnée. Un raccourci reçu pendant qu'une deuxième fenêtre du même classeur est active laisse sa sélection intacte. Une insertion de ligne et un renommage déplacent le tracé et son historique. Une feuille supprimée puis recréée sous le même nom laisse l'ancien tracé visible et signale sa source indisponible. La comparaison Arixcel repose sur le CIL. Clair retrouve les nœuds par leur chemin, et les branches d'expression par type et contenu correspondant.

Le volet reçoit l'identité de sa fenêtre depuis l'hôte Office. Pour évaluer une formule du classeur actif, la sélection temporaire reste dans cette fenêtre. Après évaluation, Clair restaure la fenêtre, la feuille, la plage et la cellule active.

## Raccourcis de navigation — 6 novembre 2026

Témoins Excel sur Ctrl+flèches, zones disjointes, coins, bouclage, autre fenêtre, retour à l'union, expansion et relance du tracé. Les transitions géométriques tournent aussi sans Excel. Contrat Arixcel relevé par le CIL.

## Dévoilement temporaire — 6 novembre 2026

Le témoin Excel couvre les feuilles masquées et très masquées, l'option des axes, la fermeture avec sélection active, deux propriétaires, les feuilles groupées et une sauvegarde suivie d'une réouverture. La protection empêche le dévoilement et déclenche le retour au masquage initial. Le comportement Arixcel provient du CIL. [Range.Hidden](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden) · [Worksheet.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible)

## Sélection de sous-expressions — 6 novembre 2026

Le témoin Excel passe de l'arbre au texte et du curseur au composant. Il couvre deux portions disjointes, Ctrl additif dans le texte, Maj+Bas dans l'arbre et une sélection de caractères conservée. La source est rose et les références résolues sont bleues. L'union Excel reste sur la feuille du composant courant, tandis que les autres feuilles gardent leur surbrillance. Désactiver les couleurs rétablit les fonds et garde le soulignement de la formule.

Le pont Excel est aussi vérifié séparément sur les noms locaux, références structurées, unions, intersections, débordements, axes entiers, cibles INDEX/FILTER et enfants 3D. Les interactions Arixcel sont établies par le CIL. Le témoin Clair exerce les gestionnaires WinForms et les sélections Excel.

## Ouverture des classeurs liés et noms externes — 6 novembre 2026

Le témoin Excel ferme deux sources. Un curseur dans la référence montre le chemin complet et permet de le refuser. Une référence à un nom global propose l'autre source. Après ouverture, le volet retrouve la formule, son nœud sélectionné et le retour vers l'ancienne racine. Excel garde sa fenêtre et sa sélection. **Entrée** depuis le nom global sélectionne sa plage physique dans le classeur rouvert. Une modification de la formule entre la proposition et la confirmation invalide la demande.

Le parseur couvre les noms absolus de classeur et de feuille, les homonymes, les cycles entre classeurs et les chemins relatifs dans la définition d'un nom. Les déclencheurs, le refus mémorisé et le rafraîchissement d'Arixcel sont établis par le CIL. Le témoin local utilise des fichiers `.xlsx`. L'ouverture via URL, l'authentification et les fichiers protégés restent à comparer. [Workbooks.Open](https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open) · [AutomationSecurity](https://learn.microsoft.com/en-us/office/vba/api/excel.application.automationsecurity)

## Dépendances entre classeurs ouverts — 6 novembre 2026

Le témoin Excel utilise trois classeurs aux noms proches, chacun avec une feuille `Data`. Il vérifie des antécédents à la même adresse dans les trois classeurs, l'évaluation d'une formule externe dans sa propre feuille, un dépendant dans un classeur pair, des rectangles homonymes et leur surlignage, la navigation explicite et le retour dans l'arbre. Fermer un pair garde le surlignage de la source. La capture suivante signale le fichier fermé et conserve les résultats des classeurs ouverts.

Le parseur est testé séparément sur les références 3D vers un classeur ouvert, les chemins et feuilles avec apostrophes, les lignes et colonnes entières et les intersections. Le témoin Excel exerce le volet par ses gestionnaires. Le comportement d'Arixcel pour l'ouverture du fichier lié provient du CIL.

## Surbrillance de la trace dans Excel — 6 novembre 2026

Le volet reprend la palette Arixcel relevée par REA. Source rose `#FF99FF`, antécédents bleus `#66CCFF`, dépendants verts `#99FFCC`. Le test du volet vérifie une source automatique, deux plages antécédentes sur des feuilles différentes, la couleur de la source du tracé enfant, le retour à la couleur précédente et le retrait par la case.

Le test Excel des rectangles colore une colonne complète, superpose Carte et Flux, puis vérifie le retrait de chaque couche. Une cellule modifiée manuellement dans la colonne conserve son fond, tandis que sa voisine est restaurée. Une feuille protégée conserve la sélection de couleurs précédente après un remplacement refusé. Le fond conditionnel continue à déterminer le rendu, avec ses règles, sa priorité et `StopIfTrue`. Un rectangle uniforme de 1 048 576 cellules est enregistré en un bloc et restauré après réouverture. Les tests XML couvrent aussi une feuille complète et les chevauchements de blocs.

Le journal enregistre encore des couleurs scalaires après réouverture. Un fond direct sur une colonne auparavant vide peut étendre la zone utilisée d'Excel au-delà du retrait du fond. Les cibles fusionnées et les dégradés refusent une nouvelle coloration dans ce moteur. Les couleurs du nouvel arbre sont comparées à l'original par inspection statique, et les effets ci-dessus sont exécutés dans Clair avec Excel.

## Suivi de la sélection Excel et retrait des couleurs — 6 novembre 2026

Le chemin de sélection d'Arixcel appelle `Range.Select` indépendamment du réglage de surbrillance. Les tests Clair vérifient dans Excel l'activation de la feuille du résultat courant avec Ctrl+Maj+Bas, l'union disjointe sur cette feuille, le repli d'un en-tête sur la cellule active et l'isolement d'un autre classeur actif. Les feuilles masquées restent masquées. Les tests appellent les gestionnaires du volet et ne mesurent pas le focus dans la fenêtre Office.

Dans la session Excel ouverte, le retrait de la dernière couche de couleurs directes restaure les propriétés de fond et de police capturées à la première écriture. Les tests superposent deux propriétaires, retirent leurs couches et vérifient thème, nuances, motif et couleur résultante. Ils couvrent aussi une retouche manuelle du motif à RGB identique et une reprise après échec de restauration. La projection de sauvegarde utilise toujours les couleurs scalaires du journal existant.

## Sélection multiple dans l'arbre — 6 novembre 2026

L'inspection statique d'Arixcel 9.0.9623.19616 établit les gestes Ctrl et Maj, l'exclusion des en-têtes du prochain tracé et la racine à deux catégories. Dans Clair, le test sur un classeur Excel jetable sélectionne deux plages sur `North` et `South` dans l'ordre inverse des onglets, vérifie les dix blocs dépendants sur la feuille racine `Audit`, puis revient à la sélection précédente. Il vérifie aussi Ctrl+Maj+Bas, l'union disjointe sur une seule feuille, la sélection d'un en-tête, la restauration après changement de regroupement et le retour dans un arbre de cellules.

Les quatre comparaisons directes de la section suivante repassent sur le même classeur. Pour la nouvelle sélection multiple, la comparaison Arixcel repose sur le CIL. Le test Clair exécute les gestionnaires du volet et vérifie les effets dans Excel. La surbrillance automatique dans la grille demeure ouverte.

## Antécédents et dépendants de plusieurs cellules — 6 novembre 2026

Comparaison avec Arixcel 9.0.9623.19616 sur deux copies au même SHA-256 de `ArixcelMultiTrace.xlsx`. Clair passe par les commandes du volet connectées à Excel. Les quatre listes et les compteurs concordent.

| Direction et sélection | Résultat commun |
| --- | --- |
| Antécédents de `Audit!B2:B4` | Deux plages, `North!B2:B4` et `South!B2:B4`. `North!B3` est vide et appartient à la plage. |
| Antécédents de `Audit!H2;H4` | Deux plages, `North!B2:B5` et `South!B2:B4`. Les formules référencent des colonnes entières. |
| Antécédents de `Audit!J2:J3` | Une plage, `North!B5`. `Audit!J2` appartient déjà à la source. |
| Dépendants de `North!B2;B4` | Huit blocs sur `Audit`, dans l'ordre `B2`, `D2:D3`, `E2:E3`, `H2`, `B3`, `D4`, `D5`, `D6`. Le regroupement affiche `Audit` avec le même total de huit. |

Les **32 contrôles Excel** du tracé vérifient aussi les zones au-delà de la partie utilisée, la navigation au clavier, le tracé depuis un bloc, le retour à une sélection disjointe et la place de l'arbre dans le volet. Les **69 groupes sans Excel** passent. Le nouveau groupe vérifie l'union sans doublon, la géométrie d'une colonne entière, l'exclusion de la source et le regroupement selon les formules R1C1.

## Dépendances et flux 3D — 6 novembre 2026

Arixcel 9.0.9623.19616 et Clair ont utilisé deux copies au même SHA-256 de `Arixcel3DReferences.xlsx`. Les témoins ci-dessous comparent des liens physiques. `Fev` est masquée.

| Commande Arixcel | Résultat natif et comparaison Clair |
| --- | --- |
| Dépendants de `Mar!B2` | Neuf cellules sur `Audit` dans les deux outils. `B2`, `B3`, `B4`, `B6`, `B7`, `B8`, `B9`, `B10`, `B11`. `B10` retient la constante `100` dans son `MAX`. |
| Dépendants de `Mar!B3` | `Audit!B4` et `Audit!B11` dans les deux outils. |
| Antécédents de `Audit!B2:B4` | Arixcel affiche `Jan!B2:B3`, `Fev!B2:B3` et `Mar!B2:B3`. L'union des antécédents de chacune des trois cellules dans Clair donne les six mêmes cellules. |
| Flux du classeur, zone centrale `Audit` | Cinq blocs d'antécédents externes et un bloc de cellules alimentées, identiques. Les plages externes sont `Jan!B2:B3`, `Fev!B2:B3`, `Mar!B2:B3`, `'North 1'!D2` et `'North O''Brien'!D2`. Le bloc interne est `Audit!B2:B11`. |

Les **87 assertions de Flux** incluent une branche conditionnelle inactive, un nom défini, une cellule vide explicitement incluse dans le périmètre, un lien hors périmètre et des bornes non résolues. Les **96 contrôles Excel** des références 3D couvrent aussi les commandes Antécédents et Dépendants du volet. Les **68 groupes sans Excel** passent. Le graphe couvre le classeur capturé. Dans le volet, la racine reste une cellule à la fois. [Preuves du flux](FLOW-EVIDENCE.md)

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

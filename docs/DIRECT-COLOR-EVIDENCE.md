# Couleurs directes : référence observée et reconstruction en cours

Relevé du 15 octobre 2026. Cible : Arixcel Explorer **9.0.9623.19616**, installé depuis le même artefact que le MSI fourni. SHA-256 de l'assembly installé : `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`, identité statique REA `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9`.

Le nouveau moteur décrit ci-dessous est raccordé à l'interface sur la branche de développement de la [PR #1](https://github.com/Baedys/audit-excel/pull/1), sans nouvelle livraison binaire. La livraison publique 0.6 continue à utiliser les règles conditionnelles. Aucune parité complète n'est revendiquée. L'empreinte de l'assembly original installé a été revérifiée le 3 novembre 2026 et reste identique.

## Observation réelle de l'original

Une instance Excel `/x` détenue par le contrôle, un classeur jetable, deux feuilles dans l'ordre Z puis A. En Z : B2=10, B3=2, B4=`=B2*B3`, B5=`=B4+1`, B7=`=1+1` ; A!B2=`=Z!B4*3`. Aucun autre document manipulé. ArixcelExplorer.Connect a été lu à true ; aucune activation ou modification de licence.

Les choix déjà enregistrés lors de l'ouverture affichaient Scope Workbook, Focus Worksheet Z et vue Inflows–Outflows. Ce n'est pas une preuve des réglages usine. Les essais de changement de vue et de regroupement n'ont pas abouti ; aucune conclusion runtime sur Inputs–Outputs ou Group by Sheet.

L'écran observé est une fenêtre de résultats dédiée et non modale : arbre avec colonnes de focus, activation et compteur, sélecteur View, menu More, bouton OK. Le menu expose notamment Switch View (Ctrl+S), Group by Sheet (Alt+G), Refresh (Alt+R), Reselect (Alt+S), Expand/Collapse All (Alt+Z) et Close & Keep Colors (Alt+X). Présence observée, fonctionnement de ces raccourcis non testé. La composition actuelle de MainForm Clair ne reproduit pas encore cet écran.

Séquence observée dans la fenêtre native :

1. Cocher INSIDE–OUTFLOWS colore Z!B4 ; parent central en état mixte.
2. Calculation Flow > Remove Colors retire cette couleur et décoche les nœuds. Fenêtre toujours ouverte.
3. Réactiver, puis Keep Colors et confirmer : couleur conservée, nœuds décochés, fenêtre toujours ouverte.
4. Fermer l'analyse, en ouvrir une nouvelle, activer puis retirer une nouvelle couche : Z!B4 conserve la couleur devenue permanente.

Après la séquence, lecture COM de la seule fixture : Z!B4 possède `Interior.Color=15652797`, `Interior.ColorIndex=24`, `FormatConditions.Count=0`, formule `=B2*B3`, valeur 20. Les autres cellules Z!B2/B3/B5/B7 n'ont pas de fond et n'ont aucune règle CF. Cela établit le format direct dans ce cas ; le code couleur relevé ne constitue pas une preuve de palette usine.

Les dix captures et le journal complet restent locaux dans `work/arixcel-runtime-0.7.md`. Le classeur n'a pas été sauvegardé. Il a été fermé sans enregistrer ; empreinte du fichier inchangée. Quit a retourné, mais le processus sans fenêtre restait en arrière-plan : terminaison après nouvelle vérification PID, démarrage, exécutable et ligne de commande de cette instance. Aucun processus de l'utilisateur fermé.

**Portée :** une fenêtre et un classeur. Pas de test runtime original de superposition multi-fenêtres, sauvegarde/réouverture du classeur, formats complexes ou panne. Les observations suivantes sont statiques.

## Couches, géométrie et éditions externes

Les sources CIL ciblées et leurs limites sont conservées dans `work/flow-global-lifecycle-evidence.md`, `work/highlight-lifetime-evidence.md`, `work/highlight-persistence-evidence.md` et le relevé précédent `work/flow-permanence-evidence.md`.

- `Manager.RemoveAllHighlights`, RVA307000, parcourt les modèles de flux de la fin vers le début. Les modèles fermés avec conservation sont inclus ; les fenêtres ouvertes restent ouvertes. `Model.RemoveHighlights`, RVA305744, synchronise les cases dans finally. Ce chemin ne filtre pas le classeur actif.
- `ForgetLayer`, RVA322424, traite chaque propriété séparément. Une couche non transparente devient la base anonyme ; toutes les couches inférieures sont abandonnées, même d'un autre propriétaire. Les supérieures restent. Une cible transparente seule disparaît. Aucun setter de couleur dans cette opération.
- `FastRange` conserve un `Range` COM et un rectangle. `Sheet.CleanUp`, RVA323716, peut relire le rectangle actuel de ce Range. Les identités de feuilles et de classeurs proviennent d'IUnknown, pas de leurs noms.
- `GetValidTargets`, RVA325356, compare la couleur directe courante au sommet attendu. Les fragments modifiés manuellement sont exclus, leur ancien historique abandonné ; les fragments encore conformes restent restaurables. Une ancienne adresse A1 seule ne suffit donc pas.
- Feuille devenue invalide : historique retiré sans écrire de formats. Le relevé ne prouve pas une purge immédiate au moment de la fermeture.
- Les retraits forcés de l'original peuvent perdre l'historique après échec d'un setter. Une transaction globale fiable n'est pas établie par ce code ; Clair ne doit pas présenter cette propriété comme héritée de l'original.

Les bases originales représentent RGB, automatic (-4105) ou none (-4142). Ni liaison au thème, ni tint, ni motif complet dans ExcelColor. Les couches de fond et de police sont indépendantes. Détails supplémentaires et RVAs dans [FLOW-EVIDENCE.md](FLOW-EVIDENCE.md).

## Sauvegarde et réouverture : résultat statique seulement

BeforeSave projette l'historique dans une partie `Workbook.CustomXMLParts`. La structure originale conserve, par propriété et rectangle, **base et sommet**, sans les modèles, catégories, cases ni l'empilement complet. Le rectangle provient du Range actuel. Les autres namespaces ne sont pas ciblés par le code examiné.

À la réouverture, le code propose trois chemins :

- Oui : reconstruire l'historique, puis tenter de restaurer les couleurs originales, sans forcer les suppressions échouées.
- Annuler : reconstruire et conserver l'historique temporaire pour plus tard.
- Non, après seconde confirmation : abandonner l'historique sans modifier les couleurs visibles.

Les couches reconstruites utilisent un nouveau propriétaire conservé par AppEvents, distinct des modèles CalcFlow. Les commandes globales de flux ne le ciblent donc pas directement. Les anciennes métadonnées sont retirées en mémoire ; BeforeSave reconstruit ensuite celles encore nécessaires. Ce constat ne prouve pas l'écriture immédiate du fichier sur disque.

Une reconstruction Clair doit employer son propre namespace et séparer propriétaires restaurés et modèles d'analyse. Le comportement réel save/reopen de l'original reste à contrôler ; aucun journal complet ni recréation de fenêtres ne découle de ces données.

## Nouveau moteur Clair et raccordement local

`DirectColorLayers` contient des piles indépendantes par cellule opaque et propriété. Base anonyme, calques nommés et propriétaires séparés ; mise à jour en place, montée explicite, retrait, permanence et abandon d'une propriété modifiée extérieurement. Copies et vues publiques indépendantes.

`DirectColorTransaction` prépare une copie, valide et capture toutes les écritures, puis publie après réussite. Une écriture qui modifie Excel avant de lever est elle aussi restaurée. Échec de restauration : journal de réparation et nouvelles mutations bloquées. La disparition vérifiée d'une référence autorise un abandon ciblé ; reçu conservé dans `AbandonedRepairHistory`, avec avertissement distinct d'une restauration réussie.

`ExcelDirectColorCell` utilise les formats intrinsèques. Les instantanés de thème, nuances et motifs servent au retour arrière après échec et au retrait de la dernière couche dans la session ouverte. Le moteur compare ces propriétés au dernier format écrit ou adopté avant de préparer une action ou de projeter l'historique. Une retouche manuelle de même RGB devient la nouvelle base d'une activation explicite. Le journal de réouverture conserve les couleurs scalaires et les index spéciaux. Le moteur direct laisse les règles CF intactes. Le masque de carte décrit plus bas possède son propre journal de fonds. Fusion, dégradés et texte multicolore restent refusés avant mutation.

Contre-essai Clair du 3 novembre : après coloration uniforme d'une police en noir, modifier la couleur d'un seul caractère produit `Font.ColorIndex=DBNull` et `Font.Color=DBNull`. Le retrait levait alors « Couleurs hétérogènes non prises en charge. » et ne restaurait même plus le remplissage. La lecture de réconciliation distingue désormais ce résultat mixte d'une erreur COM : seul l'historique de la police est abandonné, sans setter sur ses caractères ; remplissage et autres cellules restent restaurables. Même exclusion avant adoption de métadonnées non chargées et avant projection Save d'un historique reporté. Le refus initial de recolorer une police déjà multicolore reste explicite. Cette correction s'appuie sur le principe statique d'exclusion des éditions externes ; aucun essai différentiel de texte enrichi sur Arixcel n'est revendiqué.

`ExcelDirectHighlights` partage une session entre les bridges du même Excel. Opérations incrémentales par cible ; une modification de palette utilise ExistingOnly et ne réactive pas une propriété abandonnée après édition manuelle. Les bindings du modèle gardent les Range, même après retrait ou permanence, tant que le modèle reste ouvert. `BindDirectHighlightTargets` doit être appelé sur toute l'appartenance d'un modèle fraîchement capturé avant de rendre la main à Excel, y compris pour les cellules jamais activées. Une nouvelle capture remplace ces bindings. Cellules supprimées et classeurs fermés ne sont jamais retrouvés au nom seul.

Le travail local raccorde désormais les flux et cartes à cette session. Chaque fenêtre de flux possède son propriétaire. Les cartes de l'interface actuelle partagent désormais leur modèle de feuille physique et leur palette dans la session Excel : une seconde vue réutilise le même propriétaire et sa fermeture ne détache pas la carte. Les modèles dont tous les bindings sont vérifiés invalides sont purgés avant les opérations ; une erreur COM ou une protection ne constitue pas cette preuve. Retrait et permanence globaux synchronisent les cases sans fermer les fenêtres. Les cellules désactivées restent liées pour suivre une insertion ; l'actualisation remappe les choix survivants et refuse une racine supprimée. Les gestes explicites réélèvent uniquement leurs propres membres ; une palette ne réactive pas une propriété modifiée manuellement.

Les snapshots Excel retiennent les identités IUnknown du classeur et de ses feuilles. Un classeur homonyme rouvert ne remplace donc pas une cible fermée ; renommage et SaveAs suivent l'objet vivant. Les anciennes règles CF de Clair disposent de commandes de retrait séparées.

La persistance utilise le namespace propre `https://clair-excel.org/schemas/direct-colors/v1`, sans accès réseau. XML strict : DTD, commentaires, instructions de traitement, contenu inconnu et recouvrements refusés ; limites de 64 parties, 128 Mi caractères cumulés et 500 000 propriétés. Le document projette base/sommet et rectangles actuels, sans anciens modèles ni fenêtres. L'adoption ne modifie aucun format ; Oui restaure, Annuler conserve un propriétaire restauré distinct, Non confirmé abandonne l'historique. Refuser la confirmation conserve l'historique temporaire. Une commande permet sa restauration ultérieure dans la même session.

Avant une nouvelle coloration, les métadonnées encore inconnues sont adoptées pour préserver la vraie base. Avant Save, l'historique non chargé est conservé ; une projection déjà connue ne se réimporte pas. Le remplacement des seules parties propres à Clair possède un chemin de récupération après mutation suivie d'erreur. Les propriétaires restaurés restent exclus des commandes globales flux/carte.

Un watcher partagé entre complément et fenêtres d'un même processus gère Open, Activate et BeforeSave. Un essai réel a révélé que ComEventsHelper livrait BeforeSave hors du thread d'origine en autonome, annulant l'enregistrement. Il a été remplacé par un sink AppEvents sous StandardOleMarshalObject et IConnectionPoint. Le test événementiel suivant valide 21 contrôles réels : Save, réouverture, Oui/Non/Annuler, restauration différée, refus de permanence et partage/libération des abonnements. Les décisions sont injectées par le test ; la présentation native des boîtes n'est pas une preuve différentielle contre Arixcel.

La fermeture autonome projette les métadonnées en mémoire avant de libérer le dernier watcher ; aucun Save disque n'est imposé. Après fermeture du dernier processus Clair, les insertions ultérieures dans Excel ne sont plus suivies. Pour conserver ce suivi, le complément doit rester chargé. L'application autonome et le complément chargés simultanément ont des registres distincts : leur coopération n'est pas validée.

Une preuve statique supplémentaire (`work/map-direct-cf-cil.txt`) montre que la carte originale masque temporairement les fonds des règles CF de la feuille, puis les restaure : StoredFormat ctor RVA348996, Dispose RVA349164, Map.StoreConditionalFormats RVA114984. Le checkpoint antérieur `448b2c5` laissait ces règles actives ; le raccordement CF suivant traite ce décalage. Remplissages complexes, identité visuelle, performances et autres cas structurels restent dans [PARITY.md](PARITY.md).

## Carte originale : contrôle réel du 3 novembre

Une copie jetable contient Z!B4=`=B2*B3`, base directe 1193046, règle CF verte 65280 et police bleue 16711680 ; Z!H10=20 porte une règle CF jaune 65535 et police rouge 255. Les règles ont priorités 1/2 et StopIfTrue vrai/faux. Arixcel est connecté ; Clair n'est pas enregistré dans cette instance. Le ruban original est commandé par Computer Use et les relevés COM vérifient PID, démarrage, exécutable et chemin exact du classeur détenu. Les fichiers de preuve restent dans `work/native-startup-d7a1a4c8d6a548ccb56aa9a9d7e71667/`.

- Bascule sur une cellule ordinaire : B4 devient 49407, H10 devient 7471103. Les deux règles renvoient `Interior.ColorIndex=DBNull` ; leurs formules, priorités, StopIfTrue et couleurs de police restent identiques. Nouvelle bascule : base directe de B4 et fonds CF vert/jaune restaurés.
- Carte limitée à la colonne entière B : B4 est colorée ; H10 conserve son absence de fond direct, mais son fond CF est lui aussi masqué. La portée de ce masquage dépasse donc la bande cartographiée.
- Cartes sur B puis H : les deux couleurs directes coexistent. Retirer B restaure pourtant les fonds CF des deux règles ; la carte H demeure directe à 7471103, mais son affichage redevient jaune 65535. Un compteur partagé conservant le masquage jusqu'au dernier modèle **ne reproduirait pas** ce cas original.
- Sauvegarde avec une seule carte B active : dans le XLSX, les deux `dxf` CF ne contiennent plus de `fill`, alors que leurs polices restent présentes. Une partie `workbookData` propre à Arixcel apparaît. Après fermeture puis réouverture du classeur dans la même instance Excel, le dialogue propose de restaurer les couleurs de trois blocs. Choisir Oui rétablit B4 à 1193046, mais les fonds CF restent `DBNull` pour B4 et H10. La restauration des fonds CF après cette sauvegarde échoue donc dans l'original observé. Aucun redémarrage du processus n'a été testé par cette séquence.

La propriété des modèles est aussi distincte de celle des flux : le store original est statique au contexte Excel et la bascule retire toutes les cartes intersectant les zones demandées, sans créer les autres zones de cette même demande. Plusieurs bandes disjointes d'une feuille peuvent porter des cartes. Le détail statique et ses limites sont conservés dans `work/map-target-ownership-contract-20261103.md`.

## Raccordement de la bascule native par zones

Le callback `ExcelAddIn.ShowMap` appelle désormais la bascule de sélection, sans créer de fenêtre d'exploration. L'application autonome propose la même action dans « Toutes les cartes… ». Les actions explicites de coloration et retrait par feuilles restent disponibles dans cette vue.

`GetMapSelectionAreas` réplique les rectangles sur les feuilles sélectionnées de la fenêtre Excel. Les feuilles protégées sont exclues ; si aucune feuille de calcul admissible ne reste, une erreur explicite est renvoyée. Sélection ordinaire : zone utilisée de la feuille. Ligne ou colonne entière : intersection de la bande et de la zone utilisée. Les rectangles exactement identiques sont dédupliqués, sans fusion des bandes distinctes.

Le relevé CIL de Common précise le calcul des bornes : jusqu'à **1 000 000 cellules**, la zone utilisée conserve aussi les blancs formatés ; au-delà, `ShrinkToValues` garde son origine et sonde des bandes de 100 colonnes aux bords gauche/droit, puis de 1 000 lignes aux bords haut/bas. Il utilise Value2 et End vers le haut/gauche ; il ne calcule pas une enveloppe exacte par Find. Une intersection vide devient une cellule à l'origine de la bande demandée. Détails dans `work/map-smart-used-range-contract-20261103.md`.

La commande prépare d'abord un descripteur d'identité et les zones, puis utilise `FormulaMapTargets` pour décider globalement retrait ou création. Le retrait ne capture aucun contenu de cellules. La création ne lit que les rectangles ciblés, en conservant les noms définis et les identités des feuilles : un UsedRange énorme sur une autre feuille ne bloque plus une petite bande. Le plafond de capture reste de 250 000 cellules réellement lues ; les grandes zones denses restent une limite ouverte.

Chaque modèle conserve un `ExcelMapRange` physique avec tickets IUnknown du classeur et de la feuille. Insertions et renommages suivent la référence ; fermeture/réouverture homonyme ne la remplace pas. Les catégories transparentes et les blancs appartiennent aussi au modèle. Leur activation ultérieure par palette suit donc les références retenues. La classification d'une bande ne compare pas les formules à un voisin situé hors de sa borne haute ou gauche.

Le déchargement du complément projette désormais l'historique avant de libérer le watcher, même si aucune fenêtre Clair n'a été ouverte. Une erreur de projection est propagée ; les abonnements sont libérés dans finally. Ce raccordement ne constitue pas à lui seul une nouvelle validation du chargement COM natif ni de l'identité visuelle du ruban.

L'utilisateur a choisi une **restauration complète des fonds CF après sauvegarde/réouverture**. Le défaut original observé ci-dessus ne fait donc pas partie du comportement demandé.

## Raccordement local des fonds conditionnels

Ce raccordement appartient à la branche de développement ; les archives 0.6 antérieures ne sont pas mises à jour. Après les couleurs directes, chaque carte acquiert les fonds CF dérivables de toute sa feuille. Le retrait d'un propriétaire restaure ses fonds avant ses couches directes, même si une autre bande reste active. La permanence conserve seulement les couleurs directes : les fonds CF d'origine sont rétablis. Les polices, bordures, conditions, priorités, StopIfTrue et AppliesTo ne reçoivent aucun setter de ce raccordement.

`ExcelConditionalFill` capture neuf getters : ColorIndex, Color, ThemeColor, TintAndShade, Pattern, PatternColorIndex, PatternColor, PatternThemeColor et PatternTintAndShade. Un scalaire, DBNull et une propriété indisponible sont distincts. L'essai réel montre des sentinelles de thème 0, -1 et -4142 ; elles ne sont pas toutes des valeurs de setter valides. La restitution écrit les propriétés définissant le fond qui diffèrent, puis exige une relecture exacte. Réassigner systématiquement tous les getters changeait notamment un PatternColor nul en zéro : cette approche n'est pas utilisée.

Le journal propre `https://clair-excel.org/schemas/conditional-fills/v1` conserve base complète, sommet attendu et intention Pending. XML strict, DTD et contenu inconnu refusés, bornes de 10 000 règles et 16 Mi caractères. Le journal de récupération est préparé avant le premier setter de masque. Un échec qui mutile puis lève déclenche une restitution ; une réparation incomplète conserve les bases et bloque les nouvelles mutations et la sauvegarde ordinaire.

Les projections directes et CF sont validées et ajoutées avant retrait des anciennes parties. Une compensation vérifie les contenus réels, même si Office recrée les IDs ou si Add/Delete a modifié avant de lever. Une compensation incomplète possède un reçu immuable et une API de reprise ; seuls les contenus issus des snapshots et projections de ce reçu sont admissibles, sans suppression d'un contenu étranger modifié entre-temps. L'erreur ne prétend pas une atomicité que la compensation n'a pas démontrée.

BeforeSave coordonne les deux journaux. Open/Activate détecte aussi un journal **CF seul**, sans aucune couleur directe. Oui restaure les deux familles ; Annuler adopte sans setter et conserve la restauration ultérieure ; conserver les couleurs directes ne signifie jamais oublier les fonds CF. La fermeture et le déchargement projettent aussi les propriétaires CF suivis.

L'identité COM d'une règle n'est pas utilisable comme identité logique : plusieurs accès à la même règle ont renvoyé des IUnknown distincts dans la sonde détenue. Après une modification de collection, lire un ancien wrapper précédait également un arrêt du serveur Excel. Le suivi ne conserve donc aucun wrapper CF entre opérations : il conserve des témoins gérés et des Range AppliesTo physiques, puis utilise des wrappers fraîchement énumérés. Un index de collection sert seulement de raccourci validé par le prédicat relatif et la géométrie courante. La collection servant au raccourci est libérée avant les setters, comme dans le chemin d’énumération complet. Les modes `$` des cellules, lignes et colonnes entières font partie des témoins. Une correspondance absente ou ambiguë conserve la base ; elle ne prouve pas la suppression de la règle.

Un setter CF sur une feuille non protégée peut échouer sur une autre feuille protégée du même groupe Excel. Le port isole donc temporairement la feuille physique cible lorsque sa fenêtre est groupée. Il restitue d'abord la sélection, puis le groupe et la feuille active, vérifie les objets et les rectangles, et rétablit les flags ; les erreurs du setter et du contexte sont conservées ensemble. Une fenêtre déjà seule ne provoque aucune activation. Le groupe protégé possède des assertions de fenêtre, feuilles physiques, sélection, cellule active et flags.

Une sonde détenue a aussi comparé 21 paires de fenêtres avant/après Activate : même HWND, IUnknown différents, avec une seule fenêtre de classeur. Les comparaisons de fenêtre utilisent donc le handle natif de Window, tout en retenant les vrais objets COM et l'appartenance à Workbook.Windows. Les feuilles et classeurs conservent leurs identités physiques. Le contrat de [Window.hWnd est documenté par Microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.window.hwnd).

Les bordures CF utilisent les côtés xlLeft/xlRight/xlTop/xlBottom, distincts des indices xlEdge de Range.Borders. La fixture et les témoins emploient les côtés CF après un refus 1004 observé avec xlEdgeBottom. Ce contrat est documenté par [Microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.formatcondition.borders).

Limites à vérifier ou compléter : références absolues hors de l'ancre AppliesTo lors d'insertions, règles remplacées ou dupliquées indiscernables, changements externes des témoins, familles CF sans Interior exposé, gradients avec points de couleur, performances et redémarrage complet de l'hôte. Un changement externe de fond détectable n'est pas requalifié en sommet temporaire par une sauvegarde ; il reste un conflit conservant la base. Cette prudence ne constitue pas une preuve de parité de tous les scénarios d'édition CF avec l'original.

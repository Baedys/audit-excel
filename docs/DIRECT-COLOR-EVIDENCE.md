# Couleurs directes : référence observée et reconstruction en cours

Relevé du 15 octobre 2026. Cible : Arixcel Explorer **9.0.9623.19616**, installé depuis le même artefact que le MSI fourni. SHA-256 de l'assembly installé : `8684a9eb5eee58d3ab38628da5aba4cc62749b244f967391bf37c357bf3d43a6`, identité statique REA `ev_8791460ac50faf2c5956c3480900dbd89abb4b8c110bde762c61fa9ca5b0eed9`.

Le nouveau moteur décrit ci-dessous est **local, non livré et non raccordé à l'interface**. La livraison publique 0.6 continue à utiliser les règles conditionnelles. Aucune parité complète n'est revendiquée.

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

`ExcelDirectColorCell` utilise les formats intrinsèques. Les snapshots riches de thème, nuances et motifs servent au **rollback d'échec** ; le retrait normal du moteur logique restitue une couleur résolue ou un index spécial, pas une liaison au thème. CF utilisateur inchangée et susceptible de masquer le format direct. Fusion, dégradés et texte multicolore restent refusés avant mutation.

`ExcelDirectHighlights` partage une session entre les bridges du même Excel. Opérations incrémentales par cible ; une modification de palette utilise ExistingOnly et ne réactive pas une propriété abandonnée après édition manuelle. Les bindings du modèle gardent les Range, même après retrait ou permanence, tant que le modèle reste ouvert. `BindDirectHighlightTargets` doit être appelé sur toute l'appartenance d'un modèle fraîchement capturé avant de rendre la main à Excel, y compris pour les cellules jamais activées. Une nouvelle capture remplace ces bindings. Cellules supprimées et classeurs fermés ne sont jamais retrouvés au nom seul.

Le travail local raccorde désormais les flux et cartes à cette session. Chaque fenêtre de flux possède son propriétaire ; les cartes gardent un modèle par feuille/fenêtre. Retrait et permanence globaux synchronisent les cases sans fermer les fenêtres. Les cellules désactivées restent liées pour suivre une insertion ; l'actualisation remappe les choix survivants et refuse une racine supprimée. Les gestes explicites réélèvent uniquement leurs propres membres ; une palette ne réactive pas une propriété modifiée manuellement.

Les snapshots Excel retiennent les identités IUnknown du classeur et de ses feuilles. Un classeur homonyme rouvert ne remplace donc pas une cible fermée ; renommage et SaveAs suivent l'objet vivant. Les anciennes règles CF de Clair disposent de commandes de retrait séparées.

La persistance utilise le namespace propre `https://clair-excel.org/schemas/direct-colors/v1`, sans accès réseau. XML strict : DTD, commentaires, instructions de traitement, contenu inconnu et recouvrements refusés ; limites de 64 parties, 128 Mi caractères cumulés et 500 000 propriétés. Le document projette base/sommet et rectangles actuels, sans anciens modèles ni fenêtres. L'adoption ne modifie aucun format ; Oui restaure, Annuler conserve un propriétaire restauré distinct, Non confirmé abandonne l'historique. Refuser la confirmation conserve l'historique temporaire. Une commande permet sa restauration ultérieure dans la même session.

Avant une nouvelle coloration, les métadonnées encore inconnues sont adoptées pour préserver la vraie base. Avant Save, l'historique non chargé est conservé ; une projection déjà connue ne se réimporte pas. Le remplacement des seules parties propres à Clair possède un chemin de récupération après mutation suivie d'erreur. Les propriétaires restaurés restent exclus des commandes globales flux/carte.

Un watcher partagé entre complément et fenêtres d'un même processus gère Open, Activate et BeforeSave. Un essai réel a révélé que ComEventsHelper livrait BeforeSave hors du thread d'origine en autonome, annulant l'enregistrement. Il a été remplacé par un sink AppEvents sous StandardOleMarshalObject et IConnectionPoint. Le test événementiel suivant valide 21 contrôles réels : Save, réouverture, Oui/Non/Annuler, restauration différée, refus de permanence et partage/libération des abonnements. Les décisions sont injectées par le test ; la présentation native des boîtes n'est pas une preuve différentielle contre Arixcel.

La fermeture autonome projette les métadonnées en mémoire avant de libérer le dernier watcher ; aucun Save disque n'est imposé. Après fermeture du dernier processus Clair, les insertions ultérieures dans Excel ne sont plus suivies. Pour conserver ce suivi, le complément doit rester chargé. L'application autonome et le complément chargés simultanément ont des registres distincts : leur coopération n'est pas validée.

Une preuve statique supplémentaire (`work/map-direct-cf-cil.txt`) montre que la carte originale masque temporairement les fonds des règles CF de la feuille, puis les restaure : StoredFormat ctor RVA348996, Dispose RVA349164, Map.StoreConditionalFormats RVA114984. Clair préserve ces règles sans les masquer ; elles peuvent donc cacher les couleurs directes. Cette différence reste ouverte. Remplissages complexes, identité visuelle, performances et autres cas structurels restent dans [PARITY.md](PARITY.md).

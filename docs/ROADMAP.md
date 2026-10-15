# Suite de la reconstruction

1. Étendre le chargement COM désormais vérifié localement à Excel 32/64 bits et aux autres versions ; terminer groupes, raccourcis et options du ruban.
2. Remplacer la lecture lexicale par un analyseur syntaxique couvrant les opérateurs de références, noms relatifs, tables et expressions LET/LAMBDA.
3. Étendre l'arbre syntaxique et l'évaluation Excel ajoutés en 0.2 : branches actives, cibles de recherches, paramètres LAMBDA, longues formules et valeurs de tous les nœuds ; vérifier la fidélité contre la référence.
4. Résoudre les références externes vers les classeurs ouverts et les plages déversées, en indiquant les frontières non disponibles.
5. Comparer plages, feuilles et classeurs avec alignement explicite des lignes/colonnes et options de formats.
6. Étendre carte et flux sur grille 0.4 : options d'analyse, regroupement en blocs, commande rendant les couleurs permanentes, progression et comparaison runtime contre Arixcel.
7. Indexer les dépendants, rendre les longues analyses annulables et tester de grands modèles.
8. Tester installation/désinstallation, Excel 32/64 bits et versions prises en charge ; produire un paquet signé.

La version 0.2 répare le chargement COM, ajoute l'arbre syntaxique, l'évaluation contextualisée, l'édition avec annulation et la classification scope/focus des flux. La version 0.3 ajoute la carte R1C1, les couleurs réversibles sur grille et une palette persistante. La version 0.4 ajoute les couleurs de flux, leur palette, l'activation par groupe/cellule et la fermeture avec retrait ou conservation. Les exigences restantes et les preuves de chaque étape figurent dans [PARITY.md](PARITY.md). L'équivalence complète reste ouverte.

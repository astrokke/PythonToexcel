PythonXlsx - Automatisation de Planning

Ce script Python automatise la transformation de fichiers Excel contenant des données de planification de formation.

Fonctionnement:

Le script lit les fichiers Excel (.xlsx) présents dans le dossier a-faire, les traite et génère des fichiers de planning formatés dans le dossier fait. 

Traitement des données:

Identification de la colonne de date: Le script recherche automatiquement la colonne contenant les dates des séances.
Tri par date: Les données sont triées par ordre chronologique.

Fusion de lignes consécutives: Les lignes consécutives ayant la même matière, le même formateur et la même modalité sont fusionnées. Le nombre de jours est calculé en conséquence.

Calcul des jours facturables: Les jours d'autoformation ne sont pas considérés comme facturables.

Ajout de totaux: Le nombre total de jours et d'heures  est calculé, pour les jours facturables et non facturables.

Formatage du fichier de sortie:

Le fichier de sortie est un fichier Excel (.xlsx) nommé planning YP - [Nom de la session ] - du [date].xlsx

Le planning est formaté avec un en-tête spécifique, des couleurs pour différencier les types de séances et un logo (si le fichier logo_diginamic.png est présent dans le même répertoire).

Les marges, l'orientation (paysage) et la mise en page sont optimisées pour l'impression.

Utilisation:

Préparer les fichiers Excel:
Placer les fichiers Excel à traiter dans le dossier a-faire.
S'assurer que les fichiers Excel ont une colonne contenant les dates des séances.
Exécuter le script:
Ouvrir un terminal et naviguer jusqu'au répertoire contenant le script PythonXlsx.py
Installer les bibliothèques nécessaires 
Exécuter la commande python3 PythonXlsx.py.
Récupérer les fichiers de planning:
Les fichiers de planning formatés seront générés dans le dossier fait.
Remarques
Le script nécessite Python 3 et la bibliothèque pandas. Installer les dépendances 
Le formatage du fichier de sortie est spécifique à un certain type de planning. Des modifications du code source peuvent être nécessaires pour l'adapter à d'autres types de planning.
Le script suppose que la première ligne du fichier Excel contient les en-têtes de colonnes.

# python-replace-docx

## Description
python-replace-docx est un script Python conçu pour effectuer des remplacements de texte dans des fichiers DOCX. Le script prend en charge la conservation de la mise en page des documents d'origine.

## Prérequis
Avant d'utiliser le script, assurez-vous d'avoir installé les éléments suivants :
- Python : Téléchargez et installez Python à partir du site officiel : [https://www.python.org](https://www.python.org/ftp/python/3.11.3/python-3.11.3-amd64.exe)
- openpyxl : Vous pouvez installer cette bibliothèque Python en exécutant la commande suivante dans votre terminal : `pip install openpyxl`
- docx : Vous pouvez installer cette bibliothèque Python en exécutant la commande suivante dans votre terminal : `pip install docx`

## Utilisation
1. Téléchargez le fichier `convert_mep.py` si vous souhaitez conserver la mise en page d'origine, ou téléchargez `convert_no_mep.py` si vous ne souhaitez pas conserver la mise en page.
2. Créez un dossier sur votre ordinateur et placez le fichier `.py` téléchargé à l'intérieur.
3. Dans le même répertoire, créez un dossier `input_files` et placez tous les fichiers DOCX que vous souhaitez convertir à l'intérieur.
4. Créez un dossier `output` qui contiendra les fichiers convertis.
5. Placez le fichier Excel (.xlsx) à la racine, dans le même répertoire que le script `convert*.py`.
6. Dans la deuxième colonne du fichier Excel, entrez les champs que vous souhaitez remplacer. Il est préférable de mettre un champ par ligne, mais si une ligne a plusieurs possibilités de remplacement, vous pouvez les séparer par des points-virgules (`;`).
7. Les autres colonnes du fichier Excel doivent contenir les valeurs de remplacement correspondantes pour chaque champ.
8. Exécutez le script à partir de votre terminal en utilisant la commande suivante : `python convert_mep.py` ou `python convert_no_mep.py`, en fonction du fichier que vous avez téléchargé.
9. Les fichiers convertis seront générés dans le dossier `output`.


Assurez-vous d'avoir les permissions nécessaires pour lire les fichiers d'entrée et écrire les fichiers de sortie.
Le script ne pourra pas fonctionner si l'un des fichiers dans le dossier "output" est ouvert par word, veillez à fermer les fichiers avant de lancer le script.

**Note :** Veillez à sauvegarder une copie de sauvegarde de vos fichiers d'origine avant de les convertir, au cas où vous auriez besoin de revenir en arrière ou de les restaurer ultérieurement.

## Remarque
Ce script est fourni tel quel, sans garantie d'aucune sorte. Veuillez l'utiliser à vos propres risques.
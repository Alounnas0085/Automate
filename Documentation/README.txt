AUTOMATE CEGID – EXPORT DES PRESTATIONS

Description :
Ce script automatise la récupération des prestations depuis le site Cegid
et génère un fichier Excel.

Fonctionnement :

1. Double-cliquer sur lancer_cegid.bat
2. Le script ouvre le navigateur
3. Se connecte automatiquement
4. Extrait les données
5. Génère un fichier Excel dans le dossier "Extractions"

Structure :

* Lancement : scripts + configuration
* Extractions : fichiers Excel générés

Configuration :
Le fichier config.json contient :

* email
* mot de passe

Au lancement :

* possibilité de modifier ou conserver les identifiants

Prérequis :

* Python installé
* Playwright installé

Commandes d’installation :
pip install playwright openpyxl
python -m playwright install chromium

Auteur : Ali LOUNNAS

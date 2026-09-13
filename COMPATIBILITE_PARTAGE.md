# Vérification du 13 septembre 2026

Horizon Chantier doit fonctionner sur macOS et Windows, avec le code et les chantiers sur un partage. Les modifications concernent le code source ; les applications déjà construites dans `dist` ne sont pas mises à jour.

## Corrections vérifiées localement

- Le dossier de données peut être placé dans ou à côté de MonProjetPython. Le bouton « Choisir le dossier des chantiers » permet de choisir son emplacement sur chaque poste, y compris un lecteur réseau ou un partage UNC Windows. Le choix est enregistré dans le profil du poste, pas dans les fichiers communs. Un partage choisi puis inaccessible ne provoque pas de basculement silencieux sur un autre chantier.
- Les configurations du bordereau sont lues et écrites explicitement en UTF-8 ; les noms accentués sont conservés. Elles ne contiennent que des noms de fichiers relatifs, sans chemin Mac ou Windows.
- La copie, son original, `bordereau_public.json`, `Avenants.xlsx` et `Revision_global.xlsx` doivent rester ensemble dans le dossier du chantier ; le PR reste dans `data/prix_de_revient.xlsx`. Les liens Excel ajoutés sont relatifs.
- Le pilotage lit maintenant les montants situés après les libellés fusionnés P:S et U:V. Les précédentes corrections de présentation avaient rendu cette lecture incorrecte.
- La branche Excel Windows distingue une formule d'une constante lors de l'écriture dans les zones autorisées du PR.
- Les lanceurs Mac et Windows utilisent le dossier du code ; le lanceur Windows accepte un dossier UNC grâce à `pushd`. Ils utilisent le Python du poste, sans réutiliser un environnement virtuel créé sur un autre système.
- Les scripts de construction ciblent désormais `main.py` avec `main.spec`, et non l'ancienne application `app_devis.py`. Le bundle Mac n'est construit que sur macOS. Aucun ancien exécutable n'a été remplacé.

## Résultats et limites

36 tests exécutés sur macOS : création, import, calcul, clôture, export des seuls prix, conservation des formules, synthèse, déplacement des dossiers et accents. Les branches d'ouverture Windows et d'écriture Excel Windows ont été simulées avec des doubles de test. Cela ne constitue pas une exécution réelle sous Windows.

Le workflow `.github/workflows/compatibilite.yml` est préparé pour exécuter les mêmes tests sur macOS et Windows si ce dépôt est utilisé dans GitHub. Il n'a pas été exécuté à distance dans cette session. `Verifier_Horizon_Windows.cmd` permet de lancer les tests sur un PC sans GitHub.

Restent à valider sur les postes réels : ouverture dans Excel Windows/Mac, actualisation des liaisons et caches de formules, export PDF Excel, autorisations et verrouillage du partage SMB, et construction/lancement des exécutables. Aucun de ces points ne doit être présenté comme validé par les seuls tests Python.

Le générateur de copie publique reconnaît pour l'instant la structure de Mariemont (articles B, quantité K, prix L). La préparation automatique d'autres structures de bordereau n'est pas validée : elle exige une reconnaissance explicite de leurs colonnes. Le parcours de création existant utilise encore les modèles ; ces tests ne prouvent pas une migration automatique de tous les nouveaux chantiers.

## Installation du code partagé

Python 3.10 ou supérieur, avec Tkinter, et Excel de bureau sur chaque poste pour l'automatisation. Installer les dépendances sur chaque poste avec `python -m pip install -r requirements.txt` (ou `py -3 -m pip ...` sous Windows). Lancer `Lancer_Horizon_Windows.cmd` ou `Lancer_Horizon_Mac.command`. Pour construire une application, installer également PyInstaller et lancer le script de construction sur le système cible.

Référence consultée pour l'interface Excel : [documentation officielle xlwings Range](https://docs.xlwings.org/en/stable/api/range.html). Les vérifications de ce rapport reposent sur les tests du projet ; cette documentation ne vaut pas validation du logiciel.

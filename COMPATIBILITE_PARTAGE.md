# Vérification du 13 septembre 2026

## Sauvegarde PDF Historique sans classeur reconnu — 11 octobre 2026

La capture de l'utilisateur montre le blocage « Aucun fichier Excel ouvert n'a été trouvé ». Ce cas a été reproduit avec l'ancienne méthode : sans classeur actif ni référence mémorisée valide, elle s'arrêtait sans permettre une sauvegarde. La cause précise de l'absence de classeur reconnu sur le poste de l'utilisateur n'est pas établie.

Dans `main.py`, le bouton conserve la priorité au fichier Excel actif, puis au dernier fichier mémorisé encore présent. Si aucun fichier n'est retrouvé, le sélecteur de fichiers s'ouvre dans le chantier sélectionné. Le classeur choisi est exporté dans son `Historique_Sauvegardes` avec le nom horodaté existant. Annuler ne crée rien. Un fichier choisi hors du chantier est refusé. Aucun modèle, calcul métier ni fichier de chantier n'est modifié par cette correction technique ; elle couvre les chantiers existants et futurs et les dossiers déplacés.

Validation : 16 tests locaux réussis sur ce parcours et la compatibilité du partage, incluant les chemins accentués, la référence devenue absente après déplacement, la conservation des sauvegardes successives et les erreurs d'export. Les comportements Excel Mac/Windows ouvert ou fermé sont simulés dans les tests. Un export réel dans Excel Mac a ensuite créé un PDF lisible depuis un classeur d'essai fermé du projet, dont l'empreinte est restée identique ; la détection et la sélection de fichier sont pilotées dans cet essai. Résultat : `SORTIES/pdf_historique_20261011/verification_excel.json`. Aucun essai réel Excel Windows n'est revendiqué.

Relancer Horizon avec le `main.py` actualisé. Les postes Mac/Windows utilisant une copie locale doivent recevoir ce fichier avec les modules existants ; les exécutables doivent être reconstruits sur leur système cible. Les autres postes et exécutables n'ont pas été actualisés ici ; les fichiers de chantier actuels hors projet n'ont pas été consultés. Sauvegarde préalable : `SAUVEGARDES/Avant_correction_pdf_historique_20261011_155923.zip`.

## Total général des prix de vente dans la synthèse PR — 11 octobre 2026

`synthese_pr.py` affiche sur la première page le « Total général des prix de vente HT (options comprises) », égal à la vente de base plus la vente des options. Les deux récapitulatifs restent séparés. Une vente sans résultat enregistré rend ce total non disponible ; un vrai zéro reste affiché. Aucun classeur ni modèle n'est modifié : les synthèses des chantiers existants et futurs prennent cette ligne lors de leur prochaine génération.

Validation : 12 tests locaux réussis, dont six cas de total général (avec/sans options, zéros, résultat absent). PDF générés et texte vérifié depuis les copies des cinq chantiers et du modèle présentes dans `SORTIES/zeros_pr_20261008/pret`, sources inchangées. Contrôle visuel de la première page de Mariemont : base 414 465,97 € + options 84 848,40 € = total 499 314,37 €. Les copies 001, 002 et du modèle comportent des formules sans résultat enregistré : leur total reste explicitement non disponible. Ces copies ne prouvent pas les montants actuels du partage. Résultats conservés dans `SORTIES/total_general_synthese_pr_20261011`.

Pour les postes Mac et Windows utilisant ce code, relancer Horizon puis régénérer la synthèse. Les copies locales doivent recevoir `synthese_pr.py` ; les exécutables doivent être reconstruits sur leur système cible. Les PR actuels du partage hors projet, les postes distants et les exécutables n'ont pas été consultés ni mis à jour ici. Aucun chemin propre au poste ni dépendance supplémentaire n'est introduit.

## Original Mariemont indice C effectivement installé — 11 octobre 2026

À la demande explicite de l'utilisateur, le fichier du Bureau `2420_MARIEMONT_METRE RECAP_LOT 1- dernière version MAC_261005_IND_C.xlsx` a été copié à l'identique dans le chantier réel `004_Mariemont`. Sa feuille est `LOT 1 - IND. C`. La référence `original` de `bordereau_public.json` et son empreinte ont été actualisées : le bouton « Postes / Métré » cible désormais cet indice C. L'indice `250924` indiqué dans le compte rendu précédent est dépassé pour cet accès.

La configuration précédente est sauvegardée dans `Sauvegarde/Metre_original_avant_indice_C_20261011`. L'original du Bureau, le PR, l'avancement et les anciens fichiers sont conservés. La copie de l'original et sa source ont la même empreinte SHA-256 : `3d63dc3828f0be3fb1494e549217e81a7983f7e0524a7b87f2866397413f8337`. Préparation, installation et vérification : `SORTIES/metre_original_indice_c_20261011`.

Limite constatée : le nouveau métré utilise les articles en A, quantités en J et prix unitaires en K. L'avancement existant conserve l'ancien format. L'ancien export automatique des prix refuse ce nouveau classeur avant toute écriture ; son adaptation n'a pas été réalisée dans cette intervention de raccordement à l'original. Cette installation ne vaut donc pas validation de l'export ni de la soumission complète.

Le chemin du document dans la configuration reste relatif, utilisable sur Mac et Windows et après déplacement du chantier. La cible du bouton a été vérifiée avec les fichiers réels puis avec une copie du chantier déplacée ; l'ouverture par le système est simulée. Relancer Horizon avec le `main.py` actualisé pour bénéficier du bouton. Les postes qui consultent ce même partage utilisent la nouvelle référence ; les copies locales et exécutables des autres postes n'ont pas été actualisés ici. Les autres chantiers conservent leurs propres originaux.

## Accès au métré original — 11 octobre 2026

Pour le 004 Mariemont, l'utilisateur soumissionne uniquement le lot 1. Le bouton existant « Postes / Métré » ouvre désormais le fichier `original` de `bordereau_public.json`. La configuration de préparation disponible dans le projet référence `2420_MARIEMONT_METRE RECAP_LOT 1-MAC_250924.xlsx`. Le bouton ne parcourt pas les autres lots, les anciennes versions ou les copies d'avancement. Il conserve l'accès à `Metre_detaille.pdf` pour les chantiers sans bordereau configuré ; un original configuré mais absent déclenche un message sans substitution.

La correction de `main.py` utilise la référence relative propre à chaque chantier existant ou futur, sans imposer le lot 1 aux autres marchés et sans modifier modèles, calculs ou fichiers de soumission. Neuf tests locaux réussis sur l'ouverture et le bordereau public, avec plusieurs lots/versions présents et un futur chantier déplacé. L'ouverture par le système et les dialogues sont simulés. Les fichiers réels du 004 restent hors du dossier de travail autorisé : leur configuration actuelle et l'absence d'une publication plus récente ne sont pas vérifiées ici. Aucun lot n'a été déplacé ou supprimé.

Relancer les postes utilisant le code partagé ; pour les copies locales, actualiser `main.py` avec les modules existants. Les exécutables et autres postes n'ont pas été mis à jour ici.

## Copies de secours administratives — 11 octobre 2026

À l'ouverture et à l'actualisation du dossier Administratif, les fichiers nommés « Copie de secours… » et les sauvegardes Word `.wbk` sont rangés dans `Administratif/Copies_de_secours`. Leurs sous-dossiers sont conservés. Les versions différentes reçoivent un suffixe sans écrasement ; le fichier source n'est retiré qu'après vérification de la copie. Les pièces courantes restent inchangées. Le dossier de secours reste accessible par le bouton existant « Dossier (ajout manuel) », mais ses fichiers sont exclus de la liste courante et du PDF regroupé pour l'envoi. L'export exclut aussi les copies encore dispersées, sans déplacer lui-même les sources.

La règle commune s'applique aux chantiers existants et futurs à l'ouverture de leur dossier administratif ; aucun modèle ni calcul n'est modifié. Les fichiers réels du chantier montré, hors du dossier de travail autorisé, n'ont pas été déplacés ici. Pour les postes utilisant le code partagé, relancer Horizon ; les copies locales doivent recevoir ensemble `main.py`, `bibliotheque_administrative.py` et `dossier_administratif_pdf.py`. Les exécutables et postes distants n'ont pas été actualisés ici.

Validation : 33 tests locaux réussis sur les bibliothèques, les imports, le rangement sans perte et le PDF. Essai réel de la fenêtre Tk Mac avec documents temporaires reprenant les noms de la capture : liste nettoyée, originaux inchangés, nouvelle sauvegarde détectée automatiquement et deux versions conservées. Les tests portent aussi sur les conflits de noms, un échec de copie et un futur chantier déplacé.

## Offres : bibliothèque générale puis import choisi — 11 octobre 2026

Toutes les offres passent par la bibliothèque générale. À son ouverture, les offres reçues et les anciennes copies des chantiers sont reprises dans la générale, sans effacer les documents des chantiers. Les accès « Anciennes offres à classer » et « Classer les offres reçues » sont retirés de l'accueil : la reprise accompagne déjà l'ouverture de la bibliothèque.

« Ajouter des offres » ajoute les nouveaux fichiers à la bibliothèque générale. Les fenêtres d'offres des chantiers proposent l'ajout depuis cette bibliothèque ; l'import externe direct et le dépôt dans ces fenêtres sont retirés. La catégorie destinataire se choisit dans la fenêtre de bibliothèque, puis « Importer dans le chantier » copie le fichier ou dossier sélectionné. Aucun choix de catégorie dans une seconde fenêtre n'est nécessaire. Les factures conservent leur parcours existant.

Correction dans `main.py`, commune aux chantiers existants et futurs, sans changement de modèle ni de calcul. La reprise des données réelles se fera à l'ouverture de la bibliothèque avec ce code. Les fichiers réels du 004 et des autres chantiers, situés hors du dossier de travail autorisé, n'ont pas été consultés ni modifiés ici. Les documents déjà dans les chantiers sont conservés : leur présence ne permet pas de distinguer un import choisi d'une ancienne copie automatique.

Validation : 22 tests locaux sur la reprise, l'import, les bibliothèques et les factures ; essai réel de la fenêtre Tk Mac avec fichiers temporaires, vérifiant l'ajout uniquement dans la générale, l'import explicite dans le chantier, la conservation des originaux et la visibilité des boutons. Script : `SORTIES/bibliotheque_offres_20261011/verifier_parcours.py`. Les dialogues de sélection des fichiers sont pilotés pendant cet essai. Aucun essai réel Windows n'est revendiqué.

Les postes utilisant le code partagé prennent la correction au redémarrage. Les copies locales doivent recevoir `main.py` avec les modules existants ; les exécutables doivent être reconstruits sur leur système cible. Aucun poste distant ni exécutable n'a été actualisé ici. Aucun chemin propre à un poste n'est ajouté à la logique commune.

## Zéros dans les PR et leurs modèles — 8 octobre 2026

`zeros_pr.py` complète réellement les cases numériques vides de la feuille Chiffrage par zéro. Les cellules de texte, unités, titres, styles et données déjà saisies restent conservées. Pour les divisions simples concernées, un diviseur vide ou nul donne zéro ; le calcul existant est conservé lorsque le diviseur est renseigné. Les autres erreurs ne sont pas converties en zéro. Cette adaptation des calculs concerne uniquement le PR, pas le pilotage ni les états d'avancement.

`main.py` applique cette préparation à l'ouverture du PR et à la création d'un chantier. La reprise a été appliquée aux cinq PR 001 à 005 et aux deux fichiers modèles : 52 496 cases complétées par zéro. `SORTIES/zeros_pr_20261008/installation.json` atteste son application ; les sept originaux sont sauvegardés dans `SAUVEGARDES/PR_zeros_20261008.zip`. Les classeurs ouverts ou modifiés depuis la préparation sont protégés contre le remplacement. Aucun document annexe, image ou lien n'est réenregistré.

Validation : 23 tests locaux réussis sur les zéros, la création et la synthèse, ainsi que 65 tests sur les cahiers, pourcentages et consommateurs planning. Les contrôles des fichiers réels, de la création depuis le modèle et de la synthèse 005 sont consignés dans ce dossier de reprise. Les essais de recalcul Excel Mac sur les copies 001 à 003 restent distincts des contrôles de fichiers réalisés sur les autres classeurs.

Pour les postes utilisant le code partagé, relancer Horizon. Les copies locales doivent recevoir `main.py`, `zeros_pr.py`, `synthese_pr.py` et leurs modules existants ; les exécutables doivent être reconstruits sur leur système cible. Les postes distants et exécutables ne sont pas mis à jour par cette intervention. Aucun chemin propre au poste n'a été introduit dans la logique commune.

## Synthèse d'un PR partiellement renseigné — 8 octobre 2026

`synthese_pr.py` compte les saisies numériques laissées vides à zéro, y compris les heures et jours. Les moyennes sans montant ou heure à pondérer valent zéro. L'ancien calcul de jours `Iligne/Cligne/8`, enregistré en division par zéro avec zéro heure et un effectif vide ou nul, contribue également pour zéro. Une formule sans résultat enregistré ou une autre erreur conserve son alerte. Les options restent séparées des totaux de base.

La lecture commune couvre les chantiers existants et futurs sans modifier leurs classeurs ou modèles. Pour chaque poste Mac/Windows utilisant une copie locale, actualiser `synthese_pr.py` puis relancer Horizon ; les exécutables exigent une nouvelle construction. Aucun poste distant ni exécutable n'a été mis à jour ici. Validation : 31 tests réussis sur la synthèse et ses consommateurs planning ; 4 tests VBA ignorés faute d'interpréteur de développement. Le contrôle sur le PR réel 005 et le PDF généré est conservé dans `SORTIES/005_Chapelle_Perwez/simulation_synthese_20261008`.

## Import des dossiers d'offres — 8 octobre 2026

Dans la bibliothèque générale des offres, « Importer dans le chantier » importe désormais le dossier fournisseur sélectionné. Le double-clic continue à ouvrir les dossiers et à importer les fichiers. La copie et le choix de catégorie restent inchangés.

Correction dans `main.py`, commune aux chantiers existants et futurs, sans reprise de données ni modification de modèle. Les postes utilisant ce code la prennent au redémarrage ; les copies locales doivent recevoir ce fichier et les exécutables doivent être reconstruits sur leur système cible. Aucun poste distant ni exécutable n'a été mis à jour ici.

Validation : 25 tests locaux réussis, dont sélection des dossiers Gobert/Wienberg et navigation avec interface simulée ; copie de deux dossiers vérifiée avec documents temporaires. Le test reproduit l'échec du bouton avant correction. Les offres réelles et le chantier 04 ne sont pas dans le dossier de travail autorisé : leur import effectif reste à vérifier.

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

## Dossier administratif par Gmail — 30 septembre 2026

Dans le dossier Administratif, « Envoyer le dossier par mail » regroupe les pièces et sous-dossiers en un PDF à la racine du chantier. Les originaux et les exports précédents sont conservés. Le PDF sélectionné dans Finder/Explorateur est à glisser dans le message Gmail ouvert avec son objet et son texte ; le destinataire et l'envoi restent à la main de l'utilisateur.

PDF et images JPG/PNG/TIFF/BMP sont regroupés directement. Word et Excel utilisent les applications de bureau déjà employées par Horizon. Une pièce illisible, protégée ou de format non pris en charge bloque la préparation en indiquant le fichier, sans publier de dossier incomplet.

Le code commun couvre les chantiers existants et futurs, y compris après déplacement ; aucune modification des modèles ou des données des chantiers n'est nécessaire. Les postes utilisant le code partagé prennent la fonction au redémarrage. Pour les copies locales, mettre à jour ensemble `main.py`, `dossier_administratif_pdf.py` et `mail_soumission.py`, avec les dépendances du projet. Les exécutables déjà construits doivent être reconstruits sur leur système cible ; ils n'ont pas été mis à jour ici.

Validation locale : 24 tests ciblés réussis et essai réel de la fenêtre Tk sur Mac avec deux PDF temporaires. L'ouverture Gmail, Word/Excel et les commandes Windows sont simulés dans les tests ; aucun mail n'a été envoyé. Les postes distants n'ont pas été accessibles pendant cette intervention.

## Options dans le pilotage — 30 septembre 2026

Une option simplement chiffrée, sans quantité portée dans l'état, apparaît à zéro. Dès qu'elle est portée dans l'état, le pilotage affiche son montant et ses réalisés du mois et cumulé. Son montant complète le marché ; seules ses quantités exécutées alimentent le réalisé. La soumission initiale est conservée. Après clôture, les quantités précédentes/cumulées conservent la prise en compte de l'option malgré le retour du mois à zéro.

La lecture commune utilise les lignes d'options des formats d'état déjà pris en charge. Elle complète les synthèses calculées uniquement lorsque leurs formules excluent ces lignes, afin d'éviter les doubles comptes. Les synthèses renseignées manuellement restent les montants de référence. Cette lecture ne modifie aucun classeur ni modèle ; les chantiers existants et futurs en bénéficient à l'ouverture du pilotage avec le code actualisé.

Validation : 61 tests locaux réussis (options, pilotage, bordereau, clôture et création), PDF réellement généré et contrôlé visuellement sur une page. Pour une copie locale du programme, actualiser ensemble `main.py` et `options_pilotage.py`. Les exécutables et postes distants n'ont pas été mis à jour pendant cette intervention.

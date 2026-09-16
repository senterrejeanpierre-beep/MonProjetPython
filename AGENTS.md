# Contrainte permanente du projet
Horizon Chantier est partagé et utilisé sur Windows et macOS. Toute correction doit considérer les deux systèmes et les déplacements des dossiers de chantier sur le partage.
- Avant de modifier un classeur modèle partagé ou une règle de calcul métier, décrire la modification et le fichier touché. Distinguer explicitement ces changements des corrections techniques de fiabilité. Ne pas déduire une nouvelle règle métier d'une simple alerte de contrôle du PR.
- Ne pas introduire de chemin propre à un Mac ou un PC dans la logique commune ou les fichiers partagés.
- Conserver les préférences absolues de montage du partage sur chaque poste ; conserver les références aux documents du chantier relatives.
- Vérifier la création, le PR, l'avancement, la clôture, le pilotage et l'export lorsque leurs contrats changent.
- Distinguer tests locaux, branches simulées et essais réels Excel/Windows. Ne jamais annoncer une validation Windows à partir de tests exécutés seulement sur Mac.

# Règles métier confirmées le 13 septembre 2026
- Le pilotage actuel est figé : conserver sa présentation, ses indicateurs et ses calculs.
- Le pilotage doit lire les montants de la synthèse en bas de l'état d'avancement, que ces montants soient alimentés manuellement ou automatiquement. Exception confirmée le 16 septembre 2026 : les avenants en moins sont lus dans le fichier Avenants du chantier pour le seul pilotage ; ils ne figurent jamais dans la synthèse de l'état d'avancement et ne réduisent pas le réalisé. Ne pas substituer les autres montants lus dans les fichiers Avenants ou Révisions.
- Soumission HT, réalisé cumulé et réalisé du mois doivent être présents et lisibles. Un vrai zéro est valide, notamment en l'absence d'avenants ou de révisions ; une cellule manquante ou une formule sans résultat n'est pas un zéro.
- La partie droite de l'état d'avancement et sa synthèse sont figées, sauf correction d'une erreur de calcul confirmée.
- La clôture conserve la soumission et le cumulé, reporte les quantités cumulées dans les précédentes et remet les quantités du mois à zéro, sans ajouter deux fois le mois au cumulé.
- Le fichier Avenants ajuste le montant du marché dans le pilotage ; il ne prouve aucune exécution. Seuls les postes réellement exécutés dans l'état d'avancement alimentent le réalisé du mois et le réalisé cumulé. Ne pas ajouter automatiquement un montant d'avenant du fichier annexe aux totaux exécutés ou facturables de l'état.
- Pour un autre format de bordereau public, adapter uniquement la correspondance des colonnes du document source (articles, quantités, prix, montants, options). Réutiliser la partie droite et le pilotage ; ne pas les reconstruire à chaque marché.
- La reconnaissance générale de nouveaux formats n'est pas encore implémentée. Le générateur actuel cible Mariemont.
- Écart constaté le 13 septembre puis corrigé le 16 septembre : le pilotage ne doit plus substituer aux montants de la synthèse les avenants en plus ou les révisions des fichiers annexes. La seule lecture annexe conservée est celle des avenants en moins pour le pilotage.

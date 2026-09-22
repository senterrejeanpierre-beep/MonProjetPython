# Interdiction permanente : indépendance des logiciels
- Règle absolue : intervenir uniquement sur ce que l'utilisateur demande explicitement. Ne rien ajouter d'autre, ne prendre aucune initiative fonctionnelle ou visuelle et ne modifier ni les polices ni la présentation sans demande explicite. Horizon Chantier est un logiciel métier et doit conserver ce caractère.
- Le dossier du projet courant est une frontière stricte pour toute intervention. Ne jamais consulter, parcourir ou modifier les dossiers des autres logiciels de l'utilisateur.
- Ne créer aucune dépendance entre ses logiciels et ne toucher à aucun fichier, environnement Python, dépendance ou réglage global partagé susceptible d'affecter un autre logiciel.
- Cette interdiction est déjà décidée et reste applicable entre les sessions, sans nouvelle confirmation. La généralisation à tous les chantiers et postes concerne uniquement Horizon Chantier ; elle n'autorise jamais une intervention dans un autre logiciel.
- Les données et préférences propres à Horizon prévues par son fonctionnement ne constituent pas une autorisation de sortir du dossier de travail pendant une intervention. Respecter le périmètre explicitement autorisé par l'utilisateur.

# Contrainte permanente du projet
Tout travail sur Horizon Chantier doit couvrir tous les chantiers existants et futurs et tous les postes Windows et macOS. Cette portée est la règle par défaut, sans nouvelle confirmation de l'utilisateur ; seule une demande explicite de sa part peut limiter une intervention à un chantier ou un poste.
- Un chantier cité ou montré en capture (par exemple Mariemont) est un cas d'exemple, pas une limitation du périmètre.
- Pour chaque changement, traiter le code commun, la reprise des chantiers existants concernés et la création des futurs chantiers (modèles partagés compris lorsque nécessaire). Préserver les données et particularités métier de chaque chantier ; généraliser le fonctionnement ne signifie pas recopier les données d'un chantier dans les autres.
- Prévoir l'utilisation et la mise à jour sur tous les postes, avec leurs emplacements de partage propres. Ne pas considérer une correction appliquée à un seul classeur ou au seul poste de développement comme un déploiement général terminé.
- Avant d'annoncer la fin, vérifier cette portée générale et signaler précisément tout chantier ou poste restant bloqué ou inaccessible, sans prétendre l'avoir mis à jour. Ne pas attendre que l'utilisateur demande à nouveau la généralisation.
Horizon Chantier est partagé et utilisé sur Windows et macOS. Toute correction doit considérer les deux systèmes et les déplacements des dossiers de chantier sur le partage.
- Cette compatibilité est une exigence permanente, déjà confirmée par l'utilisateur : ne pas lui demander de la répéter ni de la reconfirmer à chaque correction. Concevoir et relire chaque changement pour Windows et macOS, y compris l'ouverture des fichiers, les chemins et les fenêtres de dialogue.
- Avant de modifier un classeur modèle partagé ou une règle de calcul métier, décrire la modification et le fichier touché. Distinguer explicitement ces changements des corrections techniques de fiabilité. Ne pas déduire une nouvelle règle métier d'une simple alerte de contrôle du PR.
- Ne pas introduire de chemin propre à un Mac ou un PC dans la logique commune ou les fichiers partagés.
- Conserver les préférences absolues de montage du partage sur chaque poste ; conserver les références aux documents du chantier relatives.
- Vérifier la création, le PR, l'avancement, la clôture, le pilotage et l'export lorsque leurs contrats changent.
- Distinguer tests locaux, branches simulées et essais réels Excel/Windows. Ne jamais annoncer une validation Windows à partir de tests exécutés seulement sur Mac.
- Dans les comptes rendus, mentionner une limite de validation seulement si elle est utile à la décision de l'utilisateur ; ne pas répéter systématiquement que Windows n'a pas été testé. Ne jamais présenter une vérification locale comme un essai réel sur l'autre système.

# Règles métier confirmées le 13 septembre 2026
- Numérotation des articles : suivre la logique métier, avec une suite continue lors de l'ajout de postes. Les lignes vides, titres et en-têtes Excel ne consomment pas de numéro d'article. Les dossiers et leurs liens reprennent le numéro métier ; préserver les documents lors d'une renumérotation.
- Le pilotage actuel est figé : conserver sa présentation, ses indicateurs et ses calculs.
- Le pilotage doit lire les montants de la synthèse en bas de l'état d'avancement, que ces montants soient alimentés manuellement ou automatiquement. Exception confirmée le 16 septembre 2026 : les avenants en moins sont lus dans le fichier Avenants du chantier pour le seul pilotage ; ils ne figurent jamais dans la synthèse de l'état d'avancement et ne réduisent pas le réalisé. Ne pas substituer les autres montants lus dans les fichiers Avenants ou Révisions.
- Soumission HT, réalisé cumulé et réalisé du mois doivent être présents et lisibles. Un vrai zéro est valide, notamment en l'absence d'avenants ou de révisions ; une cellule manquante ou une formule sans résultat n'est pas un zéro.
- La partie droite de l'état d'avancement et sa synthèse sont figées, sauf correction d'une erreur de calcul confirmée.
- La clôture conserve la soumission et le cumulé, reporte les quantités cumulées dans les précédentes et remet les quantités du mois à zéro, sans ajouter deux fois le mois au cumulé.
- Le fichier Avenants ajuste le montant du marché dans le pilotage ; il ne prouve aucune exécution. Seuls les postes réellement exécutés dans l'état d'avancement alimentent le réalisé du mois et le réalisé cumulé. Ne pas ajouter automatiquement un montant d'avenant du fichier annexe aux totaux exécutés ou facturables de l'état.
- Pour un autre format de bordereau public, adapter uniquement la correspondance des colonnes du document source (articles, quantités, prix, montants, options). Réutiliser la partie droite et le pilotage ; ne pas les reconstruire à chaque marché.
- La reconnaissance générale de nouveaux formats n'est pas encore implémentée. Le générateur actuel cible Mariemont.
- Écart constaté le 13 septembre puis corrigé le 16 septembre : le pilotage ne doit plus substituer aux montants de la synthèse les avenants en plus ou les révisions des fichiers annexes. La seule lecture annexe conservée est celle des avenants en moins pour le pilotage.
